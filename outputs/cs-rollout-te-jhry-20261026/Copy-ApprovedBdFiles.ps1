[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$NodePath,
    [Parameter(Mandatory=$true)][string]$ParserPath,
    [switch]$Apply,
    [switch]$WorkerOnly,
    [switch]$ChangedInputFilesOnly,
    [ValidatePattern('^[0-9]{8}-[0-9]{6}$')]
    [string]$RunTag = (Get-Date -Format 'yyyyMMdd-HHmmss')
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Set-Location -LiteralPath $repo
$runRelative = 'outputs/cs-rollout-te-jhry-20261026/applied-' + $RunTag
$runRoot = Join-Path $repo $runRelative
if (Test-Path -LiteralPath $runRoot) { throw 'Run directory already exists.' }
[void][IO.Directory]::CreateDirectory($runRoot)
$extractor = 'outputs/te-availability-role-fix/extract-powerquery.ps1'
$validator = 'outputs/allocation-diagnosis/validate-m.cjs'
$files = @(
    '1. Input/1-AllocationExtracted.xlsx',
    '1. Input/2-DemandExtract.xlsx',
    '2. Calculations/Capacity-ShiftAvailability.xlsx',
    '2. Calculations/Capacity.xlsx',
    '2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx',
    '2. Calculations/CS/CapacityDistrib(A.2)-shifts.xlsx',
    '2. Calculations/CS/CapacityDistrib(B)-shifts.xlsx'
)
if ($WorkerOnly) { $files = @('1. Input/Worker Reconciliation.xlsx') }
if ($ChangedInputFilesOnly) {
    if ($WorkerOnly) { throw 'Select only one restricted file group.' }
    $files = @('1. Input/1-AllocationExtracted.xlsx','1. Input/2-DemandExtract.xlsx','2. Calculations/Capacity-ShiftAvailability.xlsx')
}
function ScopedPath([string]$unit, [string]$relative) {
    if ($unit -notin @('BD','TE','JH-RY') -or $relative -notin $files) { throw 'Unapproved copy target.' }
    $unitRoot = [IO.Path]::GetFullPath((Join-Path $repo ('CLIENT/DATExx-Whiddon/UNITS/' + $unit)))
    $path = [IO.Path]::GetFullPath((Join-Path $unitRoot $relative))
    if (-not $path.StartsWith($unitRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Resolved path escaped its approved unit.'
    }
    return $path
}
function Hash([string]$path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
function CheckClosed([string]$path) {
    $owner = Join-Path (Split-Path -Parent $path) ('~$' + (Split-Path -Leaf $path))
    if (Test-Path -LiteralPath $owner) { throw ('Excel owner file present: ' + $path) }
    $handle = [IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $handle.Dispose()
}
function MCore([string]$path) {
    $raw = [IO.File]::ReadAllText($path).Replace([string][char]13 + [char]10,[string][char]10)
    $match = [regex]::Match($raw,'(?m)^section\s')
    if (-not $match.Success) { throw ('Missing M section: ' + $path) }
    $raw.Substring($match.Index).Trim()
}
$manifest = [ordered]@{
    Operation='User-authorized exact BD workbook copies to TE and JH-RY'
    Status='Preparing'; FilesPerUnit=$files.Count; Units=@('TE','JH-RY')
    SourceWorkbooks=@(); Backups=@(); Copies=@()
    WorkbookRefreshPerformed=$false; PowerQuerySyncPerformed=$false
    SourceExtractionPerformed=$true
    Note='Copied cached results remain BD results until each destination is recalculated, refreshed and saved.'
}
$manifestPath = Join-Path $runRoot 'copy-manifest.json'
function Publish { $manifest | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath $manifestPath -Encoding utf8 }
$donors = @{}
$targets = @()
$written = @()
try {
    foreach ($relative in $files) {
        $source = ScopedPath 'BD' $relative
        if (-not (Test-Path -LiteralPath ($source + '_PowerQuery.m'))) { throw ('Missing exact donor M source: ' + $relative) }
        CheckClosed $source
        $gateFolder = Join-Path $runRoot ('donor-gates/' + $relative)
        [void][IO.Directory]::CreateDirectory($gateFolder)
        $gateM = Join-Path $gateFolder 'workbook.m'
        $gateEvidence = Join-Path $gateFolder 'extraction.json'
        & $extractor -WorkbookPath $source -OutputMPath $gateM -EvidencePath $gateEvidence | Out-Null
        $parserResult = & $NodePath $validator $ParserPath $gateM
        if ($LASTEXITCODE -ne 0) { throw ('M validation failed: ' + $relative) }
        $parserResult | Set-Content -LiteralPath (Join-Path $gateFolder 'parser.json') -Encoding utf8
        $evidence = Get-Content -LiteralPath $gateEvidence -Raw | ConvertFrom-Json
        $donors[$relative] = @{Path=$source; Hash=$evidence.workbookSHA256; MPath=$gateM; MHash=(Hash $gateM)}
        $manifest.SourceWorkbooks += [ordered]@{
            File=$relative; WorkbookSHA256=$evidence.workbookSHA256; ExtractedMSHA256=(Hash $gateM)
            AdjacentMMatchesSavedWorkbook=((MCore ($source+'_PowerQuery.m')) -ceq (MCore $gateM))
            Syntax='Passed'; DonorChanged=$false
        }
        foreach ($unit in @('TE','JH-RY')) {
            $target = ScopedPath $unit $relative
            if (-not (Test-Path -LiteralPath $target)) { throw ('Missing approved target: ' + $target) }
            CheckClosed $target
            $targets += @{Unit=$unit; Relative=$relative; Path=$target; MPath=($target+'_PowerQuery.m'); Hash=(Hash $target)}
        }
    }
    $manifest.Status='Prepared'
    Publish
    if (-not $Apply) { Write-Output ('Prepared only: ' + $runRelative); exit 0 }
    # Back up every existing workbook and adjacent source before the first replacement.
    foreach ($target in $targets) {
        foreach ($path in @($target.Path,$target.MPath)) {
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $suffix = if ($path -eq $target.Path) { '' } else { '_PowerQuery.m' }
            $backupRelative = $runRelative + '/backups/' + $target.Unit + '/' + $target.Relative + $suffix
            $backup = Join-Path $repo $backupRelative
            [void][IO.Directory]::CreateDirectory((Split-Path -Parent $backup))
            $before = Hash $path
            Copy-Item -LiteralPath $path -Destination $backup
            if ((Hash $backup) -ne $before -or (Hash $path) -ne $before) { throw 'Backup verification failed.' }
            $manifest.Backups += @{Target=[IO.Path]::GetRelativePath($repo,$path); Backup=$backupRelative; SHA256=$before}
        }
    }
    $manifest.Status='Backed up'
    Publish
    foreach ($target in $targets) {
        $donor = $donors[$target.Relative]
        CheckClosed $target.Path
        if ((Hash $target.Path) -ne $target.Hash -or (Hash $donor.Path) -ne $donor.Hash) { throw 'A workbook changed after preflight.' }
        $written += $target.Path
        Copy-Item -LiteralPath $donor.Path -Destination $target.Path -Force
        if ((Hash $target.Path) -ne $donor.Hash) { throw 'Workbook copy hash differs from its donor.' }
        # Extract from this exact copied target; no M definition is written into Excel.
        $written += $target.MPath
        $targetEvidence = Join-Path $runRoot ('target-gates/' + $target.Unit + '/' + $target.Relative + '.json')
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $targetEvidence))
        & $extractor -WorkbookPath $target.Path -OutputMPath $target.MPath -EvidencePath $targetEvidence | Out-Null
        if ((Hash $target.MPath) -ne $donor.MHash -or (Hash $target.Path) -ne $donor.Hash) { throw 'Copied source/workbook verification failed.' }
        $manifest.Copies += [ordered]@{
            Source='CLIENT/DATExx-Whiddon/UNITS/BD/' + $target.Relative
            Target='CLIENT/DATExx-Whiddon/UNITS/' + $target.Unit + '/' + $target.Relative
            PreviousWorkbookSHA256=$target.Hash; WorkbookSHA256=$donor.Hash
            ExtractedSource='CLIENT/DATExx-Whiddon/UNITS/' + $target.Unit + '/' + $target.Relative + '_PowerQuery.m'
            ExtractedMSHA256=$donor.MHash; ExactWorkbookCopy=$true; ExactSourceMatch=$true
        }
        Publish
    }
    foreach ($relative in $files) {
        if ((Hash $donors[$relative].Path) -ne $donors[$relative].Hash) { throw 'A BD donor changed during the operation.' }
    }
    $manifest.Status='Completed'
    Publish
    Write-Output ([ordered]@{Status='Completed'; Workbooks=$manifest.Copies.Count; MatchingMSources=$manifest.Copies.Count; Backups=$manifest.Backups.Count; Manifest=$runRelative+'/copy-manifest.json'; Refreshed=$false} | ConvertTo-Json)
}
catch {
    $manifest.Status='Failed'; $manifest.Error=$_.Exception.Message
    $manifest.Rollback=@()
    foreach ($backup in $manifest.Backups) {
        $path = Join-Path $repo $backup.Target
        if ($written -notcontains $path) { continue }
        try {
            Copy-Item -LiteralPath (Join-Path $repo $backup.Backup) -Destination $path -Force
            $manifest.Rollback += @{Target=$backup.Target; Restored=((Hash $path) -eq $backup.SHA256)}
        } catch { $manifest.Rollback += @{Target=$backup.Target; Restored=$false; Error=$_.Exception.Message} }
    }
    Publish
    throw
}
