[CmdletBinding()]
param(
    [ValidatePattern('^[0-9]{8}-[0-9]{6}$')]
    [string]$RunTag = (Get-Date -Format 'yyyyMMdd-HHmmss')
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Set-Location -LiteralPath $repo
$outputRelative = 'outputs/cs-rollout-te-jhry-20261026/alignment-' + $RunTag
$outputRoot = Join-Path $repo $outputRelative
if (Test-Path -LiteralPath $outputRoot) { throw 'Alignment directory already exists.' }
[void][IO.Directory]::CreateDirectory($outputRoot)
$files = @(
    '1. Input/1-AllocationExtracted.xlsx',
    '1. Input/2-DemandExtract.xlsx',
    '1. Input/Worker Reconciliation.xlsx',
    '2. Calculations/Capacity-ShiftAvailability.xlsx',
    '2. Calculations/Capacity.xlsx',
    '2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx',
    '2. Calculations/CS/CapacityDistrib(A.2)-shifts.xlsx',
    '2. Calculations/CS/CapacityDistrib(B)-shifts.xlsx'
)
function Core([string]$path) {
    $raw = [IO.File]::ReadAllText($path).Replace([string][char]13 + [char]10,[string][char]10)
    $match = [regex]::Match($raw,'(?m)^section\s')
    if (-not $match.Success) { throw ('Missing M section: ' + $path) }
    $raw.Substring($match.Index).Trim()
}
$donors = @{}
$workbooks = @()
$exports = @()
foreach ($unit in @('BD','TE','JH-RY')) {
    foreach ($relative in $files) {
        $workbook = 'CLIENT/DATExx-Whiddon/UNITS/' + $unit + '/' + $relative
        $source = $workbook + '_PowerQuery.m'
        if (-not (Test-Path -LiteralPath $source)) { throw ('Missing exact extracted source: ' + $source) }
        $gateFolder = Join-Path $outputRoot ('extractions/' + $unit + '/' + $relative)
        [void][IO.Directory]::CreateDirectory($gateFolder)
        $extracted = Join-Path $gateFolder 'workbook.m'
        $evidencePath = Join-Path $gateFolder 'evidence.json'
        & 'outputs/te-availability-role-fix/extract-powerquery.ps1' -WorkbookPath $workbook -OutputMPath $extracted -EvidencePath $evidencePath | Out-Null
        $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
        $body = Core $extracted
        if ($unit -eq 'BD') { $donors[$relative] = @{Hash=$evidence.workbookSHA256; M=$body} }
        $exports += [pscustomobject][ordered]@{
            Unit=$unit; File=$relative; MExportMatchesSavedWorkbook=((Core $source) -ceq $body)
            WorkbookUnchanged=$evidence.workbookUnchanged
        }
        if ($unit -ne 'BD') {
            $workbooks += [pscustomobject][ordered]@{
                Unit=$unit; File=$relative
                ExactWorkbookMatchToBD=($evidence.workbookSHA256 -eq $donors[$relative].Hash)
                SavedMDefinitionsMatchBD=($body -ceq $donors[$relative].M)
                AdjacentMMatchesBD=((Core $source) -ceq $donors[$relative].M)
                WorkbookSHA256=$evidence.workbookSHA256
            }
        }
    }
}
$preview = @()
foreach ($unit in @('BD','TE','JH-RY')) {
    $path = 'CLIENT/DATExx-Whiddon/UNITS/' + $unit + '/2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx'
    $archive = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $path).Path)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notmatch '^customXml/[^/]+\.xml$') { continue }
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $doc = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($doc.DocumentElement.LocalName -ne 'DataMashup') { continue }
            $bytes = [Convert]::FromBase64String($doc.DocumentElement.InnerText)
            $text = [Text.Encoding]::UTF8.GetString($bytes)
            $match = [regex]::Match($text,'<Entry\b[^>]*Type="RunBackgroundAnalysis"[^>]*/>')
            if (-not $match.Success) {
                $text = [Text.Encoding]::Unicode.GetString($bytes)
                $match = [regex]::Match($text,'<Entry\b[^>]*Type="RunBackgroundAnalysis"[^>]*/>')
            }
            $preview += @{Unit=$unit; ExplicitPreviewOff=($match.Success -and $match.Value -match 'Value="sFalse"')}
        }
    } finally { $archive.Dispose() }
}
$donorStable = @()
foreach ($relative in $files) {
    $path = 'CLIENT/DATExx-Whiddon/UNITS/BD/' + $relative
    $stable = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -eq $donors[$relative].Hash
    $donorStable += [pscustomobject]@{File=$relative; UnchangedDuringCheck=$stable}
    if (-not $stable) { throw ('BD changed during alignment check: ' + $relative) }
}
$problems = @($workbooks | Where-Object { -not $_.ExactWorkbookMatchToBD -or -not $_.SavedMDefinitionsMatchBD -or -not $_.AdjacentMMatchesBD })
$staleExports = @($exports | Where-Object { -not $_.MExportMatchesSavedWorkbook })
$report = [ordered]@{
    Scope='Eight approved rollout workbooks in BD, TE and JH-RY only'
    Status=if ($problems.Count -eq 0) { 'All destination workbooks and sources align with saved BD definitions' } else { 'Alignment differences found' }
    DestinationWorkbooksChecked=$workbooks.Count
    Workbooks=$workbooks; SourceExports=$exports; StaleSourceExports=$staleExports
    DonorStability=$donorStable
    CS_A1_PreviewSettings=$preview
    WorkbookContentsChanged=$false; Refreshed=$false
}
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputRoot 'alignment.json') -Encoding utf8
[ordered]@{
    Status=$report.Status; DestinationWorkbooksChecked=$workbooks.Count; DestinationProblems=$problems.Count
    StaleBDExports=@($staleExports | Where-Object Unit -eq 'BD' | Select-Object File)
    CS_A1_PreviewSettings=$preview; Report=$outputRelative+'/alignment.json'; Refreshed=$false
} | ConvertTo-Json -Depth 5
if ($problems.Count -gt 0 -or @($preview | Where-Object { -not $_.ExplicitPreviewOff }).Count -gt 0) { exit 1 }
