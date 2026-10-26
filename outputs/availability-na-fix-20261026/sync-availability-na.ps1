[CmdletBinding()]
param([switch] $CheckOnly)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$auditRelative = 'outputs/availability-na-fix-20261026'

# Load only the approved comparison and existing-query update functions.
# Do not execute either helper's original, differently scoped workflow.
foreach ($helper in @(
    @{Path='outputs/te-availability-role-fix/sync-powerquery.ps1'; Names=@('Get-Sha256','Normalize-Formula','Read-MDefinitions','Release-ComObject')},
    @{Path='outputs/stafflist-dynamic-roster-20261022/sync-stafflist.ps1'; Names=@('Set-VerifiedExistingQueries')}
)) {
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $helper.Path), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Approved helper did not parse: $($helper.Path)" }
    foreach ($name in $helper.Names) {
        $definition = @($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]}, $false) | Where-Object Name -eq $name)
        if ($definition.Count -ne 1) { throw "Missing approved helper: $name" }
        . ([scriptblock]::Create($definition[0].Extent.Text))
    }
}

function Assert-ClosedWorkbook([string] $Path) {
    $ownerFile = Join-Path (Split-Path $Path -Parent) '~$Capacity-ShiftAvailability.xlsx'
    if (Test-Path -LiteralPath $ownerFile) { throw "Workbook is open: $Path" }
    $stream = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
    $stream.Dispose()
}

function Assert-NoRefreshOnOpen([string] $Path) {
    # Inspect existing refresh flags without changing the workbook package.
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -ne 'xl/connections.xml' -and $entry.FullName -notmatch '^xl/queryTables/[^/]+\.xml$') { continue }
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            if (@($xml.SelectNodes('//*[@refreshOnLoad="1" or @refreshOnLoad="true"]')).Count) {
                throw "Automatic refresh on opening is enabled in $($entry.FullName): $Path"
            }
        }
    } finally { $archive.Dispose() }
}

$syntax = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'syntax-validation.json') -Raw | ConvertFrom-Json
$validated = @(Get-Content -LiteralPath (Join-Path $PSScriptRoot 'source-validation.json') -Raw | ConvertFrom-Json)
if ($syntax.Status -cne 'Passed' -or $syntax.Results.Count -ne 3 -or $validated.Count -ne 3) { throw 'Missing three-source validation.' }
$targets = @()
foreach ($unit in @('BD','TE','JH-RY')) {
    $workbookRelative = "CLIENT/DATExx-Whiddon/UNITS/$unit/2. Calculations/Capacity-ShiftAvailability.xlsx"
    $sourceRelative = $workbookRelative + '_PowerQuery.m'
    $workbookPath = Join-Path $repoRoot $workbookRelative
    $sourcePath = Join-Path $repoRoot $sourceRelative
    $boundary = @($validated | Where-Object Unit -eq $unit)
    if ($boundary.Count -ne 1 -or $boundary[0].Workbook -cne $workbookRelative -or $boundary[0].Source -cne $sourceRelative -or
        (Get-Sha256 $sourcePath) -ne $boundary[0].SourceSHA256) { throw "Validated source changed: $unit" }
    Assert-ClosedWorkbook $workbookPath
    Assert-NoRefreshOnOpen $workbookPath
    $unitFolder = Join-Path $PSScriptRoot $unit
    $beforePath = Join-Path $unitFolder 'pre-sync.saved.m'
    $beforeEvidencePath = Join-Path $unitFolder 'pre-sync.extraction.json'
    & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $workbookPath `
        -OutputMPath $beforePath -EvidencePath $beforeEvidencePath | Out-Null
    $beforeEvidence = Get-Content -LiteralPath $beforeEvidencePath -Raw | ConvertFrom-Json
    $before = Read-MDefinitions $beforePath
    $approved = Read-MDefinitions $sourcePath
    if ($before.Count -ne 79 -or $approved.Count -ne 79 -or @($before.Keys | Where-Object {-not $approved.Contains($_)}).Count) {
        throw "Query names or count changed: $unit"
    }
    $changed = @($approved.Keys | Where-Object {(Normalize-Formula $approved[$_]) -cne (Normalize-Formula $before[$_])})
    if ($changed.Count -gt 1 -or @($changed | Where-Object {$_ -cne 'WorkerRoleMappings_Prepare'}).Count) {
        throw "Unrelated source/workbook differences: $unit/$($changed -join ', ')"
    }
    if ($changed.Count) {
        $expected = (Normalize-Formula $before['WorkerRoleMappings_Prepare']).Replace('    // Ignore wholly empty list items;',
            "    // NA and N/A are placeholder roster labels, not roles. CS may still have DC Category = NA.`n    WithoutPlaceholderRoles = Table.SelectRows(RosterRoleKeys,`n        each [RosterRoleKey] <> `"NA`"),`n    // Ignore wholly empty list items;").Replace(
            'NonEmptyDefinitions = Table.SelectRows(RosterRoleKeys,','NonEmptyDefinitions = Table.SelectRows(WithoutPlaceholderRoles,')
        if ((Normalize-Formula $approved['WorkerRoleMappings_Prepare']) -cne $expected) { throw "Unexpected target-query change: $unit" }
    }
    $backupRelative = "$auditRelative/$unit/before-sync.xlsx"
    $backupPath = Join-Path $repoRoot $backupRelative
    if (-not (Test-Path -LiteralPath $backupPath)) { Copy-Item -LiteralPath $workbookPath -Destination $backupPath }
    if ((Get-Sha256 $backupPath) -ne $beforeEvidence.workbookSHA256) { throw "Backup differs from the closed workbook: $unit" }
    $targets += [pscustomobject]@{Unit=$unit;Workbook=$workbookRelative;WorkbookPath=$workbookPath;Source=$sourceRelative;SourcePath=$sourcePath;
        Before=$before;Approved=$approved;WorkbookSHA256=$beforeEvidence.workbookSHA256;SourceSHA256=$boundary[0].SourceSHA256;
        Backup=$backupRelative;ChangedQueries=$changed}
}
$targets | Select-Object Unit,Workbook,Source,WorkbookSHA256,SourceSHA256,Backup,ChangedQueries |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'sync-preflight.json') -Encoding utf8
if ($CheckOnly) { Write-Host 'PASS: all three closed target pairs validated; backups verified; no Excel opened.'; return }

$excel = $null
$receipts = [Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    foreach ($target in $targets) {
        Assert-ClosedWorkbook $target.WorkbookPath
        if ((Get-Sha256 $target.WorkbookPath) -ne $target.WorkbookSHA256 -or
            (Get-Sha256 $target.SourcePath) -ne $target.SourceSHA256) { throw "Target changed after preflight: $($target.Unit)" }
        $workbook = $null; $queries = $null
        $updated = @()
        try {
            $workbook = $excel.Workbooks.Open($target.WorkbookPath, 0, $false)
            if ($workbook.ReadOnly -or -not [string]::Equals([string]$workbook.FullName,$target.WorkbookPath,[StringComparison]::OrdinalIgnoreCase)) {
                throw "Excel opened an incorrect or read-only workbook: $($target.Unit)"
            }
            $queries = $workbook.Queries
            $updated = @(Set-VerifiedExistingQueries $queries $target.Before $target.Approved | ForEach-Object {$_})
            if (($updated -join ',') -cne ($target.ChangedQueries -join ',')) { throw "Unexpected native query changes: $($target.Unit)" }
            $workbook.Save()
            if (-not $workbook.Saved) { throw "Excel did not confirm save: $($target.Unit)" }
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            Release-ComObject $queries
            Release-ComObject $workbook
        }
        $afterPath = Join-Path (Join-Path $PSScriptRoot $target.Unit) 'after.saved.m'
        $afterEvidence = Join-Path (Join-Path $PSScriptRoot $target.Unit) 'after.extraction.json'
        & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $target.WorkbookPath `
            -OutputMPath $afterPath -EvidencePath $afterEvidence | Out-Null
        $after = Read-MDefinitions $afterPath
        if ($after.Count -ne 79 -or @($after.Keys | Where-Object {-not $target.Approved.Contains($_)}).Count) { throw "Saved query names differ: $($target.Unit)" }
        foreach ($name in $target.Approved.Keys) {
            if ((Normalize-Formula $after[$name]) -cne (Normalize-Formula $target.Approved[$name])) {
                throw "Saved query differs from approved source: $($target.Unit)/$name"
            }
        }
        Assert-ClosedWorkbook $target.WorkbookPath
        $receipt = [ordered]@{Unit=$target.Unit;Workbook=$target.Workbook;Source=$target.Source;SourceSHA256=$target.SourceSHA256;
            WorkbookSHA256Before=$target.WorkbookSHA256;WorkbookSHA256After=(Get-Sha256 $target.WorkbookPath);Backup=$target.Backup;
            UpdatedQueries=$updated;VerifiedQueries=$after.Count;ReExtractionMatches=$true;Saved=$true;Closed=$true;
            RefreshInvoked=$false;RecordedAt=(Get-Date).ToString('o')}
        $receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path (Join-Path $PSScriptRoot $target.Unit) 'sync-receipt.json') -Encoding utf8
        $receipts.Add($receipt)
        Write-Host "PASS: $($target.Unit): $($after.Count) saved queries match; workbook closed; no refresh invoked."
    }
} finally {
    if ($null -ne $excel) { $excel.Quit() }
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$receipts | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'sync-summary.json') -Encoding utf8
