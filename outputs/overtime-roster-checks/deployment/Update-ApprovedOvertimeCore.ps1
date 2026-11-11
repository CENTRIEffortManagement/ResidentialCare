[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Fa-f0-9]{64}$')]
    [string] $ExpectedSourceHash,
    [string] $RunFolder = '',
    [string] $LogPath = '',
    [string] $StopRequestPath = '',
    [string] $WorkerStatePath = '',
    [string] $OperationId = '',
    [string] $IntegrationFixReason = 'Approved integration fix following an unsaved live refresh failure'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$ExpectedSourceHash = $ExpectedSourceHash.ToUpperInvariant()
if (-not $RunFolder) {
    $RunFolder = (Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim()
}
if ([IO.Path]::IsPathRooted($RunFolder)) { throw 'RunFolder must be repository-relative.' }
$runRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $RunFolder))
$repoPrefix = $repoRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $runRoot.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Deployment evidence must stay inside the repository.' }
if (-not (Test-Path -LiteralPath $runRoot -PathType Container)) { throw 'The existing deployment evidence directory is required.' }
if (-not $OperationId) { $OperationId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8) }
if ($OperationId -notmatch '^[A-Za-z0-9_-]+$') { throw 'OperationId contains invalid path characters.' }
if (-not $LogPath) { $LogPath = Join-Path $runRoot "Core-update-$OperationId.log" }
if (-not $StopRequestPath) { $StopRequestPath = Join-Path $runRoot 'stop.txt' }

# The helper is deliberately fixed to the previously installed core pair.
$workbookRelative = 'CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx'
$sourceRelative = 'Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq'
$workbookPath = Join-Path $repoRoot $workbookRelative
$sourcePath = Join-Path $repoRoot $sourceRelative
$receiptPath = Join-Path $runRoot 'Core-sync-receipt.json'
$previousReceipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
if ($previousReceipt.Target -cne 'Core' -or $previousReceipt.Workbook -cne $workbookRelative -or $previousReceipt.Source -cne $sourceRelative -or
    -not $previousReceipt.ReExtractionMatches -or -not $previousReceipt.Saved -or -not $previousReceipt.Closed) {
    throw 'The previous exact-pair synchronization receipt is missing or unverified.'
}
$baselinePath = if ($previousReceipt.PSObject.Properties['ReExtractionSource']) {
    Join-Path $repoRoot $previousReceipt.ReExtractionSource
} else { Join-Path $runRoot 'Core-after-sync.m' }
$baselineEvidencePath = if ($previousReceipt.PSObject.Properties['ReExtractionEvidence']) {
    Join-Path $repoRoot $previousReceipt.ReExtractionEvidence
} else { Join-Path $runRoot 'Core-after-sync-extraction.json' }
$baselineEvidence = Get-Content -LiteralPath $baselineEvidencePath -Raw | ConvertFrom-Json
if ($baselineEvidence.workbookSHA256 -cne $previousReceipt.WorkbookSHA256After -or -not $baselineEvidence.workbookUnchanged) {
    throw 'The baseline re-extraction does not identify the currently receipted workbook.'
}

# Load only approved read/normalization/COM-cleanup helpers, never execute the old sync entry point.
$parseTokens = $null; $parseErrors = $null
$syncAst = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'), [ref] $parseTokens, [ref] $parseErrors)
if ($parseErrors.Count) { throw 'Approved synchronization mechanism has syntax errors.' }
foreach ($name in @('Get-Sha256', 'Normalize-Formula', 'Read-MDefinitions', 'Release-ComObject')) {
    $definition = @($syncAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false) | Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Missing approved helper: $name" }
    . ([scriptblock]::Create($definition[0].Extent.Text))
}
function Write-Log([string] $Message, [string] $Level = 'INFO') {
    $line = "[$([DateTime]::UtcNow.ToString('o'))] [$Level] Core update $Message"
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
function Write-CurrentStatus([string] $State, [string] $Message = '') { }
function Test-StopNowRequested {
    (Test-Path -LiteralPath $StopRequestPath) -and ((Get-Content -LiteralPath $StopRequestPath -Raw) -match '(?im)^mode\s*=\s*Now\s*$')
}
function Assert-NotStopNowRequested { if (Test-StopNowRequested) { throw 'Operator stop-now request observed.' } }
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if (-not $WorkerStatePath) {
    $workerArguments = @{
        ExpectedSourceHash = $ExpectedSourceHash; RunFolder = $RunFolder; LogPath = $LogPath
        StopRequestPath = $StopRequestPath; OperationId = $OperationId; IntegrationFixReason = $IntegrationFixReason
    }
    exit (Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters $workerArguments -TimeoutSeconds 1800)
}
$script:WorkerStatePath = [IO.Path]::GetFullPath($WorkerStatePath)
$script:RefreshState = @{ phase = 'Starting'; detail = $null; excel = $null; saved = $false; updatedAt = '' }

function Assert-Closed([string] $Path) {
    $lockPath = Join-Path (Split-Path $Path -Parent) ('~$' + [IO.Path]::GetFileName($Path))
    if (Test-Path -LiteralPath $lockPath) { throw "Excel lock file exists: $Path" }
    $handle = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
    $handle.Dispose()
}
function Get-OutputBindings($Book) {
    $bindings = foreach ($output in @(
        @{ Sheet = 'ShiftResults'; Table = 'OvertimeRosterChecks_ShiftResults' },
        @{ Sheet = 'FortnightResults'; Table = 'OvertimeRosterChecks_FortnightResults' },
        @{ Sheet = 'ExcessSegments'; Table = 'OvertimeRosterChecks_ExcessSegments' }
    )) {
        $sheet = $null; $table = $null; $queryTable = $null; $connection = $null
        try {
            $sheet = $Book.Worksheets.Item($output.Sheet)
            $table = $sheet.ListObjects.Item($output.Table)
            if ([string] $table.Name -cne $output.Table) { throw 'An existing loaded output does not match its required name.' }
            $queryTable = $table.QueryTable
            $connection = $queryTable.WorkbookConnection
            if ($null -eq $queryTable -or $null -eq $connection) { throw 'An existing loaded output has no query connection.' }
            [pscustomobject]@{ Sheet = $output.Sheet; Table = [string] $table.Name; Connection = [string] $connection.Name; CommandText = (@($queryTable.CommandText) -join "`n") }
        } finally {
            foreach ($value in @($connection, $queryTable, $table, $sheet)) { Release-ComObject $value }
        }
    }
    return ,@($bindings)
}

$excel = $null; $book = $null; $queries = $null; $ownsExcel = $false; $closed = $false; $quit = $false; $lease = $null
try {
    Set-RefreshPhase 'Preflight'
    Assert-NotStopNowRequested
    if ((Get-Sha256 $sourcePath) -cne $ExpectedSourceHash) { throw 'Approved integration-fix source hash does not match the explicit parameter.' }
    if ((Get-Sha256 $baselinePath) -cne $baselineEvidence.extractedMSHA256) { throw 'Baseline extracted definitions changed after verification.' }
    Assert-Closed $workbookPath
    if ((Get-Sha256 $workbookPath) -cne $previousReceipt.WorkbookSHA256After) { throw 'Core workbook changed after its verified synchronization.' }
    $before = Read-MDefinitions $baselinePath
    $approved = Read-MDefinitions $sourcePath
    if ($before.Count -ne $approved.Count -or @($before.Keys | Where-Object { -not $approved.Contains($_) }).Count) {
        throw 'Update-only operation requires identical existing and approved query names; no additions or removals are permitted.'
    }
    $changedNames = @($approved.Keys | Where-Object { (Normalize-Formula $before[$_]) -cne (Normalize-Formula $approved[$_]) })
    if (-not $changedNames.Count) { throw 'The approved source contains no integration-fix formula changes.' }
    . (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
    $lease = Enter-RefreshRunGate -DateRoot (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon') -Exclusive

    $backupDirectory = Join-Path $runRoot 'backups'
    [void][IO.Directory]::CreateDirectory($backupDirectory)
    $backupPath = Join-Path $backupDirectory "Core.before-update-$OperationId.xlsx"
    $previousReceiptPath = Join-Path $runRoot "Core-sync-receipt.before-update-$OperationId.json"
    if ((Test-Path -LiteralPath $backupPath) -or (Test-Path -LiteralPath $previousReceiptPath)) { throw 'Update evidence already exists; refusing to overwrite it.' }
    Copy-Item -LiteralPath $workbookPath -Destination $backupPath
    if ((Get-Sha256 $backupPath) -cne $previousReceipt.WorkbookSHA256After) { throw 'Current-workbook backup hash verification failed.' }
    Copy-Item -LiteralPath $receiptPath -Destination $previousReceiptPath
    $receiptHashBefore = Get-Sha256 $receiptPath
    $preExisting = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Assert-Closed $workbookPath
    if ((Get-Sha256 $workbookPath) -cne $previousReceipt.WorkbookSHA256After) { throw 'Core workbook changed after backup.' }
    $fileSnapshot = New-RefreshFileSnapshot -Path $workbookPath
    Set-RefreshPhase 'CreatingExcel'
    $excel = New-Object -ComObject Excel.Application
    Register-RefreshExcel -Excel $excel -PreExistingExcelIds $preExisting
    $ownsExcel = $true
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    Set-RefreshPhase 'Opening'
    $book = $excel.Workbooks.Open($workbookPath, 0, $false)
    Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $workbookPath
    $queries = $book.Queries
    # Never infer an empty collection from Queries.Count; verify all named definitions before changing any.
    foreach ($name in $before.Keys) {
        $query = $null
        try {
            $query = $queries.Item($name)
            if ($null -eq $query -or [string] $query.Name -cne $name -or
                (Normalize-Formula ([string] $query.Formula)) -cne (Normalize-Formula $before[$name])) {
                throw "Native definition differs from the verified baseline: $name"
            }
        } finally { Release-ComObject $query }
    }
    $bindingsBefore = Get-OutputBindings $book | ConvertTo-Json -Depth 4 -Compress
    Set-RefreshPhase 'UpdatingApprovedFormulas'
    foreach ($name in $changedNames) {
        Assert-NotStopNowRequested
        $query = $null
        try {
            $query = $queries.Item($name)
            $query.Formula = [string] $approved[$name]
            if ((Normalize-Formula ([string] $query.Formula)) -cne (Normalize-Formula $approved[$name])) { throw "Native formula update was not accepted: $name" }
        } finally { Release-ComObject $query }
    }
    foreach ($name in $approved.Keys) {
        $query = $null
        try {
            $query = $queries.Item($name)
            if ((Normalize-Formula ([string] $query.Formula)) -cne (Normalize-Formula $approved[$name])) { throw "Pre-save query differs: $name" }
        } finally { Release-ComObject $query }
    }
    if ((Get-OutputBindings $book | ConvertTo-Json -Depth 4 -Compress) -cne $bindingsBefore) { throw 'Existing loaded-output bindings changed during formula-only update.' }
    if ((Get-Sha256 $sourcePath) -cne $ExpectedSourceHash -or (Get-Sha256 $receiptPath) -cne $receiptHashBefore) { throw 'Approved source or current receipt changed during update.' }
    Assert-RefreshFileSnapshot $fileSnapshot
    Assert-NotStopNowRequested
    Set-RefreshPhase 'Saving'
    $book.Save()
    if (-not $book.Saved) { throw 'Excel did not confirm the formula-only update save.' }
    $script:RefreshState.saved = $true
    Set-RefreshPhase 'Closing'
    $book.Close($false); $closed = $true
    Set-RefreshPhase 'Quitting'
    $excel.Quit(); $quit = $true
    Release-ComObject $queries; $queries = $null; Release-ComObject $book; $book = $null; Release-ComObject $excel; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Assert-Closed $workbookPath
    Set-RefreshPhase 'ReExtracting'
    $afterPath = Join-Path $runRoot "Core-after-update-$OperationId.m"
    $afterEvidencePath = Join-Path $runRoot "Core-after-update-$OperationId-extraction.json"
    & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $workbookPath -OutputMPath $afterPath -EvidencePath $afterEvidencePath | Out-Null
    $after = Read-MDefinitions $afterPath
    if ($after.Count -ne $approved.Count -or @($after.Keys | Where-Object { -not $approved.Contains($_) }).Count) { throw 'Re-extracted query names differ after formula-only update.' }
    foreach ($name in $approved.Keys) {
        if ((Normalize-Formula $after[$name]) -cne (Normalize-Formula $approved[$name])) { throw "Re-extracted formula differs: $name" }
    }
    if ((Get-Sha256 $sourcePath) -cne $ExpectedSourceHash -or (Get-Sha256 $receiptPath) -cne $receiptHashBefore) { throw 'Approved source or current receipt changed before publishing update evidence.' }
    $newReceipt = [ordered]@{
        Target = 'Core'; Workbook = $workbookRelative; Source = $sourceRelative; SourceSHA256 = $ExpectedSourceHash
        WorkbookSHA256Before = $previousReceipt.WorkbookSHA256After; WorkbookSHA256After = (Get-Sha256 $workbookPath)
        RemovedQueries = @(); AddedQueries = @(); UpdatedQueries = $changedNames; VerifiedQueryCount = $after.Count
        LoadedOutputs = @($previousReceipt.LoadedOutputs); LoadedOutputBindingsPreserved = $true
        ReExtractionMatches = $true; Saved = $true; Closed = $true; Refreshed = $false
        Operation = 'Formula-only integration fix'; IntegrationFixReason = $IntegrationFixReason
        PreviousReceipt = [IO.Path]::GetRelativePath($repoRoot, $previousReceiptPath).Replace('\', '/')
        Backup = [IO.Path]::GetRelativePath($repoRoot, $backupPath).Replace('\', '/')
        ReExtractionSource = [IO.Path]::GetRelativePath($repoRoot, $afterPath).Replace('\', '/')
        ReExtractionEvidence = [IO.Path]::GetRelativePath($repoRoot, $afterEvidencePath).Replace('\', '/')
        RecordedAtUTC = [DateTime]::UtcNow.ToString('o')
    }
    $operationReceiptPath = Join-Path $runRoot "Core-update-$OperationId-receipt.json"
    $newReceipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $operationReceiptPath -Encoding utf8
    Copy-Item -LiteralPath $operationReceiptPath -Destination $receiptPath -Force
    Set-RefreshPhase 'Complete'
    Write-Log 'Formula-only update, save, close and exact re-extraction comparison passed; no refresh was invoked.'
    exit 0
} catch {
    $script:RefreshState.error = @{ Phase = $script:RefreshState.phase; Message = $_.Exception.Message }
    Write-RefreshWorkerState
    Write-Log $_.Exception.Message 'ERROR'
    throw
} finally {
    if ($null -ne $book -and -not $closed) { try { $book.Close($false) } catch { Write-Log $_.Exception.Message 'WARN' } }
    if ($null -ne $excel -and $ownsExcel -and -not $quit) { try { $excel.Quit() } catch { Write-Log $_.Exception.Message 'WARN' } }
    foreach ($value in @($queries, $book, $excel)) { Release-ComObject $value }
    if ($null -ne $lease) { $lease.Dispose() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
