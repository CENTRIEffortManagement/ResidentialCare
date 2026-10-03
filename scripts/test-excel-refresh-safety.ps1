param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')

function Assert-Test { param($Condition, [string] $Message) if (-not $Condition) { throw $Message } }
$script:testMessages = [Collections.Generic.List[string]]::new()
function Write-Log { param($Message, $Level) $script:testMessages.Add($Message); Write-Host $Message }
function Write-CurrentStatus { param($State, $Message) $script:lastStatusMessage = $Message }
function Test-StopNowRequested {
    return ((Test-Path $script:testStopPath) -and ((Get-Content $script:testStopPath) -contains 'mode=Now'))
}
function Assert-NotStopNowRequested { if (Test-StopNowRequested) { throw 'Stop requested.' } }

$testFolder = Join-Path ([IO.Path]::GetTempPath()) ('rc-refresh-tests-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testFolder)
$script:testStopPath = Join-Path $testFolder 'stop.txt'
$WorkerStatePath = Join-Path $testFolder 'readiness.json'
$script:RefreshState = @{ phase = ''; detail = $null; excel = $null; saved = $false; updatedAt = '' }
$script:pollSleeps = 0
$script:sleepDurations = [Collections.Generic.List[int]]::new()
$script:asyncCalls = 0
$script:calculateCalls = 0
$backgroundTarget = [pscustomobject] @{ BackgroundQuery = $true; Refreshing = $false }
$backgroundConnection = [pscustomobject] @{ Name = 'Background'; Type = 1; OLEDBConnection = $backgroundTarget }
$backgroundWorkbook = [pscustomobject] @{ Connections = @($backgroundConnection); Worksheets = @() }
$backgroundSettings = @(Disable-WorkbookBackgroundRefresh $backgroundWorkbook)
Assert-Test (-not $backgroundTarget.BackgroundQuery -and $backgroundSettings.Count -eq 1) 'Supported background refresh was not disabled.'
Restore-WorkbookBackgroundRefresh $backgroundSettings
Assert-Test $backgroundTarget.BackgroundQuery 'Original background-refresh setting was not restored.'
$listBackgroundQuery = [pscustomobject] @{ BackgroundQuery = $true; Refreshing = $false }
$listBackgroundWorkbook = [pscustomobject] @{
    Connections = @()
    Worksheets = @([pscustomobject] @{
        Name = 'Demand'; QueryTables = @()
        ListObjects = @([pscustomobject] @{ Name = 'LoadedDemand'; SourceType = 3; QueryTable = $listBackgroundQuery })
    })
}
$listBackgroundSettings = @(Disable-WorkbookBackgroundRefresh $listBackgroundWorkbook)
Assert-Test (-not $listBackgroundQuery.BackgroundQuery -and $listBackgroundSettings.Count -eq 1) 'ListObject background refresh was not disabled.'
Restore-WorkbookBackgroundRefresh $listBackgroundSettings
Assert-Test $listBackgroundQuery.BackgroundQuery 'ListObject background-refresh setting was not restored.'

$refreshTarget = [pscustomobject] @{ RefreshDate = [datetime] '2026-09-30T10:00:00' }
$refreshConnection = [pscustomobject] @{
    Name = 'Query - DemandIntervals'; Type = 1; RefreshWithRefreshAll = $true; OLEDBConnection = $refreshTarget
}
$refreshQuery = [pscustomobject] @{
    Name = 'DemandIntervals'; EnableRefresh = $true; WorkbookConnection = $refreshConnection
}
$refreshWorkbook = [pscustomobject] @{
    Worksheets = @([pscustomobject] @{ Name = 'Demand'; QueryTables = @($refreshQuery) })
}
$refreshBefore = Get-WorksheetQueryRefreshEvidence $refreshWorkbook
$refreshTarget.RefreshDate = [datetime] '2026-10-01T10:00:00'
Assert-WorksheetQueryRefreshEvidence $refreshWorkbook $refreshBefore
$staleRejected = $false
try { Assert-WorksheetQueryRefreshEvidence $refreshWorkbook (Get-WorksheetQueryRefreshEvidence $refreshWorkbook) }
catch { $staleRejected = $_.Exception.Message -like '*did not report a new connection refresh date*' }
Assert-Test $staleRejected 'An unchanged query refresh date was accepted.'
$refreshConnection.RefreshWithRefreshAll = $false
$excludedRejected = $false
try { Get-WorksheetQueryRefreshEvidence $refreshWorkbook | Out-Null }
catch { $excludedRejected = $_.Exception.Message -like '*excluded from Refresh All*' }
Assert-Test $excludedRejected 'A loaded query excluded from Refresh All was accepted.'
$refreshConnection.RefreshWithRefreshAll = $true
$listWorkbook = [pscustomobject] @{
    Worksheets = @([pscustomobject] @{
        Name = 'Demand'; QueryTables = @();
        ListObjects = @([pscustomobject] @{ Name = 'Table_ShiftDemandUnitINTERVAL'; SourceType = 3; QueryTable = $refreshQuery })
    })
}
$listBefore = Get-WorksheetQueryRefreshEvidence $listWorkbook
$refreshTarget.RefreshDate = [datetime] '2026-10-02T10:00:00'
Assert-WorksheetQueryRefreshEvidence $listWorkbook $listBefore
$listStaleRejected = $false
try { Assert-WorksheetQueryRefreshEvidence $listWorkbook (Get-WorksheetQueryRefreshEvidence $listWorkbook) }
catch { $listStaleRejected = $_.Exception.Message -like '*did not report a new connection refresh date*' }
Assert-Test $listStaleRejected 'An unchanged ListObject query refresh date was accepted.'

# Excel can return no previous refresh date on an opened workbook. That is a
# valid baseline, but saving still requires a timestamp from the current refresh.
foreach ($evidenceWorkbook in @($refreshWorkbook, $listWorkbook)) {
    foreach ($emptyDate in @($null, [DBNull]::Value, '')) {
        $refreshTarget.RefreshDate = $emptyDate
        $withoutPreviousDate = Get-WorksheetQueryRefreshEvidence $evidenceWorkbook
        $evidenceLabel = @($withoutPreviousDate.Keys)[0]
        Assert-Test ($null -eq $withoutPreviousDate[$evidenceLabel].RefreshDate) 'An empty previous refresh date was not preserved as unknown.'

        $missingAfterRejected = $false
        try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $withoutPreviousDate }
        catch { $missingAfterRejected = $_.Exception.Message -like '*no refresh date after refresh*' -and $_.Exception.Message.Contains($evidenceLabel) }
        Assert-Test $missingAfterRejected 'Missing refresh evidence after refresh was accepted or did not identify the query.'

        $refreshTarget.RefreshDate = [datetime] '2000-01-01T00:00:00'
        $oldDateRejected = $false
        try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $withoutPreviousDate }
        catch { $oldDateRejected = $_.Exception.Message -like '*predates this refresh*' }
        Assert-Test $oldDateRejected 'An old timestamp was accepted when the previous date was missing.'

        # Excel dates can have only second precision; a refresh in the same
        # second as the baseline capture must still pass.
        $captured = $withoutPreviousDate[$evidenceLabel].CapturedAt
        $refreshTarget.RefreshDate = $captured.AddTicks(-($captured.Ticks % [TimeSpan]::TicksPerSecond))
        Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $withoutPreviousDate
    }
}
$refreshTarget.RefreshDate = Get-Date
$knownBefore = Get-WorksheetQueryRefreshEvidence $refreshWorkbook
$refreshTarget.RefreshDate = $null
$lostDateRejected = $false
try { Assert-WorksheetQueryRefreshEvidence $refreshWorkbook $knownBefore }
catch { $lostDateRejected = $_.Exception.Message -like '*no refresh date after refresh*' }
Assert-Test $lostDateRejected 'A previously known refresh date becoming unavailable was accepted.'
$refreshTarget.RefreshDate = 'not-a-date'
$invalidDateRejected = $false
try { Get-WorksheetQueryRefreshEvidence $refreshWorkbook | Out-Null }
catch { $invalidDateRejected = $true }
Assert-Test $invalidDateRejected 'An invalid nonempty refresh date was treated as a missing baseline.'
$refreshTarget.RefreshDate = Get-Date
Write-Host 'PASS: missing baseline dates, current refresh evidence and stale/missing/invalid date rejection.'

$emptyRejected = $false
try { Get-WorksheetQueryRefreshEvidence ([pscustomobject] @{ Worksheets = @([pscustomobject] @{ Name = 'Empty'; QueryTables = @() }) }) | Out-Null }
catch { $emptyRejected = $_.Exception.Message -like '*No loaded worksheet queries*' }
Assert-Test $emptyRejected 'A DemandIntervals workbook with no verifiable query table was accepted.'
Write-Host 'PASS: loaded query participation and refresh-date evidence.'

# Synthetic package metadata reproduces Excel retaining a worksheet QueryTable
# after its connection was explicitly deleted. No real workbook is created/opened.
function New-DeletedQueryMetadataFixture {
    param([string] $Deleted = '1', [string] $ConnectionId = '4')
    $path = Join-Path $testFolder ([guid]::NewGuid().ToString('N') + '.zip')
    $main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rels = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $parts = @{
        'xl/workbook.xml' = "<workbook xmlns='$main' xmlns:r='$office'><sheets><sheet name='Intervals' sheetId='1' r:id='rId1'/><sheet name='Other' sheetId='2' r:id='rId2'/></sheets></workbook>"
        'xl/_rels/workbook.xml.rels' = "<Relationships xmlns='$rels'><Relationship Id='rId1' Type='$office/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='rId2' Type='$office/worksheet' Target='/xl/worksheets/sheet2.xml'/></Relationships>"
        'xl/connections.xml' = "<connections xmlns='$main'><connection id='4' name='Query - Intervals' deleted='$Deleted'/><connection id='10' name='Query - Active'/></connections>"
        'xl/worksheets/_rels/sheet1.xml.rels' = "<Relationships xmlns='$rels'><Relationship Id='rId1' Type='$office/queryTable' Target='../queryTables/queryTable1.xml'/><Relationship Id='rId2' Type='$office/table' Target='../tables/table1.xml'/></Relationships>"
        'xl/worksheets/_rels/sheet2.xml.rels' = "<Relationships xmlns='$rels'><Relationship Id='rId1' Type='$office/queryTable' Target='../queryTables/queryTable2.xml'/></Relationships>"
        'xl/queryTables/queryTable1.xml' = "<queryTable xmlns='$main' name='ExternalData_7' connectionId='$ConnectionId'/>"
        'xl/queryTables/queryTable2.xml' = "<queryTable xmlns='$main' name='ExternalData_7' connectionId='10'/>"
    }
    $archive = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($part in $parts.Keys) {
            $writer = [IO.StreamWriter]::new($archive.CreateEntry($part).Open())
            try { $writer.Write($parts[$part]) } finally { $writer.Dispose() }
        }
    }
    finally { $archive.Dispose() }
    return $path
}

$orphanQuery = [pscustomobject] @{ Name = 'ExternalData_7'; EnableRefresh = $true; WorkbookConnection = $null }
$orphanSheet = [pscustomobject] @{ Name = 'Intervals'; QueryTables = @($orphanQuery) }
$mixedWorkbook = [pscustomobject] @{ Worksheets = @($orphanSheet, $refreshWorkbook.Worksheets[0]) }
$missingConnectionRejected = $false
try { Get-WorksheetQueryRefreshEvidence $mixedWorkbook | Out-Null }
catch { $missingConnectionRejected = $_.Exception.Message -like '*has no workbook connection*' }
Assert-Test $missingConnectionRejected 'A missing connection without saved deletion evidence was accepted.'

foreach ($deletedFlag in @('1', 'true')) {
    $metadataPath = New-DeletedQueryMetadataFixture -Deleted $deletedFlag
    $metadataHash = (Get-FileHash -LiteralPath $metadataPath).Hash
    $deletedQueries = Get-DeletedWorksheetQueryMetadata -Path $metadataPath
    Assert-Test ($deletedQueries.Count -eq 1 -and $deletedQueries.ContainsKey('Intervals/ExternalData_7')) 'Deleted query was not identified by sheet, query and connection ID.'
    Assert-Test ($metadataHash -eq (Get-FileHash -LiteralPath $metadataPath).Hash) 'Reading connection metadata changed the package.'
    $mixedBefore = Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries
    Assert-Test ($mixedBefore.Count -eq 1 -and $mixedBefore.ContainsKey('Demand/DemandIntervals')) 'An active output was skipped or a deleted query was required.'
    $refreshTarget.RefreshDate = $refreshTarget.RefreshDate.AddSeconds(1)
    Assert-WorksheetQueryRefreshEvidence $mixedWorkbook $mixedBefore -DeletedQueries $deletedQueries
    $staleMixedRejected = $false
    try { Assert-WorksheetQueryRefreshEvidence $mixedWorkbook (Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries) -DeletedQueries $deletedQueries }
    catch { $staleMixedRejected = $_.Exception.Message -like '*did not report a new connection refresh date*' }
    Assert-Test $staleMixedRejected 'Deleted-query exemption bypassed active output freshness checks.'
}
foreach ($fixtureCase in @(@{ Deleted = '0'; ConnectionId = '4' }, @{ Deleted = 'false'; ConnectionId = '4' }, @{ Deleted = '1'; ConnectionId = '99' })) {
    $noDeletion = Get-DeletedWorksheetQueryMetadata -Path (New-DeletedQueryMetadataFixture @fixtureCase)
    Assert-Test ($noDeletion.Count -eq 0) 'An active or unknown connection was treated as explicitly deleted.'
}
$orphanSheet.Name = 'Other'
$wrongSheetRejected = $false
try { Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries | Out-Null }
catch { $wrongSheetRejected = $_.Exception.Message -like '*has no workbook connection*' }
Assert-Test $wrongSheetRejected 'A deleted query exemption leaked to another sheet with the same query name.'
$orphanSheet.Name = 'Intervals'
$orphanQuery.EnableRefresh = $false
$null = Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries
$orphanQuery.EnableRefresh = $true
$orphanQuery.WorkbookConnection = $refreshConnection
$reactivatedBefore = Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries
Assert-Test ($reactivatedBefore.Count -eq 2) 'A reconnected query was skipped because of old deletion metadata.'
$orphanQuery.WorkbookConnection = $null
$refreshTarget.RefreshDate = $refreshTarget.RefreshDate.AddSeconds(1)
$disappearedRejected = $false
try { Assert-WorksheetQueryRefreshEvidence $mixedWorkbook $reactivatedBefore -DeletedQueries $deletedQueries }
catch { $disappearedRejected = $_.Exception.Message -like '*disappeared during refresh*' }
Assert-Test $disappearedRejected 'An active query losing its connection during refresh was accepted.'
$onlyDeletedRejected = $false
try { Get-WorksheetQueryRefreshEvidence ([pscustomobject] @{ Worksheets = @($orphanSheet) }) -DeletedQueries $deletedQueries | Out-Null }
catch { $onlyDeletedRejected = $_.Exception.Message -like '*No loaded worksheet queries*' }
Assert-Test $onlyDeletedRejected 'A workbook with only deleted queries was accepted without active refresh evidence.'
$orphanListWorkbook = [pscustomobject] @{ Worksheets = @([pscustomobject] @{
    Name = 'Intervals'; QueryTables = @(); ListObjects = @([pscustomobject] @{ Name = 'ExternalData_7'; SourceType = 3; QueryTable = $orphanQuery })
}, $refreshWorkbook.Worksheets[0]) }
$listOrphanRejected = $false
try { Get-WorksheetQueryRefreshEvidence $orphanListWorkbook -DeletedQueries $deletedQueries | Out-Null }
catch { $listOrphanRejected = $_.Exception.Message -like '*has no workbook connection*' }
Assert-Test $listOrphanRejected 'A worksheet-only deleted query exemption was applied to a loaded table.'
Assert-Test (@($script:testMessages | Where-Object { $_ -like '*Skipping deleted worksheet query*Intervals/ExternalData_7*' }).Count -gt 0) 'Skipped deleted queries were not logged.'
Write-Host 'PASS: deleted worksheet query metadata, exact matching and active output protections.'

# Power Query can finish loading a worksheet without publishing RefreshDate.
# A foreground refresh must explicitly succeed before that output can be saved.
$refreshQuery | Add-Member NoteProperty BackgroundQuery $false
$refreshQuery | Add-Member NoteProperty Refreshing $false
$refreshQuery | Add-Member NoteProperty FetchedRowOverflow $false
$refreshQuery | Add-Member NoteProperty RefreshCalls 0
$refreshQuery | Add-Member NoteProperty RefreshResult $true
$refreshQuery | Add-Member NoteProperty RefreshError ''
$refreshQuery | Add-Member ScriptMethod Refresh {
    param($Background)
    if ($Background -ne $false) { throw 'Fixture requires foreground refresh.' }
    $this.RefreshCalls++
    if ($this.RefreshError) { throw $this.RefreshError }
    return $this.RefreshResult
}
foreach ($evidenceWorkbook in @($refreshWorkbook, $listWorkbook)) {
    $refreshTarget.RefreshDate = $null
    $baselineWithoutDate = Get-WorksheetQueryRefreshEvidence $evidenceWorkbook
    $callsBefore = $refreshQuery.RefreshCalls
    $confirmed = Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $baselineWithoutDate -ConfirmMissingDates
    Assert-Test ($confirmed -eq $true -and $refreshQuery.RefreshCalls -eq $callsBefore + 1) 'A missing timestamp was not verified by a successful foreground refresh.'
    foreach ($badResult in @($false, $null, 'True')) {
        $refreshQuery.RefreshResult = $badResult
        $rejected = $false
        try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $baselineWithoutDate -ConfirmMissingDates | Out-Null }
        catch { $rejected = $_.Exception.Message -like '*did not confirm successful foreground refresh*' }
        Assert-Test $rejected 'A cancelled or unknown foreground refresh result was accepted.'
    }
    $refreshQuery.RefreshResult = $true
    foreach ($badProperty in @('Refreshing', 'FetchedRowOverflow', 'BackgroundQuery')) {
        foreach ($badValue in @($true, $null)) {
            $refreshQuery.$badProperty = $badValue
            $rejected = $false
            try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $baselineWithoutDate -ConfirmMissingDates | Out-Null }
            catch { $rejected = $_.Exception.Message -like '*foreground refresh*' }
            Assert-Test $rejected "Unknown or true $badProperty was accepted as a completed foreground refresh."
        }
        $refreshQuery.$badProperty = $false
    }
    $refreshQuery.RefreshError = 'Synthetic provider failure'
    $providerRejected = $false
    try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $baselineWithoutDate -ConfirmMissingDates | Out-Null }
    catch { $providerRejected = $_.Exception.Message -like '*foreground refresh*Synthetic provider failure*' }
    Assert-Test $providerRejected 'A provider failure was swallowed or lost its query context.'
    $refreshQuery.RefreshError = ''
    $refreshTarget.RefreshDate = Get-Date
    $datedBefore = Get-WorksheetQueryRefreshEvidence $evidenceWorkbook
    $callsBefore = $refreshQuery.RefreshCalls
    $staleRejected = $false
    try { Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $datedBefore -ConfirmMissingDates | Out-Null }
    catch { $staleRejected = $_.Exception.Message -like '*did not report a new connection refresh date*' }
    Assert-Test ($staleRejected -and $refreshQuery.RefreshCalls -eq $callsBefore) 'Foreground fallback bypassed a known stale timestamp.'
    $refreshTarget.RefreshDate = $refreshTarget.RefreshDate.AddSeconds(1)
    $confirmed = Assert-WorksheetQueryRefreshEvidence $evidenceWorkbook $datedBefore -ConfirmMissingDates
    Assert-Test ($confirmed -eq $false -and $refreshQuery.RefreshCalls -eq $callsBefore) 'A current timestamp unnecessarily triggered another refresh.'
}
$refreshTarget.RefreshDate = $null
$mixedBefore = Get-WorksheetQueryRefreshEvidence $mixedWorkbook -DeletedQueries $deletedQueries
$null = Assert-WorksheetQueryRefreshEvidence $mixedWorkbook $mixedBefore -DeletedQueries $deletedQueries -ConfirmMissingDates
# The deleted object has no Refresh method, so the preceding check also proves
# that foreground confirmation never attempts to refresh deleted connections.
$refreshTarget.RefreshDate = Get-Date
Write-Host 'PASS: missing timestamp confirmation, cancellation, provider failures, asynchronous/unknown state and row overflow.'

$testConnection = [pscustomobject] @{ Name = 'Connection'; Type = 1; OLEDBConnection = [pscustomobject] @{ Refreshing = $false } }
$testQuery = [pscustomobject] @{ Name = 'Table'; Refreshing = $true }
$testWorkbook = [pscustomobject] @{ Connections = @($testConnection); Worksheets = @([pscustomobject] @{ Name = 'Sheet'; QueryTables = @($testQuery) }) }
$testExcel = [pscustomobject] @{ CalculationState = 1 }
$listActivityQuery = [pscustomobject] @{ Refreshing = $true }
$listActivityWorkbook = [pscustomobject] @{
    Connections = @()
    Worksheets = @([pscustomobject] @{
        Name = 'Demand'; QueryTables = @()
        ListObjects = @([pscustomobject] @{ Name = 'LoadedDemand'; SourceType = 3; QueryTable = $listActivityQuery })
    })
}
Assert-Test ('Demand/LoadedDemand' -in (Get-ExcelActivity $testExcel $listActivityWorkbook).Active) 'Active ListObject query was not detected.'
$testExcel | Add-Member ScriptMethod CalculateUntilAsyncQueriesDone { $script:asyncCalls++ }
$testExcel | Add-Member ScriptMethod Calculate { $script:calculateCalls++; $this.CalculationState = 1 }
function Start-ResponsiveSleep {
    param($Seconds)
    $script:pollSleeps++
    $script:sleepDurations.Add([int] $Seconds)
    # The async drain runs first; a query table and calculation must still block final readiness.
    if ($script:pollSleeps -eq 2) { $testQuery.Refreshing = $false; $testExcel.CalculationState = 0 }
    if ($script:pollSleeps -eq 6) { $testExcel.CalculationState = 0 }
}
Wait-ExcelReadyToSave $testExcel $testWorkbook
Assert-Test ($script:asyncCalls -eq 0 -and $script:calculateCalls -eq 1) 'Unsafe async drain was called or calculation was bypassed.'
Assert-Test ($script:pollSleeps -eq 8) 'Connection and calculation readiness gates did not wait for completion.'
Assert-Test ($script:sleepDurations[0] -eq 15) 'The pre-drain Power Query settling interval was omitted.'
Assert-Test ([string]::IsNullOrWhiteSpace([string] $script:RefreshState.detail)) 'Successful readiness left a stale COM probe detail.'
$originalActivityFunction = ${function:Get-ExcelActivity}
$script:transientReads = 1
function Get-ExcelActivity {
    param($Excel, $Workbook)
    if ($script:transientReads -gt 0) { $script:transientReads--; throw 'Transient null connection.' }
    & $originalActivityFunction $Excel $Workbook
}
$script:pollSleeps = 0
Wait-ExcelQuietPolls $testExcel $testWorkbook -RequireCalculationDone
Assert-Test ($script:pollSleeps -eq 3) 'Transient read counted as quiet or was not retried.'
$nullConnectionRejected = $false
try { Get-ExcelActivity $testExcel ([pscustomobject] @{ Connections = @($null); Worksheets = @() }) | Out-Null } catch { $nullConnectionRejected = $true }
Assert-Test $nullConnectionRejected 'A null connection was treated as completed.'

$identity = Get-RefreshProcessIdentity (Get-Process -Id $PID)
Assert-Test ($null -ne (Get-MatchingRefreshProcess $identity)) 'Matching process was not recognized.'
$identity.startTicks = '0'
Assert-Test ($null -eq (Get-MatchingRefreshProcess $identity)) 'Reused PID was accepted with a different start time.'

# Exercise real child processes that simulate blocked COM calls; never open Excel/workbooks.
$fakeWorker = Join-Path $testFolder 'fake-worker.ps1'
@'
param($WorkerStatePath, $LogPath, $StopRequestPath, $Mode, $IdentityPath)
$state = @{ phase = 'Saving'; excel = $null; saved = $false; updatedAt = '' }
if ($Mode -eq 'success') {
    $state.phase = 'Complete'; $state.saved = $true
    $state.excel = @{ id = 2147483000; name = 'EXCEL'; startTicks = '0' }
}
if ($IdentityPath) {
    $state.excel = Get-Content $IdentityPath -Raw | ConvertFrom-Json
    $state.saved = $Mode -ne 'unsaved'
    $state.phase = if ($Mode -in @('linger', 'cleanup-blocked')) { 'Complete' } else { 'Closing' }
}
$state | ConvertTo-Json | Set-Content $WorkerStatePath
if ($Mode -eq 'stop') { Set-Content $StopRequestPath 'mode=Now' }
if ($Mode -in @('hang','stop')) { Start-Sleep -Seconds 300 }
exit 0
'@ | Set-Content $fakeWorker
$baseParameters = @{ LogPath = (Join-Path $testFolder 'test.log'); StopRequestPath = $script:testStopPath }
foreach ($case in @(
    @{ Mode = 'success'; Expected = 0; Timeout = 10 },
    @{ Mode = 'false-success'; Expected = 1; Timeout = 10 },
    @{ Mode = 'hang'; Expected = 124; Timeout = 2 },
    @{ Mode = 'stop'; Expected = 3; Timeout = 10 }
)) {
    $parameters = @{} + $baseParameters
    $parameters.Mode = $case.Mode
    $actual = Invoke-SupervisedExcelRefresh -WorkerScript $fakeWorker -Parameters $parameters -TimeoutSeconds $case.Timeout -StopGraceSeconds 0
    Assert-Test ($actual -eq $case.Expected) "Case $($case.Mode): expected $($case.Expected), got $actual"
    Write-Host "PASS: $($case.Mode)"
}
Write-Host 'PASS: async drain, active query table/calculation, quiet polls, unknown state, process identity and supervisor cases.'

# Simulate lingering Excel and Mashup with owned disposable PowerShell processes.
# Keep real identity checks and termination; mock only child discovery (no WMI/Excel).
Set-Content $script:testStopPath 'mode=None'
$originalChildrenFunction = ${function:Update-RefreshChildren}
$originalStopFunction = ${function:Stop-OwnedRefreshProcesses}
function Update-RefreshChildren {
    param($ExcelIdentity, [hashtable] $Children)
    $Children['test-child'] = $script:testChildIdentity
}
foreach ($mode in @('linger', 'incomplete', 'unsaved', 'cleanup-blocked')) {
    $disposableProcesses = @()
    try {
        foreach ($index in 1..2) {
            $startInfo = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh).Source)
            $startInfo.UseShellExecute = $false
            $startInfo.CreateNoWindow = $true
            $startInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
            foreach ($argument in @('-NoProfile', '-Command', 'Start-Sleep -Seconds 300')) {
                $startInfo.ArgumentList.Add($argument)
            }
            $disposableProcesses += [Diagnostics.Process]::Start($startInfo)
        }
        $identityPath = Join-Path $testFolder 'owned-identity.json'
        Get-RefreshProcessIdentity $disposableProcesses[0] | ConvertTo-Json | Set-Content $identityPath
        $script:testChildIdentity = Get-RefreshProcessIdentity $disposableProcesses[1]
        $parameters = @{} + $baseParameters
        $parameters.Mode = $mode
        $parameters.IdentityPath = $identityPath
        $script:testMessages.Clear()
        if ($mode -eq 'cleanup-blocked') {
            function Stop-OwnedRefreshProcesses { param($State, $Children, $Worker) }
        }
        $actual = Invoke-SupervisedExcelRefresh -WorkerScript $fakeWorker -Parameters $parameters -TimeoutSeconds 20 -ExitGraceSeconds 0
        $expected = if ($mode -eq 'linger') { 0 } else { 1 }
        Assert-Test ($actual -eq $expected) "Case ${mode}: expected $expected, got $actual"
        if ($mode -eq 'cleanup-blocked') {
            Assert-Test ($script:testMessages -match 'Owned processes remain after cleanup') 'Incomplete cleanup did not block continuation.'
        } else {
            foreach ($process in $disposableProcesses) {
                Assert-Test $process.HasExited "Case ${mode}: owned process still alive after cleanup."
            }
            $expectedSaveMessage = if ($mode -eq 'unsaved') { 'save is not confirmed' } else { 'save was confirmed' }
            Assert-Test ($script:testMessages -match $expectedSaveMessage) "Case ${mode}: inaccurate save message."
        }
        if ($actual -ne 0) {
            $expectedSaveMessage = if ($mode -eq 'unsaved') { 'save is not confirmed' } else { 'save was confirmed' }
            Assert-Test ($script:lastStatusMessage -match $expectedSaveMessage) "Case ${mode}: inaccurate final save status."
        }
        Write-Host "PASS: $mode"
    } finally {
        ${function:Stop-OwnedRefreshProcesses} = $originalStopFunction
        foreach ($process in $disposableProcesses) {
            if (-not $process.HasExited) { $process.Kill(); [void] $process.WaitForExit(10000) }
            $process.Dispose()
        }
    }
}
${function:Update-RefreshChildren} = $originalChildrenFunction

function Get-RefreshFileUsers {
    param($Path)
    return [pscustomobject] @{ AppName = 'Tableau'; Process = [pscustomobject] @{ Id = 42 } }
}
Assert-NoExternalFileUsers -Path 'mock.xlsx' -AllowedProcessId 42
$externalHolderRejected = $false
try { Assert-NoExternalFileUsers -Path 'mock.xlsx' -AllowedProcessId 99 }
catch { $externalHolderRejected = $_.Exception.Message -like '*Tableau (PID 42)*' }
Assert-Test $externalHolderRejected 'External file holder was not rejected with an actionable error.'
function Get-RefreshFileUsers { param($Path) throw 'Cannot inspect file users.' }
$unknownHoldersRejected = $false
try { Assert-NoExternalFileUsers -Path 'mock.xlsx' } catch { $unknownHoldersRejected = $true }
Assert-Test $unknownHoldersRejected 'Failed file-usage inspection was treated as safe.'
Write-Host 'PASS: own Excel holder allowed, external file holder blocked, unavailable inspection fails closed.'
