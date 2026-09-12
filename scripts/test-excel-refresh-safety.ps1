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
$script:RefreshState = @{ phase = ''; excel = $null; saved = $false; updatedAt = '' }
$script:pollSleeps = 0
$script:asyncCalls = 0
$script:calculateCalls = 0
$testConnection = [pscustomobject] @{ Name = 'Connection'; Type = 1; OLEDBConnection = [pscustomobject] @{ Refreshing = $false } }
$testQuery = [pscustomobject] @{ Name = 'Table'; Refreshing = $true }
$testWorkbook = [pscustomobject] @{ Connections = @($testConnection); Worksheets = @([pscustomobject] @{ Name = 'Sheet'; QueryTables = @($testQuery) }) }
$testExcel = [pscustomobject] @{ CalculationState = 1 }
$testExcel | Add-Member ScriptMethod CalculateUntilAsyncQueriesDone { $script:asyncCalls++ }
$testExcel | Add-Member ScriptMethod Calculate { $script:calculateCalls++; $this.CalculationState = 1 }
function Start-ResponsiveSleep {
    param($Seconds)
    $script:pollSleeps++
    # Connection flags are idle throughout; a query table and calculation must still block readiness.
    if ($script:pollSleeps -eq 2) { $testQuery.Refreshing = $false; $testExcel.CalculationState = 0 }
    if ($script:pollSleeps -eq 6) { $testExcel.CalculationState = 0 }
}
Wait-ExcelReadyToSave $testExcel $testWorkbook
Assert-Test ($script:asyncCalls -eq 1 -and $script:calculateCalls -eq 1) 'Async drain/calculation were bypassed.'
Assert-Test ($script:pollSleeps -eq 8) 'Readiness did not wait for both query completion and subsequent calculation completion.'
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
