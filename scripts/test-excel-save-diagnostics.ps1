# Synthetic objects/files only. Executes the actual worker save guard, not Excel.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-save-diagnostic-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
$WorkerStatePath = Join-Path $testRoot 'worker.json'
$script:RefreshState = @{ phase=''; saved=$false; updatedAt=''; detail=$null }
$script:checks = 0
$script:sleeps = 0
$script:becomeCleanOnSleep = $false
function Write-Log { param($Message, $Level) }
function Write-CurrentStatus { param($State, $Message) }
function Assert-NotStopNowRequested { }
function Assert-SaveTest {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}
function Start-ResponsiveSleep {
    param($Seconds)
    $script:sleeps++
    if ($script:becomeCleanOnSleep) { $script:fixtureWorkbook.Saved = $true }
}
$resolvedWorkbookPath = Join-Path $testRoot 'target.data'
[IO.File]::WriteAllText($resolvedWorkbookPath, 'original')
$excel = [pscustomobject] @{
    Ready=$true; CalculationState=0; Calculation=-4105
    CalculateBeforeSave=$true; EnableEvents=$true
}
$workbook = [pscustomobject] @{
    Saved=$false; ReadOnly=$false; FullName=$resolvedWorkbookPath
    Connections=@(); Worksheets=@(); SaveCalls=0; SaveResult=$true; ThrowOnSave=$false
}
$workbook | Add-Member ScriptMethod Save {
    $this.SaveCalls++
    if ($this.ThrowOnSave) { throw 'Synthetic save failure.' }
    [IO.File]::WriteAllText($this.FullName, "saved $($this.SaveCalls)")
    $this.Saved = $this.SaveResult
}
$script:fixtureWorkbook = $workbook

Write-ExcelSaveDiagnostic $excel $workbook $resolvedWorkbookPath 'ObservedFalse' -ObservedSaved $false -UseObservedSaved
$snapshot = $script:RefreshState.saveDiagnostics[-1]
Assert-SaveTest ($snapshot.Saved -is [bool] -and -not $snapshot.Saved -and $snapshot.SavedType -eq 'System.Boolean') 'Raw Boolean false/type were lost.'
Assert-SaveTest ($workbook.SaveCalls -eq 0 -and -not $workbook.Saved) 'Read-only diagnostic mutated or saved the workbook.'
Write-ExcelSaveDiagnostic $excel $workbook $resolvedWorkbookPath 'ObservedNull' -ObservedSaved $null -UseObservedSaved
Assert-SaveTest ($null -eq $script:RefreshState.saveDiagnostics[-1].Saved -and $null -eq $script:RefreshState.saveDiagnostics[-1].SavedType) 'Unknown Saved state was converted to false/true.'
$excel | Add-Member ScriptProperty Ready { throw 'Synthetic busy property.' } -Force
Write-ExcelSaveDiagnostic $excel $workbook $resolvedWorkbookPath 'UnreadableProbe'
Assert-SaveTest ($null -eq $script:RefreshState.saveDiagnostics[-1].ExcelReady -and $script:RefreshState.saveDiagnostics[-1].Errors.Contains('ExcelReady')) 'Unreadable probe was hidden or prevented recording.'
$excel | Add-Member NoteProperty Ready $true -Force

$workerPath = Join-Path $PSScriptRoot '../CLIENT/DATExx-Whiddon/runner/src/Invoke-AllUnitsExcelWorkbookRefresh.ps1'
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($workerPath, [ref] $null, [ref] $parseErrors)
Assert-SaveTest (@($parseErrors).Count -eq 0) 'Worker has syntax errors.'
$workerTry = @($ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] })[0]
$first = @($workerTry.Body.Statements | Where-Object { $_.Extent.Text -eq "Set-RefreshPhase 'Saving'" })[0]
$last = @($workerTry.Body.Statements | Where-Object { $_.Extent.Text -like 'if (-not $savedAfterSave)*' })[0]
$saveBlock = [scriptblock]::Create((@($workerTry.Body.Statements | Where-Object {
    $_.Extent.StartOffset -ge $first.Extent.StartOffset -and $_.Extent.EndOffset -le $last.Extent.EndOffset
} | ForEach-Object { $_.Extent.Text }) -join "`n"))
foreach ($case in @('Success','False','Unknown','BecomesCleanLater','SaveError')) {
    $script:RefreshState = @{ phase=''; saved=$false; updatedAt=''; detail=$null }
    $script:sleeps = 0
    $script:becomeCleanOnSleep = $case -eq 'BecomesCleanLater'
    $workbook.Saved = $false; $workbook.SaveCalls = 0
    $workbook.SaveResult = if ($case -eq 'Success') { $true } elseif ($case -eq 'Unknown') { $null } else { $false }
    $workbook.ThrowOnSave = $case -eq 'SaveError'
    $failure = ''
    try { . $saveBlock } catch { $failure = $_.Exception.Message }
    $expectedFailure = $case -ne 'Success'
    Assert-SaveTest (($failure.Length -gt 0) -eq $expectedFailure) "$case changed the original save decision."
    Assert-SaveTest ($workbook.SaveCalls -eq 1) "$case retried or skipped Save."
    Assert-SaveTest (-not $script:RefreshState.saved) "$case promoted diagnostic evidence to a completion receipt."
    if ($case -eq 'Success') {
        Assert-SaveTest ($script:sleeps -eq 0) 'Successful save gained an unconditional delay.'
        Assert-SaveTest ($script:RefreshState.saveDiagnostics[-1].Saved -eq $true) 'Successful raw saved status was not recorded.'
    } elseif ($case -eq 'SaveError') {
        Assert-SaveTest ($failure -like '*Synthetic save failure*' -and $script:RefreshState.saveDiagnostics[-1].Point -eq 'SaveThrew') 'Save exception was masked or not diagnosed.'
    } else {
        Assert-SaveTest ($script:sleeps -eq 2 -and $failure -like '*without confirming*') "$case did not retain the failed decision with bounded observations."
        if ($case -eq 'BecomesCleanLater') {
            Assert-SaveTest ($workbook.Saved -and $script:RefreshState.saveDiagnostics[-1].Saved) 'Fixture did not observe a later clean state.'
        }
    }
}
# A getter exception must also remain a failure with its original reason.
$unreadable = [pscustomobject] @{ ReadOnly=$false; FullName=$resolvedWorkbookPath; Connections=@(); Worksheets=@(); SaveCalls=0 }
$unreadable | Add-Member ScriptMethod Save { $this.SaveCalls++ }
$workbook = $unreadable
$failure = ''
try { . $saveBlock } catch { $failure = $_.Exception.Message }
Assert-SaveTest ($failure -like '*property*Saved*cannot be found*' -and $workbook.SaveCalls -eq 1) 'Unreadable Saved property was accepted or masked.'
Assert-SaveTest ($script:RefreshState.saveDiagnostics[-1].Point -eq 'SavedReadFailed') 'Saved getter failure was not diagnosed.'
Write-Host "PASS: $script:checks save diagnostic checks; raw/unknown states, observational probes, unchanged save decisions, bounded failure observations and no Excel."
