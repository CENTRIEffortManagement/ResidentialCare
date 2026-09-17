# Regression for incidental repository readers and delayed Excel identity.
# Only synthetic .data files and fake objects are used; Excel is never started.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-file-access-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
$script:checks = 0
$script:messages = [Collections.Generic.List[string]]::new()
$script:stopRequested = $false
function Assert-AccessTest {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}
function Assert-AccessError {
    param([scriptblock] $Action, [string] $Pattern)
    $message = ''
    try { & $Action } catch { $message = $_.Exception.Message }
    Assert-AccessTest ($message -like $Pattern) "Expected '$Pattern', got '$message'"
}
function Write-Log { param($Message, $Level) $script:messages.Add($Message) }
function Write-CurrentStatus { param($State, $Message) }
function Assert-NotStopNowRequested { if ($script:stopRequested) { throw 'Stop requested.' } }
$script:RefreshState = @{ phase = 'WaitingToSave'; detail = $null; saved = $false; updatedAt = '' }
$WorkerStatePath = Join-Path $testRoot 'receipt.json'
$target = Join-Path $testRoot 'target.data'
[IO.File]::WriteAllText($target, 'unchanged target')
try {
    $script:queries = 0
    function Get-RefreshFileUsers {
        param($Path)
        $script:queries++
        [pscustomobject]@{ AppName = 'Owned Excel'; Process = [pscustomobject]@{Id=42} }
        if ($script:queries -le 2) { [pscustomobject]@{ AppName='Git for Windows'; Process=[pscustomobject]@{Id=99} } }
    }
    Assert-NoExternalFileUsers $target -AllowedProcessId 42 -WaitSeconds 1 -PollMilliseconds 10
    Assert-AccessTest ($script:queries -eq 3) 'temporary Git use is rechecked until absent, not whitelisted'
    Assert-AccessTest ([IO.File]::ReadAllText($target) -eq 'unchanged target') 'access wait never rewrites the target'
    $script:queries = 0
    Wait-RefreshFileAccess $target -AllowedProcessId 42 -WaitSeconds 3
    Assert-AccessTest ($script:queries -eq 3 -and ($script:messages -join ' ') -match 'Git for Windows') 'worker waits with holder diagnostics, then continues'
    Assert-AccessTest ($null -eq $script:RefreshState.detail) 'resolved access wait clears its stale warning'

    function Get-RefreshFileUsers { param($Path) [pscustomobject]@{AppName='Persistent holder';Process=[pscustomobject]@{Id=99}} }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    Assert-AccessError { Assert-NoExternalFileUsers $target -WaitSeconds .06 -PollMilliseconds 10 } '*Persistent holder*after waiting*'
    Assert-AccessTest ($timer.Elapsed.TotalSeconds -lt 1) 'persistent contention has a bounded wait'
    Assert-AccessError { Assert-NoExternalFileUsers $target } '*Persistent holder*after waiting 0*'
    $script:stopRequested = $true
    Assert-AccessError { Wait-RefreshFileAccess $target -WaitSeconds 60 } '*Stop requested*'
    $script:stopRequested = $false
    function Get-RefreshFileUsers { param($Path) throw 'File-user inspection unavailable.' }
    Assert-AccessError { Wait-RefreshFileAccess $target } '*File-user inspection unavailable*'

    # A holder disappearing is not permission to overwrite someone else's edit.
    function Get-RefreshFileUsers { param($Path) [IO.File]::WriteAllText($Path, 'external change during access wait') }
    Assert-AccessError { Wait-RefreshFileAccess $target } '*Workbook changed while waiting*'

    $workbook = [pscustomobject]@{ Reads=0; ReadOnly=$false; Expected=$target }
    $workbook | Add-Member ScriptProperty FullName { $this.Reads++; if ($this.Reads -lt 3) { return '' }; return $this.Expected }
    Wait-RefreshWorkbookIdentity $workbook $target -WaitSeconds 1 -PollMilliseconds 10
    Assert-AccessTest ($workbook.Reads -eq 3) 'temporarily empty FullName waits for exact identity before continuing'
    $workbook = [pscustomobject]@{ Reads=0; ReadOnly=$false; Expected=$target }
    $workbook | Add-Member ScriptProperty FullName { $this.Reads++; if ($this.Reads -eq 1) { throw 'Excel busy' }; return $this.Expected }
    Wait-RefreshWorkbookIdentity $workbook $target -WaitSeconds 1 -PollMilliseconds 10
    Assert-AccessTest ($workbook.Reads -eq 2) 'transient identity-read failure is retried'
    Assert-AccessError { Wait-RefreshWorkbookIdentity ([pscustomobject]@{FullName=$target;ReadOnly=$true}) $target -WaitSeconds 0 } '*read-only*'
    Assert-AccessError { Wait-RefreshWorkbookIdentity ([pscustomobject]@{FullName=(Join-Path $testRoot 'wrong.data');ReadOnly=$false}) $target -WaitSeconds 0 } '*identity mismatch*'
    Assert-AccessError { Wait-RefreshWorkbookIdentity ([pscustomobject]@{FullName='target.data';ReadOnly=$false}) $target -WaitSeconds 0 } '*identity mismatch*'
    Assert-AccessError { Wait-RefreshWorkbookIdentity ([pscustomobject]@{FullName='';ReadOnly=$false}) $target -WaitSeconds .04 -PollMilliseconds 10 } '*did not confirm*No refresh or save was permitted*'
    Assert-AccessError { Wait-RefreshWorkbookIdentity $null $target -WaitSeconds 0 } '*did not confirm*'
    $script:stopRequested = $true
    Assert-AccessError { Wait-RefreshWorkbookIdentity ([pscustomobject]@{FullName='';ReadOnly=$false}) $target } '*Stop requested*'
    $script:stopRequested = $false

    # Verify both linear and expanded-worker entry points use identical guards.
    foreach ($file in @('CLIENT/DATExx-Whiddon/runner/src/Invoke-AllUnitsExcelWorkbookRefresh.ps1', 'CLIENT/DATExx-Whiddon/UNITS/runner/src/Invoke-ExcelWorkbookRefresh.ps1')) {
        $text = [IO.File]::ReadAllText((Join-Path $repoRoot $file))
        Assert-AccessTest (($text -split 'Wait-RefreshFileAccess').Count -eq 3) "$file waits before opening and before saving"
        Assert-AccessTest (($text -split 'Wait-RefreshWorkbookIdentity').Count -eq 3) "$file verifies exact identity after opening and after saving"
        Assert-AccessTest ($text.IndexOf("Set-RefreshPhase 'WaitingToSave'") -lt $text.IndexOf('$workbook.Save()')) "$file does not save before checking users"
        $parseErrors = $null
        [void] [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $file), [ref]$null, [ref]$parseErrors)
        Assert-AccessTest ($parseErrors.Count -eq 0) "$file parses cleanly"
    }
    Write-Host "PASS: $script:checks file-access/identity checks; no Excel opened, refreshed or saved."
} finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-file-access-*') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
