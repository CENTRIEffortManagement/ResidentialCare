# Regression for menu results closing before they can be read. The launcher
# runs only in a disposable fixture with mock validation/dispatch and no Excel worker.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-launcher-' + [guid]::NewGuid().ToString('N'))
$fixtureRepo = Join-Path $testRoot 'repo'
$fixtureDate = Join-Path $fixtureRepo 'CLIENT/DATExx-Whiddon'
$fixtureRunner = Join-Path $fixtureDate 'runner'
$beforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$script:checks = 0

function Assert-LauncherTest {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}

function Invoke-LauncherFixture {
    param([string] $Arguments = '', [string[]] $InputLines = @(), [switch] $ExpectPause, [switch] $CheckGateReleased)
    $info = [Diagnostics.ProcessStartInfo]::new($env:ComSpec)
    $launcher = Join-Path $fixtureDate 'Run-BatchRefreshRunner.cmd'
    $info.Arguments = '/d /s /c ""{0}" {1}"' -f $launcher, $Arguments
    $info.WorkingDirectory = $fixtureRepo
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    $info.RedirectStandardInput = $true; $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    $stderr = $process.StandardError.ReadToEndAsync()
    $lines = [Collections.Generic.List[string]]::new()
    try {
        foreach ($line in $InputLines) { $process.StandardInput.WriteLine($line) }
        $process.StandardInput.Flush()
        if ($ExpectPause) {
            $timer = [Diagnostics.Stopwatch]::StartNew()
            $paused = $false
            while ($timer.Elapsed.TotalSeconds -lt 10) {
                $lineTask = $process.StandardOutput.ReadLineAsync()
                while (-not $lineTask.IsCompleted -and $timer.Elapsed.TotalSeconds -lt 10) { Start-Sleep -Milliseconds 25 }
                if (-not $lineTask.IsCompleted) { break }
                $line = $lineTask.GetAwaiter().GetResult()
                if ($null -eq $line) { break }
                $lines.Add($line)
                if ($line -eq 'Press any key to close this window.') { $paused = $true; break }
            }
            Assert-LauncherTest $paused "launcher displayed the hold-open instruction: $($lines -join ' | ')"
            Start-Sleep -Milliseconds 200
            Assert-LauncherTest (-not $process.HasExited) 'menu process is still waiting after displaying its result'
            if ($CheckGateReleased) {
                . (Join-Path $fixtureRunner 'src/RefreshRunGate.ps1')
                Assert-RefreshRunGateAvailable $fixtureDate
                Assert-LauncherTest $true 'completed menu released run exclusion before waiting for a key'
            }
            $process.StandardInput.WriteLine('x')
            $process.StandardInput.Flush()
        }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        Assert-LauncherTest ($process.WaitForExit(10000)) 'launcher returned instead of hanging'
        $lines.Add($stdout.GetAwaiter().GetResult())
        return @{ ExitCode = $process.ExitCode; Output = ($lines -join "`n"); ErrorText = $stderr.GetAwaiter().GetResult() }
    }
    finally {
        # Only the disposable process tree created above can be terminated.
        if (-not $process.HasExited) { $process.Kill($true); $process.WaitForExit() }
        $process.Dispose()
    }
}

try {
    [void] [IO.Directory]::CreateDirectory($fixtureRunner)
    foreach ($unit in @('Unit1', 'Unit2')) { [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureDate "UNITS/$unit")) }
    Copy-Item -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Destination $fixtureRepo
    Copy-Item -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd') -Destination $fixtureDate
    $sourceRunner = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner'
    foreach ($file in (Get-ChildItem -LiteralPath $sourceRunner -Recurse -File | Where-Object Extension -in @('.ps1', '.psd1'))) {
        $target = Join-Path $fixtureRunner ([IO.Path]::GetRelativePath($sourceRunner, $file.FullName))
        [void] [IO.Directory]::CreateDirectory((Split-Path $target -Parent))
        Copy-Item -LiteralPath $file.FullName -Destination $target
    }
    [IO.File]::WriteAllText((Join-Path $fixtureRunner 'ResidentialCare-BatchApproval.psd1'), @'
@{
    Approved = $false
    Evidence = ''
    CatalogueFingerprint = ''
    OrganisationScopeConfirmed = $false
    OrganisationUnits = @('Unit1', 'Unit2')
    AdditionalInputs = @{}
    ExclusiveJobs = @()
}
'@)
    # Even an unexpected dispatch cannot reach an Excel engine in this fixture.
    $stub = "throw 'Excel execution is forbidden in launcher regression tests.'"
    foreach ($name in @('Invoke-AllUnitsExcelWorkbookRefresh.ps1', 'Invoke-OrgExcelWorkbookRefresh.ps1', 'Invoke-BatchWorkbook.ps1')) {
        [IO.File]::WriteAllText((Join-Path $fixtureRunner "src/$name"), $stub)
    }
    # Only the fixture overrides validation and dispatch. Production safety checks
    # stay intact and have their own tests. Legacy false approval fields above must
    # not stop an explicit run after the user-requested gate removal.
    $fixtureOverrides = @'

function Test-BatchPlan {
    param($Plan, [switch] $InspectFileUsers)
    if (Test-Path -LiteralPath (Join-Path $Plan.DateRoot 'fixture-missing-input')) {
        return @('Missing input in launcher fixture.')
    }
    return @()
}
function Invoke-BatchSchedule {
    param($Plan, $State, $RunDirectory, $MaxParallelBatches, $Visible, $TimeoutMinutes)
    Write-BatchJson (Join-Path $Plan.DateRoot 'fixture-dispatch.json') @{ Jobs = @($Plan.Jobs.Id); Limit = $MaxParallelBatches }
    if (Test-Path -LiteralPath (Join-Path $Plan.DateRoot 'fixture-fail-run')) {
        throw 'Simulated refresh failure; no workbook was opened.'
    }
    foreach ($record in $State.Jobs.Values) { $record.Status = 'Completed'; $record.ExitCode = 0 }
    $State.Status = 'Completed'; $State.ExitCode = 0
    if (Test-Path -LiteralPath (Join-Path $Plan.DateRoot 'fixture-partial-run')) {
        $State.Jobs[$Plan.Jobs[0].Id].Status = 'Failed'
        $State.Jobs[$Plan.Jobs[0].Id].ExitCode = 1
        $State.Status = 'PartialOrFailed'; $State.ExitCode = 1
    }
    Update-BatchSummary $Plan $State
    Write-BatchJson (Join-Path $RunDirectory 'state.json') $State
    return $State.ExitCode
}
'@
    [IO.File]::AppendAllText((Join-Path $fixtureRunner 'src/BatchScheduling.ps1'), $fixtureOverrides)
    $dispatchPath = Join-Path $fixtureDate 'fixture-dispatch.json'
    $result = Invoke-LauncherFixture -InputLines @('2', 'C1', 'Unit1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0) 'C1 menu reaches mock execution without separate approval'
    Assert-LauncherTest ($result.Output -match 'Starting the selected refresh now' -and $result.Output -match 'Live log:') 'live dispatch announces refresh and its readable log instead of claiming Excel was not opened'
    Assert-LauncherTest ($result.Output -notmatch 'REFRESH NOT STARTED|LIVE REFRESH GATED|The runner stopped') 'successful C1 menu has no stale approval block or failure message'
    Assert-LauncherTest ([string]::IsNullOrWhiteSpace($result.ErrorText)) 'C1 menu has no red exception/traceback'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest (($dispatch.Jobs -join ',') -eq 'Unit1/Workers,Unit1/Availability,Unit1/Staff' -and $dispatch.Limit -eq 7) 'exact C1 files and seven-batch limit reach dispatch'

    $result = Invoke-LauncherFixture -InputLines @('2', 'C2', 'Unit1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed; exit 0') 'successful C2 menu keeps the completion result visible'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest ($dispatch.Jobs.Count -eq 9 -and @($dispatch.Jobs | Where-Object { $_ -notlike 'Unit1/Role:*' }).Count -eq 0) 'C2 still dispatches exactly the nine Unit1 role files'
    $result = Invoke-LauncherFixture -InputLines @('1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed; exit 0') 'Run all keeps the completion result visible'
    $dispatchBefore = [IO.File]::ReadAllText($dispatchPath)
    $result = Invoke-LauncherFixture -InputLines @('7') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'File access and selection validation passed') 'successful menu validation stays visible'
    $result = Invoke-LauncherFixture -InputLines @('8') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed') 'menu status stays visible'
    Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'validation and status menus never dispatch'

    $failPath = Join-Path $fixtureDate 'fixture-fail-run'
    [IO.File]::WriteAllText($failPath, 'fixture failure')
    $result = Invoke-LauncherFixture -InputLines @('2', 'C1', 'Unit1') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.Output -match 'Simulated refresh failure') 'failed C1 menu remains visible and preserves its exit code'
    $errorFiles = @(Get-ChildItem (Join-Path $fixtureDate 'RunLogs/BatchRefresh') -Recurse -Filter 'coordinator-error-*.json')
    Assert-LauncherTest ($errorFiles.Count -eq 1 -and (Get-Content $errorFiles[0].FullName -Raw) -match 'Simulated refresh failure') 'startup/dispatch exception has a durable diagnostic record'
    [IO.File]::Delete($failPath)

    $partialPath = Join-Path $fixtureDate 'fixture-partial-run'
    [IO.File]::WriteAllText($partialPath, 'fixture partial run')
    $result = Invoke-LauncherFixture -InputLines @('2', 'C2', 'Unit1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.Output -match 'PartialOrFailed; exit 1') 'partial completion waits once and preserves its failure exit code'
    [IO.File]::Delete($partialPath)

    $result = Invoke-LauncherFixture -InputLines @('2', 'C1??', 'Unit1') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.Output -match 'Unknown or empty batch selection') 'invalid menu input remains visible'
    Assert-LauncherTest ([string]::IsNullOrWhiteSpace($result.ErrorText)) 'invalid menu input uses plain-language error output'

    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C1 -RefreshSelected'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -notmatch 'REFRESH NOT STARTED|Press any key') 'scripted C1 run reaches mock dispatch without pausing'
    $result = Invoke-LauncherFixture -Arguments '-Batches O1 -RefreshSelected'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Organisation consumer Units: Unit1, Unit2') 'organisation run uses configured scope without a separate confirmation gate'
    $dispatchBefore = [IO.File]::ReadAllText($dispatchPath)
    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C1'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'selection without an explicit run command still never dispatches'
    $missingPath = Join-Path $fixtureDate 'fixture-missing-input'
    [IO.File]::WriteAllText($missingPath, 'fixture validation failure')
    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C1 -ValidateSelectionOnly'
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.Output -match 'VALIDATION:' -and $result.Output -notmatch 'Press any key') 'validation failures remain nonzero and do not pause scripted calls'
    Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'validation never dispatches'
    [IO.File]::Delete($missingPath)
    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C1 -RunAll -ValidateSelectionOnly'
    Assert-LauncherTest ($result.ExitCode -eq 1 -and [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'invalid selectors still fail without dispatch'
    $result = Invoke-LauncherFixture -Arguments '-RunAll -ValidateSelectionOnly'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'validation-only wins over RunAll'
    $current = Get-Content (Join-Path $fixtureDate 'RunLogs/BatchRefresh/current.json') -Raw | ConvertFrom-Json
    $recoveryPath = Join-Path $fixtureDate "RunLogs/BatchRefresh/$($current.RunId)/state-recovery-fixture.json"
    [IO.File]::WriteAllText($recoveryPath, '{"Status":"PartialOrFailed"}')
    $result = Invoke-LauncherFixture -Arguments '-ShowStatus'
    Assert-LauncherTest ($result.Output -match 'state.json may be stale') 'status warns about unpublished terminal recovery state'
    $result = Invoke-LauncherFixture -Arguments "-ResumeRun $($current.RunId) -RefreshSelected"
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.ErrorText -match 'Final state publication failed' -and [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'resume rejects stale state when final recovery snapshot exists'
    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C1 -ShowPlan'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Preview only' -and $result.Output -notmatch 'Press any key') 'scripted preview returns successfully without pausing'
    $result = Invoke-LauncherFixture -InputLines @('0')
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -notmatch 'Press any key') 'explicit menu exit does not pause'
    Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'preview and menu Exit never dispatch'
    Assert-LauncherTest (-not (Import-PowerShellDataFile (Join-Path $fixtureRunner 'ResidentialCare-BatchApproval.psd1')).Approved) 'legacy false approval fields need not be falsified to permit explicit runs'
    Assert-LauncherTest (@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object Id -notin $beforeExcel).Count -eq 0) 'no Excel process was started'
    Write-Host "PASS: $script:checks launcher checks; successful and failed menus stay open, scripted calls and Exit do not pause, validation does not run. No Excel opened."
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]] '\/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-launcher-*') { throw 'Unsafe launcher fixture cleanup target.' }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
