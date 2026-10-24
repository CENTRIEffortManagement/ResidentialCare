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
    Assert-LauncherTest (($dispatch.Jobs -join ',') -eq 'Unit1/Workers,Unit1/Availability,Unit1/Staff' -and $dispatch.Limit -eq 12) 'exact C1 files and twelve-batch limit reach dispatch'

    $result = Invoke-LauncherFixture -InputLines @('2', 'C2', 'Unit1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed; exit 0') 'successful C2 menu keeps the completion result visible'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest ($dispatch.Jobs.Count -eq 9 -and @($dispatch.Jobs | Where-Object { $_ -notlike 'Unit1/Role:*' }).Count -eq 0) 'C2 still dispatches exactly the nine Unit1 role files'
    $dispatchBefore = [IO.File]::ReadAllText($dispatchPath)
    $result = Invoke-LauncherFixture -InputLines @('2', 'C2,O1,O2,O5,O6', 'Unit1', '0') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Selection cancelled' -and
        $result.Output.Contains('Unit1/Capacity (U4)') -and $result.Output.Contains('Unit1/Effort (U4)')) 'mixed role/reporting menu stops and names the missing consolidation workbooks'
    Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore -and [string]::IsNullOrWhiteSpace($result.ErrorText)) 'rejected menu selection never dispatches and keeps its plain-language error visible'
    $result = Invoke-LauncherFixture -InputLines @('2', 'o1,o2,o5', '') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Selection cancelled' -and $result.Output -notmatch 'Available Unit IDs:') 'organisation-only menu blank choice cancels without a Unit prompt'
    Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'organisation gap never dispatches'
    $result = Invoke-LauncherFixture -InputLines @('2', 'C1,C2,U4,01,02,05,06', 'Unit1', 'S') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Saved intermediate outputs explicitly reused' -and
        $result.Output -match 'will not propagate') 'reported menu selection can explicitly use saved O3/O4 outputs with a visible limitation'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest ($dispatch.Jobs.Count -eq 22 -and 'Org/Cost' -notin $dispatch.Jobs -and 'Org/History' -notin $dispatch.Jobs -and
        'Org/HistoryRead' -notin $dispatch.Jobs) 'saved-input menu dispatches exactly the requested batches'
    $savedCurrent = Get-Content -LiteralPath (Join-Path $fixtureDate 'RunLogs/BatchRefresh/current.json') -Raw | ConvertFrom-Json
    $savedDirectory = Join-Path $fixtureDate "RunLogs/BatchRefresh/$($savedCurrent.RunId)"
    $savedPlan = Get-Content -LiteralPath (Join-Path $savedDirectory 'plan.json') -Raw | ConvertFrom-Json
    Assert-LauncherTest ($savedPlan.UseSavedIntermediates -and $savedPlan.SavedIntermediates.Count -eq 3 -and
        [IO.File]::ReadAllText((Join-Path $savedDirectory 'coordinator.log')).Contains('Saved intermediate outputs explicitly reused') -and
        [IO.File]::ReadAllText((Join-Path $savedDirectory 'planned-notes.md')).Contains('will not propagate')) 'saved-input choice is durable in the plan, log and forecast'
    $result = Invoke-LauncherFixture -Arguments '-ShowStatus'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Saved intermediate outputs explicitly reused') 'status keeps the saved-output limitation visible'
    $result = Invoke-LauncherFixture -Arguments "-ResumeRun $($savedCurrent.RunId) -RefreshSelected"
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Saved intermediate outputs explicitly reused') 'resume preserves the exact saved-output choice without asking again'
    $result = Invoke-LauncherFixture -InputLines @('2', 'o1,o2,o3,o4,o5') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -notmatch 'Available Unit IDs:') 'organisation-only menu selection skips the Unit prompt'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest (($dispatch.Jobs -join ',') -eq 'Org/StaffAll,Org/EffortAll,Org/Outcomes,Org/Inefficiencies,Org/Cost,Org/History,Org/HistoryRead,Org/Reporting') 'organisation-only menu dispatches the selected batches once'
    $result = Invoke-LauncherFixture -InputLines @('2', 'o1,o2,o3,o4,05') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Interpreting batch ID 05 as O5' -and $result.Output -notmatch 'Available Unit IDs:') 'zero-prefixed organisation ID is accepted and explained without a Unit prompt'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest (($dispatch.Jobs -join ',') -eq 'Org/StaffAll,Org/EffortAll,Org/Outcomes,Org/Inefficiencies,Org/Cost,Org/History,Org/HistoryRead,Org/Reporting') 'zero-prefixed O5 selects the intended organisation jobs once'
    $result = Invoke-LauncherFixture -InputLines @('2', '06') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Interpreting batch ID 06 as O6' -and $result.Output -notmatch 'Available Unit IDs:') 'Leave Balance menu selection accepts 06 without a Unit prompt'
    Assert-LauncherTest ($result.Output.Contains('O6 - Leave Balance Stats [LBStats.xlsx, LBReduction.xlsx, TableauConnect-LB.xlsx]')) 'menu lists the Leave Balance batch title and ordered files'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest (($dispatch.Jobs -join ',') -eq 'Org/LeaveBalanceStats,Org/LeaveBalanceReduction,Org/LeaveBalanceReporting') 'Leave Balance menu dispatches the three workbooks once in order'
    $result = Invoke-LauncherFixture -InputLines @('2', 'U4,O1', 'Unit1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Available Unit IDs:') 'mixed batch selection still asks for Units'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest (@($dispatch.Jobs | Where-Object { $_ -like 'Unit2/*' }).Count -eq 0 -and @($dispatch.Jobs | Where-Object { $_ -like 'Org/*' }).Count -eq 2) 'mixed selection scopes Unit work and retains organisation work'
    $result = Invoke-LauncherFixture -InputLines @('1') -ExpectPause -CheckGateReleased
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed; exit 0' -and $result.Output -match 'Run all .*Leave Balance Stats') 'Run all is labelled with its final batch and keeps the completion result visible'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest ($dispatch.Jobs[-1] -eq 'Org/LeaveBalanceReporting' -and @($dispatch.Jobs | Where-Object { $_ -like 'Org/*' }).Count -eq 11) 'menu option 1 dispatches the complete organisation chain and final Leave Balance batch'
    $dispatchBefore = [IO.File]::ReadAllText($dispatchPath)
    $result = Invoke-LauncherFixture -InputLines @('7') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'File access and selection validation passed') 'successful menu validation stays visible'
    $result = Invoke-LauncherFixture -InputLines @('8') -ExpectPause
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Completed' -and $result.Output -notmatch 'Saved intermediate outputs explicitly reused') 'menu status stays visible without a false saved-intermediate warning for a full run'
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
    foreach ($action in @('-RefreshSelected', '-ValidateSelectionOnly', '-ShowPlan')) {
        $runCount = @(Get-ChildItem -LiteralPath (Join-Path $fixtureDate 'RunLogs/BatchRefresh') -Directory).Count
        $result = Invoke-LauncherFixture -Arguments "-Units Unit1 -Batches C2,O1,O2,O5,O6 $action"
        Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.ErrorText -match 'Selection skips intermediate workbooks') "scripted $action rejects the incomplete chain"
        Assert-LauncherTest ([IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore -and
            @(Get-ChildItem -LiteralPath (Join-Path $fixtureDate 'RunLogs/BatchRefresh') -Directory).Count -eq $runCount) 'gap rejection precedes dispatch and run-directory creation'
    }
    foreach ($action in @('-ValidateSelectionOnly', '-ShowPlan')) {
        $result = Invoke-LauncherFixture -Arguments "-Units Unit1 -Batches C2,O1,O2,O5,O6 -UseSavedIntermediates $action"
        Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'Saved intermediate outputs explicitly reused' -and
            $result.Output -notmatch 'Saved intermediate outputs: S or 0|Press any key' -and
            [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) "explicit saved-output $action stays read-only and never prompts"
    }
    $result = Invoke-LauncherFixture -Arguments '-Units Unit1 -Batches C2,O1,O2,O5,O6 -UseSavedIntermediates -RefreshSelected'
    Assert-LauncherTest ($result.ExitCode -eq 0 -and $result.Output -match 'will not propagate' -and $result.Output -notmatch 'Press any key') 'scripted saved-output run requires no menu confirmation'
    $dispatch = Get-Content -LiteralPath $dispatchPath -Raw | ConvertFrom-Json
    Assert-LauncherTest ($dispatch.Jobs.Count -eq 17 -and 'Unit1/Capacity' -notin $dispatch.Jobs -and 'Unit1/Effort' -notin $dispatch.Jobs) 'scripted saved-output run keeps the exact requested scope'
    $dispatchBefore = [IO.File]::ReadAllText($dispatchPath)
    # Reconstruct an old incomplete plan with a current fingerprint to ensure the
    # selection check, independently of fingerprint invalidation, protects resume.
    . (Join-Path $fixtureRunner 'src/BatchPlanning.ps1')
    . (Join-Path $fixtureRunner 'src/BatchScheduling.ps1')
    . (Join-Path $fixtureRunner 'src/BatchReporting.ps1')
    $resumeCatalogue = Get-BatchCatalogue $fixtureRepo
    $incompletePlan = [pscustomobject]@{
        SchemaVersion = 1; Fingerprint = $resumeCatalogue.Fingerprint; DateRoot = $resumeCatalogue.DateRoot
        SequenceProfile = $resumeCatalogue.SequenceProfile; Roles = $resumeCatalogue.Roles; OrgUnits = $resumeCatalogue.OrgUnits
        Jobs = @($resumeCatalogue.Jobs | Where-Object { $_.Id -like 'Unit1/Role:*' -or $_.Batch -eq 'O1' })
    }
    $gapResumeDirectory = Join-Path $fixtureDate 'RunLogs/BatchRefresh/gap-resume'
    [void] [IO.Directory]::CreateDirectory($gapResumeDirectory)
    Write-BatchJson (Join-Path $gapResumeDirectory 'plan.json') (ConvertTo-PortableBatchPlan $incompletePlan)
    Write-BatchJson (Join-Path $gapResumeDirectory 'state.json') (New-BatchState $incompletePlan 'gap-resume')
    $result = Invoke-LauncherFixture -Arguments '-ResumeRun gap-resume -RefreshSelected'
    Assert-LauncherTest ($result.ExitCode -eq 1 -and $result.ErrorText -match 'Selection skips intermediate workbooks' -and
        [IO.File]::ReadAllText($dispatchPath) -eq $dispatchBefore) 'resume reconstructs and rejects an incomplete saved selection before dispatch'
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
