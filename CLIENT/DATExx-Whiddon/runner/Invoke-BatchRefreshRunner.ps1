[CmdletBinding()]
param(
    [ValidateSet('Current', 'ParallelInputs-Unit1Priority')] [string] $SequenceProfile = 'Current',
    [switch] $ExportGantt, [string] $ReportRun,
    [switch] $RunAll,
    [string[]] $Units, [string[]] $Batches, [string[]] $Roles, [string[]] $Workbooks,
    [string] $StartAtWorkbook, [int] $StartAtSequence, [int] $EndAtSequence,
    [switch] $IncludeOrg, [switch] $IncludeDependencies,
    [switch] $RefreshSelected, [switch] $ShowPlan, [switch] $ValidateSelectionOnly,
    [switch] $ShowStatus, [string] $ResumeRun,
    [ValidateSet('AfterCurrent', 'AfterBatch', 'Now')] [string] $StopMode,
    [ValidateRange(1, 7)] [int] $MaxParallelBatches = 7,
    [ValidateSet('true', 'false')] [string] $VisibleOverride = 'false',
    [ValidateRange(0.01, 1440)] [double] $TimeoutMinutesOverride = 30
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$interactiveMenu = $PSBoundParameters.Count -eq 0
$holdSuccessfulMenuResult = $false
. (Join-Path $PSScriptRoot 'src/BatchPlanning.ps1')
. (Join-Path $PSScriptRoot 'src/BatchScheduling.ps1')
. (Join-Path $PSScriptRoot 'src/BatchReporting.ps1')
. (Join-Path $PSScriptRoot 'src/RefreshRunGate.ps1')
. (Join-Path $PSScriptRoot 'src/ExcelRefreshSafety.ps1')
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$gate = $null
$runDirectory = $null

function Show-ResolvedBatchPlan {
    param($Plan, $Catalogue)
    Write-Host "Selected: $($Plan.Jobs.Count) files in $(@($Plan.Jobs | Group-Object BatchKey).Count) batches."
    Write-Host "Sequence profile: $($Plan.SequenceProfile)"
    Write-Host "Parallel batches: $MaxParallelBatches (maximum 7); files within each batch are sequential."
    Write-Host "Catalogue fingerprint: $($Catalogue.Fingerprint)"
    if (@($Plan.Jobs | Where-Object Unit -eq 'Org').Count) {
        Write-Host "Organisation consumer Units: $($Plan.OrgUnits -join ', '). Unit selection does not change workbook filters."
    }
    $selectedIds = @($Plan.Jobs | ForEach-Object Id)
    $selectedPaths = @($Plan.Jobs | ForEach-Object Path)
    foreach ($batchKey in @($Plan.Jobs | ForEach-Object BatchKey | Select-Object -Unique)) {
        $batchJobs = @($Plan.Jobs | Where-Object BatchKey -eq $batchKey)
        Write-Host "`n$batchKey [$($batchJobs[0].Batch)]"
        foreach ($job in $batchJobs) {
            Write-Host "  $($job.Id) | $($job.RelativePath)"
            $waits = @($job.Dependencies | Where-Object { $_ -in $selectedIds })
            if ($waits.Count) { Write-Host "    After: $($waits -join ', ')" }
            foreach ($path in $job.Reads) {
                if ($path -notin $selectedPaths) {
                    Write-Host "    Existing input - not refreshed this run: $([IO.Path]::GetRelativePath($Plan.DateRoot, $path))"
                }
            }
        }
    }
    Write-Host "`nTechnical refresh only; business reconciliation is not evaluated."
}

try {
    # Status/stop must remain available even if a catalogue is broken or files vanish.
    $project = Get-Content -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Raw | ConvertFrom-Json
    $logRoot = Resolve-BatchPath $repoRoot $project.batchRunner.logFolder
    $currentPath = Join-Path $logRoot 'current.json'
    $reportRoot = Resolve-BatchPath $repoRoot $(if ($project.batchRunner.PSObject.Properties['reportFolder']) { $project.batchRunner.reportFolder } else { 'docs' })
    if ($ReportRun) {
        if ($ReportRun -notmatch '^[a-zA-Z0-9-]+$' -or $RunAll -or $RefreshSelected -or $ResumeRun -or $StopMode) { throw 'ReportRun is report-only; do not combine it with execution or stop options.' }
        $reportDirectory = Resolve-BatchPath $logRoot $ReportRun
        $reportPlan = Read-BatchJson (Join-Path $reportDirectory 'plan.json')
        $reportState = Read-BatchJson (Join-Path $reportDirectory 'state.json') -AsHashtable
        if ($reportState.Status -eq 'Running' -or @(Get-ChildItem -LiteralPath $reportDirectory -Filter 'state-recovery-*.json').Count) { throw 'An active run or recovery snapshot requires inspection before final reporting.' }
        if (-not $reportPlan.PSObject.Properties['SequenceProfile']) { $reportPlan | Add-Member SequenceProfile 'Current' }
        $null = Write-BatchActualReport $reportPlan $reportState $reportDirectory
        Publish-BatchReports $reportDirectory $reportRoot $reportPlan.SequenceProfile $ReportRun
        exit 0
    }
    if ($interactiveMenu) {
        Write-Host '1. Run all (up to 7 parallel batches)'
        Write-Host '2. Run selected batches    3. Run selected Units    4. Run selected roles'
        Write-Host '5. Run exact files        6. Run legacy range      7. Preview / validate all'
        Write-Host '8. Status                 9. Resume               10. Stop    0. Exit'
        $menuOption = Read-Host 'Option'
        switch ($menuOption) {
            '1' { $RunAll = $true }
            '2' {
                $pickCatalogue = Get-BatchCatalogue $repoRoot $SequenceProfile
                foreach ($line in (Get-BatchPickList $pickCatalogue)) { Write-Host $line }
                $Batches = @(Read-Host 'Batch IDs, comma-separated')
                $Units = @(Read-UnitSelection $pickCatalogue -BlankMeansAll)
                $RefreshSelected = $true
            }
            '3' {
                $pickCatalogue = Get-BatchCatalogue $repoRoot $SequenceProfile
                $Units = @(Read-UnitSelection $pickCatalogue)
                $RunAll = $true
            }
            '4' {
                $Roles = @(Read-Host 'Saved role names, comma-separated')
                $pickCatalogue = Get-BatchCatalogue $repoRoot $SequenceProfile
                $Units = @(Read-UnitSelection $pickCatalogue -BlankMeansAll)
                $RefreshSelected = $true
            }
            '5' { $Workbooks = @(Read-Host 'Exact Date-relative workbook paths, comma-separated'); $RefreshSelected = $true }
            '6' { $StartAtSequence = [int] (Read-Host 'Start global sequence'); $EndAtSequence = [int] (Read-Host 'End global sequence'); $RefreshSelected = $true }
            '7' { $ValidateSelectionOnly = $true }
            '8' { $ShowStatus = $true }
            '9' { $ResumeRun = Read-Host 'Run ID'; $RefreshSelected = $true }
            '10' { $StopMode = Read-Host 'AfterCurrent, AfterBatch or Now'; if ($StopMode -notin @('AfterCurrent', 'AfterBatch', 'Now')) { throw 'Invalid stop mode.' } }
            default { exit 0 }
        }
        $holdSuccessfulMenuResult = $true
        if (($menuOption -eq '2' -and -not (Expand-BatchArguments $Batches)) -or
            ($menuOption -eq '3' -and -not (Expand-BatchArguments $Units)) -or
            ($menuOption -eq '4' -and -not (Expand-BatchArguments $Roles)) -or
            ($menuOption -eq '5' -and -not (Expand-BatchArguments $Workbooks))) { throw 'Empty selection; nothing was started.' }
    }
    if ($RefreshSelected -and -not ($RunAll -or $ResumeRun -or $Units -or $Batches -or $Roles -or $Workbooks -or $StartAtWorkbook -or $StartAtSequence)) { throw 'Specify a selection or use RunAll explicitly.' }
    if ($ShowStatus -or $StopMode) {
        if (-not (Test-Path -LiteralPath $currentPath)) { Write-Host 'No batch run has been recorded.'; exit 0 }
        $current = Read-BatchJson $currentPath
        if ($current.RunId -notmatch '^[a-zA-Z0-9-]+$') { throw 'Invalid current run ID.' }
        $directory = Resolve-BatchPath $logRoot $current.RunId
        if ($StopMode) {
            Write-BatchJson (Join-Path $directory 'stop.json') @{ Mode = $StopMode; Requested = [datetime]::UtcNow.ToString('o') }
            Write-Host "Requested $StopMode for $($current.RunId)."
        } else {
            $state = Read-BatchJson (Join-Path $directory 'state.json') -AsHashtable
            Write-Host "Run $($state.RunId): $($state.Status); maximum observed parallel batches: $($state.MaxObserved)."
            if ($state.ContainsKey('CoordinatorError')) { Write-Host "Coordinator error: $($state.CoordinatorError.Message)" }
            if (@(Get-ChildItem -LiteralPath $directory -Filter 'state-recovery-*.json' -File).Count) {
                Write-Host 'A final recovery snapshot exists: state.json may be stale. Inspect the recovery record before any rerun.'
            }
            if ($state.Status -eq 'Running' -and $state.ContainsKey('Coordinator')) {
                $owner = Get-Process -Id $state.Coordinator.Id -ErrorAction SilentlyContinue
                if ($null -eq $owner -or $owner.StartTime.ToUniversalTime().Ticks -ne $state.Coordinator.Started) {
                    Write-Host 'Coordinator is no longer running. Interrupted jobs require inspection before retry.'
                }
            }
            foreach ($key in ($state.Batches.Keys | Sort-Object)) {
                $batch = $state.Batches[$key]; Write-Host "$key | $($batch.Status) | $($batch.Completed)/$($batch.Total)"
            }
            foreach ($key in ($state.Jobs.Keys | Sort-Object)) {
                if ($state.Jobs[$key].Status -in @('Failed', 'Blocked', 'Stopped', 'Running')) { Write-Host "$key | $($state.Jobs[$key].Status) | $($state.Jobs[$key].Message)" }
            }
            Write-Host "Records: $([IO.Path]::GetRelativePath($repoRoot, $directory))"
        }
        exit 0
    }
    if ($ResumeRun) {
        if ($ResumeRun -notmatch '^[a-zA-Z0-9-]+$') { throw 'Invalid run ID.' }
        $savedProfilePlan = Read-BatchJson (Join-Path (Resolve-BatchPath $logRoot $ResumeRun) 'plan.json')
        $savedProfile = if ($savedProfilePlan.PSObject.Properties['SequenceProfile']) { $savedProfilePlan.SequenceProfile } else { 'Current' }
        if ($PSBoundParameters.ContainsKey('SequenceProfile') -and $SequenceProfile -ne $savedProfile) { throw 'Explicit sequence profile conflicts with the saved run.' }
        $SequenceProfile = $savedProfile
        if ($savedProfilePlan.PSObject.Properties['MaxParallelBatches']) {
            if ($PSBoundParameters.ContainsKey('MaxParallelBatches') -and $MaxParallelBatches -ne $savedProfilePlan.MaxParallelBatches) { throw 'Resume must preserve the saved slot limit.' }
            $MaxParallelBatches = $savedProfilePlan.MaxParallelBatches
        }
    }
    $catalogue = Get-BatchCatalogue $repoRoot $SequenceProfile
    if ($ResumeRun) {
        if ($ResumeRun -notmatch '^[a-zA-Z0-9-]+$') { throw 'Invalid run ID.' }
        if ($RunAll -or $Units -or $Batches -or $Roles -or $Workbooks -or $StartAtWorkbook -or $StartAtSequence -or $EndAtSequence -or $IncludeOrg -or $IncludeDependencies) { throw 'Resume uses its immutable saved selection; do not supply a new selection.' }
        $runDirectory = Resolve-BatchPath $logRoot $ResumeRun
        if (@(Get-ChildItem -LiteralPath $runDirectory -Filter 'state-recovery-*.json' -File).Count) {
            throw 'Final state publication failed. Inspect the recovery record and original workbooks; do not automatically resume stale state.json.'
        }
        $plan = Read-BatchJson (Join-Path $runDirectory 'plan.json')
        $state = Read-BatchJson (Join-Path $runDirectory 'state.json') -AsHashtable
        if (-not $plan.PSObject.Properties['LocationIdentity'] -or
            $plan.Fingerprint -ne $catalogue.Fingerprint -or
            $plan.LocationIdentity -ne (Get-BatchLocationIdentity $catalogue.DateRoot)) { throw 'Configuration, machine or Date root changed: start a fresh run.' }
        # Re-resolve immutable IDs from the current catalogue, not editable saved paths.
        $reconstructed = Select-BatchPlan -Catalogue $catalogue -Workbooks @($plan.Jobs | ForEach-Object Id)
        if (($plan.Jobs | ConvertTo-Json -Depth 30 -Compress) -ne ((ConvertTo-PortableBatchPlan $reconstructed).Jobs | ConvertTo-Json -Depth 30 -Compress)) { throw 'Saved plan differs from the current resolved catalogue.' }
        $plan = $reconstructed
        Reset-BatchResume $plan $state
    } else {
        $plan = Select-BatchPlan -Catalogue $catalogue -RunAll:$RunAll -Units $Units -Batches $Batches -Roles $Roles -Workbooks $Workbooks `
            -StartAtWorkbook $StartAtWorkbook -StartAtSequence $StartAtSequence -EndAtSequence $EndAtSequence -IncludeOrg:$IncludeOrg -IncludeDependencies:$IncludeDependencies
    }
    Show-ResolvedBatchPlan $plan $catalogue
    if ($ExportGantt -and ($ShowPlan -or $ValidateSelectionOnly -or (-not $RunAll -and -not $RefreshSelected))) {
        $previewId = 'preview-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
        $previewDirectory = Resolve-BatchPath $logRoot $previewId
        [void][IO.Directory]::CreateDirectory($previewDirectory)
        Write-BatchJson (Join-Path $previewDirectory 'plan.json') (ConvertTo-PortableBatchPlan $plan)
        $null = Write-BatchForecastReport $plan $logRoot $previewDirectory $MaxParallelBatches
        Publish-BatchReports $previewDirectory $reportRoot $SequenceProfile $previewId
    }
    if ($ShowPlan -and -not $ValidateSelectionOnly) { Write-Host 'Preview only; no file-access validation or Excel startup.'; exit 0 }
    Assert-RefreshRunGateAvailable $catalogue.DateRoot
    $issues = @(Test-BatchPlan $plan -InspectFileUsers)
    if ($issues.Count) {
        foreach ($issue in $issues) { Write-Host "VALIDATION: $issue" }
        if ($ShowPlan -or $ValidateSelectionOnly -or (-not $RunAll -and -not $RefreshSelected)) {
            throw "Validation failed for $($issues.Count) file checks. Excel was not started."
        }
        Write-Host 'Unavailable files will block their jobs and descendants. Independent validated work may continue.'
    } else {
        if ($ShowPlan -or $ValidateSelectionOnly -or (-not $RunAll -and -not $RefreshSelected)) {
            Write-Host 'File access and selection validation passed. Excel has not been opened.'
        } else { Write-Host 'File access and selection validation passed. Starting the selected refresh now.' }
    }
    if ($ShowPlan -or $ValidateSelectionOnly -or (-not $RunAll -and -not $RefreshSelected)) { exit 0 }
    $gate = Enter-RefreshRunGate -DateRoot $catalogue.DateRoot -Exclusive
    if (-not $ResumeRun) {
        $runId = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
        $runDirectory = Resolve-BatchPath $logRoot $runId
        [void] [IO.Directory]::CreateDirectory($runDirectory)
        $state = New-BatchState $plan $runId
        $plan | Add-Member MaxParallelBatches $MaxParallelBatches -Force
        Write-BatchJson (Join-Path $runDirectory 'plan.json') (ConvertTo-PortableBatchPlan $plan)
        # Freeze the forecast before any worker starts.
        $null = Write-BatchForecastReport $plan $logRoot $runDirectory $MaxParallelBatches
    } else {
        # Recheck recovery metadata after acquiring exclusion against other runners.
        Reset-BatchResume $plan $state
        $oldStop = Join-Path $runDirectory 'stop.json'
        if (Test-Path -LiteralPath $oldStop) {
            [IO.File]::Move($oldStop, (Join-Path $runDirectory ("stop-resumed-$([guid]::NewGuid().ToString('N')).json")))
        }
    }
    Write-BatchJson (Join-Path $runDirectory 'state.json') $state
    Write-BatchJson $currentPath @{ RunId = $state.RunId }
    Write-Host "Live log: $([IO.Path]::GetRelativePath($repoRoot, (Join-Path $runDirectory 'coordinator.log')))"
    $exitCode = Invoke-BatchSchedule -Plan $plan -State $state -RunDirectory $runDirectory -MaxParallelBatches $MaxParallelBatches -Visible $VisibleOverride -TimeoutMinutes $TimeoutMinutesOverride
    try {
        $null = Write-BatchActualReport $plan $state $runDirectory
        Publish-BatchReports $runDirectory $reportRoot $SequenceProfile $state.RunId
    } catch {
        Write-Warning "Refresh outcome retained; Gantt reporting failed: $($_.Exception.Message). Regenerate using -ReportRun $($state.RunId)."
        try { Write-BatchJson (Join-Path $runDirectory 'report-error.json') @{ Message = $_.Exception.Message } } catch { Write-Warning 'Could not write report error.' }
    }
    Write-Host "Batch run $($state.RunId): $($state.Status); exit $exitCode. Records: $([IO.Path]::GetRelativePath($repoRoot, $runDirectory))"
    $holdSuccessfulMenuResult = $interactiveMenu -and $exitCode -eq 0
    exit $exitCode
}
catch {
    $originalError = $_
    if ($runDirectory -and (Test-Path -LiteralPath $runDirectory)) {
        $null = Write-BatchCoordinatorError $runDirectory $originalError
    }
    $holdSuccessfulMenuResult = $false
    if ($interactiveMenu) {
        Write-Host ''
        Write-Host 'The runner could not complete this request.' -ForegroundColor Yellow
        Write-Host $originalError.Exception.Message
    } else { Write-Error $originalError -ErrorAction Continue }
    exit 1
}
finally {
    if ($null -ne $gate) { $gate.Dispose() }
    # Release run exclusion before waiting. The CMD launcher already holds failed
    # menus open; scripted calls and explicit menu Exit must never wait for input.
    if ($holdSuccessfulMenuResult) {
        Write-Host ''
        Write-Host 'Press any key to close this window.'
        if ([Console]::IsInputRedirected) { [void] [Console]::ReadLine() }
        else { [void] [Console]::ReadKey($true) }
    }
}
