param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchScheduling.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchReporting.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-batch-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
$catalogueRoot = Join-Path $testRoot 'selection-config'
$catalogueDate = Join-Path $catalogueRoot 'CLIENT/DATExx-Whiddon'
[void] [IO.Directory]::CreateDirectory((Join-Path $catalogueDate 'runner'))
# Selection tests need the legacy Units even when production no longer has them.
# Copy configuration only; never create or copy production workbooks.
Copy-Item -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Destination $catalogueRoot
Get-ChildItem -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner') -Filter '*.psd1' -File |
    Copy-Item -Destination (Join-Path $catalogueDate 'runner')
foreach ($unit in @('Unit1','Unit2','BD','TE','JH-RY')) {
    [void] [IO.Directory]::CreateDirectory((Join-Path $catalogueDate "UNITS/$unit"))
}
$script:assertions = 0
$beforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
function Assert-True {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }; $script:assertions++
}
function Assert-Throws {
    param([scriptblock] $Action, [string] $Message)
    $threw = $false; try { & $Action | Out-Null } catch { $threw = $true }
    Assert-True $threw $Message
}
function New-FakeJob {
    param([string] $Id, [string] $Batch = $Id, [string[]] $Depends = @(), [string[]] $Reads = @())
    $path = Join-Path $testRoot (($Id -replace '[^A-Za-z0-9_.-]', '-') + '.data')
    [IO.File]::WriteAllText($path, 'original')
    return [pscustomobject]@{ Id = $Id; BatchKey = $Batch; Batch = $Batch; Path = $path; Reads = $Reads; Dependencies = $Depends; Exclusive = $false }
}
function Invoke-FakeRun {
    param([object[]] $Jobs, [int] $Max = 12, [string[]] $Fail = @(), [string] $Stop = '',
        [string[]] $Slow = @(), [switch] $Uncertain, [int] $FailureCode = 1)
    $directory = Join-Path $testRoot ([guid]::NewGuid().ToString('N'))
    [void] [IO.Directory]::CreateDirectory($directory)
    $plan = [pscustomobject]@{ SchemaVersion = 1; Fingerprint = 'fixture'; DateRoot = $testRoot; Jobs = $Jobs }
    $state = New-BatchState $plan 'fixture'
    $sim = @{ Active = @{}; Events = [Collections.Generic.List[string]]::new(); Max = 0; RequestedStop = $false }
    $start = {
        param($job, $jobDirectory, $options)
        Assert-True (-not (Test-BatchConflict $job @($sim.Active.Values))) 'scheduler admitted conflicting files or two files in a batch'
        foreach ($path in $job.Reads) { Assert-True ($options.InputStamps[$path] -eq (Get-BatchFileStamp $path)) 'dispatch carries the frozen input baseline to the worker' }
        $sim.Active[$job.Id] = $job; $sim.Max = [Math]::Max($sim.Max, $sim.Active.Count)
        $sim.Events.Add("start:$($job.Id)")
        if ($Stop -and -not $sim.RequestedStop) {
            Write-BatchJson (Join-Path $directory 'stop.json') @{ Mode = $Stop }; $sim.RequestedStop = $true
        }
        return @{ Job = $job; Remaining = $(if ($job.Id -in $Slow) { 10 } else { 2 }); Directory = $jobDirectory }
    }.GetNewClosure()
    $poll = {
        param($handle)
        $job = $handle.Job
        $stopped = Test-Path -LiteralPath (Join-Path $handle.Directory 'stop.txt')
        $handle.Remaining--
        if ($handle.Remaining -gt 0 -and -not $stopped) { return $null }
        $sim.Active.Remove($job.Id); $sim.Events.Add("end:$($job.Id)")
        $code = if ($stopped) { 3 } elseif ($job.Id -in $Fail) { $FailureCode } else { 0 }
        if ($code -eq 0) { [IO.File]::WriteAllText($job.Path, ('saved-' + [guid]::NewGuid().ToString('N'))) }
        return @{ ExitCode = $code; NeedsInspection = ($code -ne 0 -and [bool] $Uncertain); Message = "Fixture exit $code" }
    }.GetNewClosure()
    $code = Invoke-BatchSchedule $plan $state $directory -MaxParallelBatches $Max -StartWorker $start -PollWorker $poll -ValidateJob { param($job) Assert-BatchFileAccessible $job.Path -OutputFile } -PollMilliseconds 0
    return @{ State = $state; Plan = $plan; Events = $sim.Events; Max = $sim.Max; Code = $code; Directory = $directory }
}
try {
    $catalogue = Get-BatchCatalogue $catalogueRoot
    $defaultRun = Select-BatchPlan $catalogue -RunAll
    Assert-True ($defaultRun.Jobs.Count -eq 90) 'Run all selects current facilities including CS without retired Workers jobs, one shared demand master and eleven organisation workbooks'
    Assert-True ((Select-BatchPlan $catalogue).Jobs.Count -eq 90) 'unqualified preview matches Run all scope'
    Assert-True (@($defaultRun.Jobs | Group-Object BatchKey).Count -eq 43) 'default run resolves three Unit batch sets including CS, one shared input batch and six organisation batches'
    Assert-True (@($defaultRun.Jobs | Where-Object Unit -in @('Unit1', 'Unit2')).Count -eq 0 -and @($defaultRun.Jobs | Where-Object Unit -eq 'Org').Count -eq 11) 'Run all excludes legacy Units and includes the complete organisation release chain'
    Assert-True ($defaultRun.Jobs[-1].Id -eq 'Org/LeaveBalanceReporting' -and (Split-Path $defaultRun.Jobs[-1].Path -Leaf) -eq 'TableauConnect-LB.xlsx') 'Run all ends at the Leave Balance reporting workbook'
    Assert-True (($catalogue.RunAllUnits -join ',') -eq 'BD,TE,JH-RY' -and $catalogue.RunAllIncludeOrg) 'saved default scope includes organisation work'
    $all = Select-BatchPlan $catalogue -RunAll -Units Unit1,Unit2
    Assert-True ($all.Jobs.Count -eq 50 -and @($all.Jobs | Group-Object BatchKey).Count -eq 22) 'explicit legacy Unit selection remains available without organisation work'
    Assert-True ($all.Jobs[0].Id -eq 'Unit1/AllocationInput') 'numeric Unit and catalogue order'
    Assert-True (@($all.Jobs | Where-Object RelativePath -match 'TestRole').Count -eq 0) 'test roles excluded'
    Assert-True (-not $catalogue.Settings.ContainsKey('Approved')) 'operational settings do not require approval metadata'
    Assert-True (($catalogue.OrgUnits -join ',') -eq 'BD,TE,JH-RY') 'organisation consumer scope matches the default facility Units'
    Assert-True (($catalogue.Units -join ',') -eq 'Unit1,Unit2,BD,TE,JH-RY') 'BD, TE and JH-RY are added after legacy Units'
    Assert-Throws { Select-BatchPlan $catalogue -RunAll -Units Unit1 -IncludeOrg } 'legacy Unit cannot be combined with current organisation jobs'
    $orgStaffAll = @($defaultRun.Jobs | Where-Object Id -eq 'Org/StaffAll')[0]
    $orgEffortAll = @($defaultRun.Jobs | Where-Object Id -eq 'Org/EffortAll')[0]
    Assert-True (($orgStaffAll.Dependencies -join ',') -eq 'BD/Staff,TE/Staff,JH-RY/Staff') 'organisation staff waits for all default facility staff outputs'
    Assert-True ('BD/Effort' -in $orgEffortAll.Dependencies -and 'TE/Effort' -in $orgEffortAll.Dependencies -and
        'JH-RY/Effort' -in $orgEffortAll.Dependencies) 'organisation effort waits for all default facility Effort outputs'
    $unitPicks = @(Get-UnitPickList $catalogue)
    Assert-True ('Unit1' -in $unitPicks -and 'Unit2' -in $unitPicks -and 'BD' -in $unitPicks -and 'TE' -in $unitPicks -and 'JH-RY' -in $unitPicks) 'Unit picker lists legacy and additional Units'
    Assert-True ('All - all listed Units' -in $unitPicks) 'Unit picker offers All'
    Assert-True ('Blank = cancel this selection; nothing will run.' -in $unitPicks) 'dedicated Unit selection explains empty-input cancellation'
    Assert-True ('Blank = all listed Units.' -in @(Get-UnitPickList $catalogue -BlankMeansAll)) 'batch and role Unit qualifiers explain blank means all'
    $pickList = @(Get-BatchPickList $catalogue)
    Assert-True ('D1 - Demand transformation [Demand-MasterRoster Manual Read.xlsx, 2-DemandExtract.xlsx]' -in $pickList) 'batch choices show ID, title and ordered comma-separated filenames'
    Assert-True (@($pickList | Where-Object { $_ -like 'D1 -*' }).Count -eq 1) 'shared Unit batch choices are not duplicated'
    Assert-True ('C2.2 - Role capacity - AINC4 [CapacityDistrib(A.1)-shifts.xlsx, CapacityDistrib(A.2)-shifts.xlsx, CapacityDistrib(B)-shifts.xlsx]' -in $pickList) 'legacy AINC4 role has a dedicated batch choice'
    Assert-True ('C2.4 - Role capacity - EN [CapacityDistrib(A.1)-shifts.xlsx, CapacityDistrib(A.2)-shifts.xlsx, CapacityDistrib(B)-shifts.xlsx]' -in $pickList) 'EN has a dedicated batch choice'
    Assert-True ('C2.5 - Role capacity - CS [CapacityDistrib(A.1)-shifts.xlsx, CapacityDistrib(A.2)-shifts.xlsx, CapacityDistrib(B)-shifts.xlsx]' -in $pickList) 'CS has a dedicated batch choice'
    Assert-True ([array]::IndexOf($pickList, ($pickList | Where-Object { $_ -like 'C2.5 -*' })) -lt [array]::IndexOf($pickList, ($pickList | Where-Object { $_ -like 'U4 -*' }))) 'all C2 role choices are listed before U4'
    Assert-True (@($pickList | Where-Object { $_ -like 'C2 - All enabled*' }).Count -eq 1) 'all-role shorthand is offered once'
    Assert-True ('O1 - Cross-unit assembly [StafMasterList-All.xlsx, Effort-All.xlsx]' -in $pickList) 'organisation choices include files'
    Assert-True ('O6 - Leave Balance Stats [LBStats.xlsx, LBReduction.xlsx, TableauConnect-LB.xlsx]' -in $pickList) 'Leave Balance batch picker shows its title and ordered files'
    Assert-True ([array]::IndexOf($pickList, 'O6 - Leave Balance Stats [LBStats.xlsx, LBReduction.xlsx, TableauConnect-LB.xlsx]') -gt [array]::IndexOf($pickList, 'O5 - Reporting release [Tableau Connection.xlsx]')) 'Leave Balance is the final organisation batch choice'
    Assert-True (-not (Test-BatchSelectionNeedsUnits $catalogue @('O6'))) 'Leave Balance batch runs once without a Unit prompt'
    $leaveBalance = Select-BatchPlan $catalogue -Batches O6
    Assert-True (($leaveBalance.Jobs.Id -join ',') -eq 'Org/LeaveBalanceStats,Org/LeaveBalanceReduction,Org/LeaveBalanceReporting' -and @($leaveBalance.Jobs | Group-Object BatchKey).Count -eq 1) 'Leave Balance selection contains exactly the three ordered workbooks once'
    $leaveBalanceStats = $leaveBalance.Jobs[0]
    $leaveBalanceReduction = $leaveBalance.Jobs[1]
    $leaveBalanceReporting = $leaveBalance.Jobs[2]
    Assert-True (($leaveBalanceStats.Dependencies -join ',') -eq 'Org/Reporting' -and ($leaveBalanceReduction.Dependencies -join ',') -eq 'Org/LeaveBalanceStats' -and ($leaveBalanceReporting.Dependencies -join ',') -eq 'Org/LeaveBalanceReduction') 'final batch follows ResidentialCare reporting, stats, reduction and Leave Balance reporting'
    Assert-True ($leaveBalanceStats.Reads.Count -eq 0 -and $leaveBalanceReduction.Reads.Count -eq 0 -and $leaveBalanceReporting.Reads.Count -eq 0) 'ordering does not invent unverified workbook inputs'
    $project = Get-Content -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Raw | ConvertFrom-Json
    $leaveBalanceRoot = Resolve-BatchWorkbookRoot $catalogueRoot $project.batchRunner.workbookRoots.LeaveBalanceStats
    Assert-True ($leaveBalanceStats.Path -eq (Resolve-BatchPath $leaveBalanceRoot 'LBStats.xlsx') -and $leaveBalanceReduction.Path -eq (Resolve-BatchPath $leaveBalanceRoot 'LBReduction.xlsx') -and $leaveBalanceReporting.Path -eq (Resolve-BatchPath $leaveBalanceRoot 'TableauConnect-LB.xlsx')) 'named root resolves the exact three LeaveBalance targets'
    $savedTitles = $catalogue.BatchTitles
    try {
        $catalogue.BatchTitles = @{}
        Assert-True ('A1 [1-AllocationExtracted.xlsx]' -in @(Get-BatchPickList $catalogue)) 'untitled batches still show their ID and files'
    } finally { $catalogue.BatchTitles = $savedTitles }
    $one = Select-BatchPlan $catalogue -RunAll -Units Unit1
    Assert-True ($one.Jobs.Count -eq 25 -and @($one.Jobs | Where-Object Unit -eq 'Org').Count -eq 0) 'Unit scope excludes org'
    Assert-Throws { Select-BatchPlan $catalogue -Units Unit1 -Batches 'A1,A3' } 'disjoint batches cannot bypass intermediate producers'
    $subset = Select-BatchPlan $catalogue -Units Unit1 -Batches 'A3,C2'
    Assert-True ($subset.Jobs.Count -eq 11 -and @($subset.Jobs | Where-Object Batch -notin @('A3','C2.1','C2.2','C2.3')).Count -eq 0) 'independent batch branches remain exact selections'
    # Regression: the 21 October selection refreshed roles and reporting while
    # leaving Capacity/Effort (and other reporting intermediates) out of the run.
    foreach ($sequenceProfile in @('Current', 'ParallelInputs-Unit1Priority')) {
        $gapCatalogue = Get-BatchCatalogue $catalogueRoot $sequenceProfile
        Assert-Throws { Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Workers' } "$sequenceProfile cannot dispatch the retired facility workbook by ID"
        $c1Ids = @('JH-RY/AllocationInput','JH-RY/Settings','JH-RY/Availability','JH-RY/Staff',
            'JH-RY/Role:CS/1','JH-RY/Role:CS/2','JH-RY/Role:CS/3')
        $c1Jobs = @(foreach ($id in $c1Ids) {
            $actual = @($gapCatalogue.Jobs | Where-Object Id -eq $id)[0]
            New-FakeJob $actual.Id $actual.BatchKey -Depends @($actual.Dependencies | Where-Object { $_ -in $c1Ids })
        })
        foreach ($job in $c1Jobs) {
            $job.Reads = @($c1Jobs | Where-Object { $_.Id -in $job.Dependencies } | ForEach-Object Path)
        }
        $c1Run = Invoke-FakeRun $c1Jobs -Slow 'JH-RY/AllocationInput','JH-RY/Settings'
        Assert-True ($c1Run.Code -eq 0 -and
            $c1Run.Events.IndexOf('start:JH-RY/Availability') -gt $c1Run.Events.IndexOf('end:JH-RY/AllocationInput') -and
            $c1Run.Events.IndexOf('start:JH-RY/Availability') -gt $c1Run.Events.IndexOf('end:JH-RY/Settings') -and
            $c1Run.Events.IndexOf('start:JH-RY/Staff') -gt $c1Run.Events.IndexOf('end:JH-RY/Availability')) "$sequenceProfile keeps the retained C1 chain ordered without Workers"
        $c1Failure = Invoke-FakeRun $c1Jobs -Fail 'JH-RY/Availability'
        Assert-True ($c1Failure.Code -eq 1 -and @('JH-RY/Staff','JH-RY/Role:CS/1','JH-RY/Role:CS/2','JH-RY/Role:CS/3' | Where-Object {
            $c1Failure.State.Jobs[$_].Status -ne 'Blocked' -or "start:$_" -in $c1Failure.Events
        }).Count -eq 0) "$sequenceProfile Availability failure still blocks Staff and CS before dispatch"
        Assert-True (@($gapCatalogue.Jobs | Where-Object { $_.Id -match '/MultiRole$' -or $_.RelativePath -match 'MutliRole|MultiRole' }).Count -eq 0) "$sequenceProfile has no MultiRole job or target"
        Assert-True (@($gapCatalogue.Jobs | ForEach-Object Reads | Where-Object { $_ -match 'MutliRole|MultiRole' }).Count -eq 0) "$sequenceProfile never requires a MultiRole file as an input"
        Assert-True (@($gapCatalogue.Jobs | ForEach-Object Dependencies | Where-Object { $_ -match '/MultiRole$' }).Count -eq 0) "$sequenceProfile has no dangling MultiRole prerequisite"
        $profileEffort = @($gapCatalogue.Jobs | Where-Object Id -eq 'BD/Effort')[0]
        $expectedEffortPrerequisites = @('BD/Capacity','BD/Demand','BD/Allocation','BD/Distribution','BD/Staff','BD/Settings')
        Assert-True ((($profileEffort.Dependencies | Sort-Object) -join ',') -eq (($expectedEffortPrerequisites | Sort-Object) -join ',')) "$sequenceProfile preserves every remaining Effort prerequisite"
        # Stand in for the real retained producers with successful metadata receipts.
        # The actual catalogue's Effort edges must allow completion without a MultiRole file or job.
        $effortProducers = @(foreach ($dependency in $profileEffort.Dependencies) {
            $producer = @($gapCatalogue.Jobs | Where-Object Id -eq $dependency)[0]
            New-FakeJob $producer.Id $producer.BatchKey
        })
        $syntheticEffort = New-FakeJob $profileEffort.Id $profileEffort.BatchKey $profileEffort.Dependencies @($effortProducers.Path)
        $effortRun = Invoke-FakeRun @($effortProducers + $syntheticEffort)
        Assert-True ($effortRun.Code -eq 0 -and $effortRun.State.Jobs[$syntheticEffort.Id].Status -eq 'Completed' -and
            @($effortProducers | Where-Object { $effortRun.Events.IndexOf("end:$($_.Id)") -gt $effortRun.Events.IndexOf("start:$($syntheticEffort.Id)") }).Count -eq 0) "$sequenceProfile completes Effort after all retained producers without MultiRole"
        # Retain the catalogue's selected edges through every reporting branch;
        # other facility prerequisites are represented by already saved external inputs.
        $downstreamIds = @('Org/StaffAll','Org/EffortAll','Org/Outcomes','Org/Inefficiencies','Org/Cost','Org/History','Org/HistoryRead','Org/Reporting')
        $selectedIds = @($effortProducers.Id) + $syntheticEffort.Id + $downstreamIds
        $downstreamJobs = @(foreach ($producer in @($gapCatalogue.Jobs | Where-Object Id -in $downstreamIds)) {
            New-FakeJob $producer.Id $producer.BatchKey @($producer.Dependencies | Where-Object { $_ -in $selectedIds })
        })
        $effortFailure = Invoke-FakeRun @($effortProducers + $syntheticEffort + $downstreamJobs) -Fail 'BD/Capacity'
        $affectedIds = @($syntheticEffort.Id) + @($downstreamIds | Where-Object { $_ -ne 'Org/StaffAll' })
        Assert-True ($effortFailure.Code -eq 1 -and @($affectedIds | Where-Object {
            $effortFailure.State.Jobs[$_].Status -ne 'Blocked' -or "start:$_" -in $effortFailure.Events }).Count -eq 0) "$sequenceProfile blocks the entire affected Effort-to-reporting chain before any descendant starts"
        Assert-True (@($effortProducers | Where-Object { $_.Id -ne 'BD/Capacity' -and
            $effortFailure.State.Jobs[$_.Id].Status -ne 'Completed' }).Count -eq 0 -and
            $effortFailure.State.Jobs['Org/StaffAll'].Status -eq 'Completed') "$sequenceProfile lets independent retained producers finish after the affected chain stops"
        $gapMessage = ''
        try { Select-BatchPlan $gapCatalogue -Units JH-RY -Batches C2,O1,O2,O5,O6 | Out-Null }
        catch { $gapMessage = $_.Exception.Message }
        Assert-True ($gapMessage -like '*Selection skips intermediate workbooks*' -and
            $gapMessage -like '*JH-RY/Capacity (U4)*' -and $gapMessage -like '*JH-RY/Effort (U4)*' -and
            $gapMessage -like '*Org/Cost (O3)*' -and $gapMessage -like '*Org/History (O4)*' -and
            $gapMessage -like '*Org/HistoryRead (O4)*') "$sequenceProfile reports every missing intermediate in the failed selection"
        Assert-True ($gapMessage -like '*IncludeDependencies*' -and $gapMessage -like '*Excel has not been opened*') 'gap error explains the remedy and that execution has not started'
        Assert-Throws { Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Role:AIN/1','Org/EffortAll' } "$sequenceProfile exact-file selection cannot bypass the same chain"
        Assert-Throws { Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Capacity','Org/EffortAll' } "$sequenceProfile cannot skip Effort when Capacity is selected"
        Assert-Throws { Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Role:AIN/1','JH-RY/Role:AIN/3' } "$sequenceProfile cannot skip the middle role workbook"
        $gapFilled = Select-BatchPlan $gapCatalogue -Units JH-RY -Batches C2,U4,O1,O2,O3,O4,O5,O6
        Assert-True ($gapFilled.Jobs.Count -eq 25 -and
            'JH-RY/Capacity' -in $gapFilled.Jobs.Id -and 'JH-RY/Effort' -in $gapFilled.Jobs.Id -and
            @($gapFilled.Jobs | Where-Object { $_.Unit -notin @('JH-RY','Org') }).Count -eq 0) "$sequenceProfile accepts the completed chain without selecting other Units or earlier inputs"
        $withDependencies = Select-BatchPlan $gapCatalogue -Units JH-RY -Batches C2,O1,O2,O5,O6 -IncludeDependencies
        Assert-True ('JH-RY/Capacity' -in $withDependencies.Jobs.Id -and 'JH-RY/Effort' -in $withDependencies.Jobs.Id) "$sequenceProfile explicit dependency expansion resolves gaps"
        Assert-True ((Select-BatchPlan $gapCatalogue -Units JH-RY -Batches U4).Jobs.Count -eq 2) "$sequenceProfile standalone consolidation can use saved upstream inputs"
        Assert-True ((Select-BatchPlan $gapCatalogue -Batches O1,O2,O3,O4,O5,O6).Jobs.Count -eq 11) "$sequenceProfile complete organisation-only chain can use saved Unit inputs"
        Assert-True ((Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Effort','Org/EffortAll').Jobs.Count -eq 2) "$sequenceProfile direct selected dependency permits unrelated saved inputs"
        Assert-True ((Select-BatchPlan $gapCatalogue -Workbooks 'BD/Role:AIN/3','TE/Effort').Jobs.Count -eq 2) "$sequenceProfile independent Units do not create a false gap"

        $manualPlan = Select-BatchPlan $gapCatalogue -Units BD,TE,JH-RY -Batches 'C1,C2,U4,01,02,05,06' -UseSavedIntermediates
        Assert-True ($manualPlan.Jobs.Count -eq 56 -and $manualPlan.UseSavedIntermediates -and
            ($manualPlan.SavedIntermediates.Id -join ',') -eq 'Org/Cost,Org/History,Org/HistoryRead' -and
            @($manualPlan.Jobs | Where-Object Batch -in @('O3','O4')).Count -eq 0) "$sequenceProfile exact reported selection deliberately reuses O3/O4 without refreshing them"
        $savedGap = Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Role:AIN/1','Org/EffortAll' -UseSavedIntermediates
        $savedConsumer = @($savedGap.Jobs | Where-Object Id -eq 'Org/EffortAll')[0]
        Assert-True ('JH-RY/Role:AIN/1' -in $savedConsumer.Dependencies -and $savedGap.Jobs.Count -eq 2) "$sequenceProfile selected ancestor still controls ordering through saved intermediates"
        Assert-True ('JH-RY/Role:AIN/1' -notin @($gapCatalogue.Jobs | Where-Object Id -eq 'Org/EffortAll')[0].Dependencies) 'saved-input selection does not mutate the catalogue'
        Assert-Throws { Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Role:AIN/1','Org/EffortAll' } 'a previous saved-input choice does not weaken the next selection'
        $portableGap = ConvertTo-PortableBatchPlan $savedGap
        $gapRoundtrip = $portableGap | ConvertTo-Json -Depth 40 | ConvertFrom-Json
        $gapReconstructed = Select-BatchPlan $gapCatalogue -Workbooks @($gapRoundtrip.Jobs.Id) -UseSavedIntermediates:$gapRoundtrip.UseSavedIntermediates
        Assert-True (($portableGap | ConvertTo-Json -Depth 40 -Compress) -eq ((ConvertTo-PortableBatchPlan $gapReconstructed) | ConvertTo-Json -Depth 40 -Compress)) 'saved-input policy, files and projected ordering survive portable resume reconstruction'
        $oneGap = Select-BatchPlan $gapCatalogue -Workbooks 'JH-RY/Role:AIN/1','JH-RY/Role:AIN/3' -UseSavedIntermediates
        $oneGapState = New-BatchState $oneGap 'one-gap'
        Assert-True ($oneGap.SavedIntermediates.Count -eq 1 -and $oneGapState.SavedIntermediates.Count -eq 1 -and
            @(Get-BatchSavedIntermediateNotice ([pscustomobject]$oneGapState)).Count -eq 3) 'a single skipped file remains an array in status records'
        $standalone = Select-BatchPlan $gapCatalogue -Units JH-RY -Batches U4 -UseSavedIntermediates
        Assert-True (-not $standalone.UseSavedIntermediates -and $standalone.SavedIntermediates.Count -eq 0 -and
            @(Get-BatchSavedIntermediateNotice ([pscustomobject](New-BatchState $standalone 'no-gap'))).Count -eq 0) 'ordinary saved upstream inputs do not produce a false intermediate warning'

        $syntheticGapJobs = @(foreach ($job in $savedGap.Jobs) { New-FakeJob $job.Id $job.BatchKey $job.Dependencies })
        foreach ($limit in @(1,12)) {
            $savedRun = Invoke-FakeRun $syntheticGapJobs -Max $limit -Slow 'JH-RY/Role:AIN/1'
            Assert-True ($savedRun.Code -eq 0 -and $savedRun.Events.IndexOf('end:JH-RY/Role:AIN/1') -lt $savedRun.Events.IndexOf('start:Org/EffortAll')) 'saved gap cannot let a downstream file overtake its selected producer'
            $savedFailure = Invoke-FakeRun $syntheticGapJobs -Max $limit -Fail 'JH-RY/Role:AIN/1'
            Assert-True ($savedFailure.Code -eq 1 -and $savedFailure.State.Jobs['Org/EffortAll'].Status -eq 'Blocked' -and
                'start:Org/EffortAll' -notin $savedFailure.Events) 'saved intermediates cannot bypass a failed selected ancestor'
        }
        [IO.File]::WriteAllText($syntheticGapJobs[0].Path, 'changed selected producer')
        Reset-BatchResume $savedRun.Plan $savedRun.State
        Assert-True ($savedRun.State.Jobs['Org/EffortAll'].Status -eq 'Pending') 'resume invalidates the selected descendant across a saved gap'
        $missingSavedInput = Join-Path $testRoot 'missing-saved-intermediate.data'
        $syntheticGapJobs[-1].Reads = @($missingSavedInput)
        Assert-True (@(Test-BatchPlan ([pscustomobject]@{ Jobs = $syntheticGapJobs })).Count -gt 0) 'saved-output choice retains file-access validation'

        $savedNotice = @(Get-BatchSavedIntermediateNotice $savedGap)
        Assert-True (($savedNotice -join ' ') -like '*JH-RY/Capacity*' -and ($savedNotice -join ' ') -like '*will not propagate*') 'saved-output notice explains the original stale-capacity failure'
        $reportDirectory = Join-Path $testRoot ("saved-gap-report-" + $sequenceProfile)
        [void] [IO.Directory]::CreateDirectory($reportDirectory)
        $savedRun.Plan | Add-Member SequenceProfile $sequenceProfile -Force
        $savedRun.Plan | Add-Member SavedIntermediates $savedGap.SavedIntermediates -Force
        $savedRun.State.Updated = [datetime]::UtcNow.ToString('o')
        $summary = Write-BatchActualReport $savedRun.Plan $savedRun.State $reportDirectory
        Assert-True ($summary.SavedIntermediates.Count -eq $savedGap.SavedIntermediates.Count -and
            [IO.File]::ReadAllText((Join-Path $reportDirectory 'actual-notes.md')).Contains('Selected upstream changes will not propagate')) 'actual report preserves the saved-output limitation'
    }
    $role = Select-BatchPlan $catalogue -Units Unit1 -Batches C2.2
    Assert-True ($role.Jobs.Count -eq 3 -and $role.Jobs[0].Role -eq 'AINC4') 'numbered role mapping'
    Assert-Throws { Select-BatchPlan $catalogue -Units BD -Batches C2.2 } 'AINC4 batch is empty for BD'
    $bdRole = Select-BatchPlan $catalogue -Units BD -Batches C2.4
    Assert-True ($bdRole.Jobs.Count -eq 3 -and $bdRole.Jobs[0].Role -eq 'EN') 'BD EN role uses the dedicated C2.4 batch'
    $csRole = Select-BatchPlan $catalogue -Batches C2.5
    Assert-True ($csRole.Jobs.Count -eq 9 -and @($csRole.Jobs | Where-Object { $_.Role -ne 'CS' -or $_.Unit -notin @('BD','TE','JH-RY') }).Count -eq 0) 'CS batch selects only the nine configured facility role workbooks'
    $bdUnit = Select-BatchPlan $catalogue -RunAll -Units BD
    Assert-True ($bdUnit.Jobs.Count -eq 27 -and @($bdUnit.Jobs | Where-Object Unit -eq 'Org').Count -eq 0) 'BD full Unit scope excludes organisation work and retired Workers'
    $everyUnit = Select-BatchPlan $catalogue -RunAll -Units All
    Assert-True ($everyUnit.Jobs.Count -eq 129 -and @($everyUnit.Jobs | Where-Object Unit -eq 'Org').Count -eq 0) 'selected All runs every Unit and the shared master once without organisation jobs'
    $bdCapacity = @($bdUnit.Jobs | Where-Object Id -eq 'BD/Capacity')[0]
    Assert-True ('BD/Role:EN/3' -in $bdCapacity.Dependencies -and 'BD/Role:AINC4/3' -notin $bdCapacity.Dependencies) 'BD consolidation waits for EN rather than AINC4'
    $role = Select-BatchPlan $catalogue -Units Unit2 -Roles RN
    Assert-True ($role.Jobs.Count -eq 3 -and $role.Jobs[0].BatchKey -eq 'Unit2/C2:RN') 'named role selection'
    $exact = Select-BatchPlan $catalogue -Workbooks 'Unit1/DemandMaster', 'Unit2/Workers'
    Assert-True ($exact.Jobs.Count -eq 2) 'new files selectable without legacy renumbering'
    Assert-Throws { Select-BatchPlan $catalogue -StartAtSequence 24 -EndAtSequence 26 } 'legacy ranges cannot skip newly added intermediate workbooks'
    $legacyRangeIds = @(foreach ($sequence in 24..26) {
        (Select-BatchPlan $catalogue -StartAtSequence $sequence -EndAtSequence $sequence).Jobs.Id
    })
    Assert-True (($legacyRangeIds | Sort-Object) -join ',' -eq 'Unit1/Effort,Unit2/AllocationInput,Unit2/DemandInput') 'legacy cross-Unit numbers preserved'
    Assert-Throws { Select-BatchPlan $catalogue -Units Unit1 -StartAtSequence 13 -EndAtSequence 13 } 'retired MultiRole legacy slot cannot select another workbook'
    $range = Select-BatchPlan $catalogue -Units Unit2 -StartAtSequence 14 -EndAtSequence 16
    Assert-True ($range.Jobs.Count -eq 3 -and $range.Jobs[0].Unit -eq 'Unit2') 'Unit-local range'
    $range = Select-BatchPlan $catalogue -StartAtWorkbook '2. Calculations/Cost/Cost..xlsx'
    Assert-True ($range.Jobs.Count -eq 4) 'start-at legacy workbook'
    $deps = Select-BatchPlan $catalogue -Units Unit1 -Batches A2 -IncludeDependencies
    Assert-True ('Unit1/Intervals' -in $deps.Jobs.Id -and 'Unit1/DemandMaster' -in $deps.Jobs.Id -and 'Unit1/Demand' -notin $deps.Jobs.Id) 'file-level dependency closure'
    Assert-Throws { Select-BatchPlan $catalogue -Workbooks 'Effort.xlsx' } 'ambiguous filename rejected'
    Assert-Throws { Select-BatchPlan $catalogue -Batches C2.99 } 'unknown batch rejected'
    Assert-Throws { Select-BatchPlan $catalogue -Roles TestRole1 } 'unapproved role rejected'
    Assert-Throws { Select-BatchPlan $catalogue -RunAll -Batches A1 } 'mixed selectors rejected'
    Assert-Throws { Select-BatchPlan $catalogue -StartAtSequence 5 -EndAtSequence 3 } 'reversed range rejected'
    Assert-Throws { Resolve-BatchPath $testRoot '../escape.data' } 'path boundary enforced'
    # Named roots may reach the sibling project, while each target remains bounded
    # by that configured directory. Guard checks use only disposable directories.
    $rootGuardParent = Join-Path $testRoot 'root-guards'
    $rootGuardRepo = Join-Path $rootGuardParent 'project'
    $rootGuardSibling = Join-Path $rootGuardParent 'sibling/Analysis'
    [void] [IO.Directory]::CreateDirectory($rootGuardRepo)
    [void] [IO.Directory]::CreateDirectory($rootGuardSibling)
    Assert-True ((Resolve-BatchWorkbookRoot $rootGuardRepo '../sibling/Analysis') -eq [IO.Path]::GetFullPath($rootGuardSibling)) 'repo-relative sibling workbook root is supported'
    Assert-Throws { Resolve-BatchWorkbookRoot $rootGuardRepo '../../outside' } 'workbook root cannot escape the shared parent'
    Assert-Throws { Resolve-BatchWorkbookRoot $rootGuardRepo $rootGuardSibling } 'absolute workbook root is rejected'
    Assert-Throws { Resolve-BatchPath $rootGuardSibling '../escape.data' } 'workbook target cannot escape its named root'
    $rootGuardAlias = Join-Path $rootGuardParent 'alias'
    try {
        [void] (New-Item -ItemType Junction -Path $rootGuardAlias -Target $rootGuardSibling)
        Assert-Throws { Resolve-BatchWorkbookRoot $rootGuardRepo '../alias' } 'workbook root rejects reparse-point aliases'
    } finally {
        if (Test-Path -LiteralPath $rootGuardAlias) { Remove-Item -LiteralPath $rootGuardAlias -Force }
    }
    $shifts = $all.Jobs | Where-Object Id -eq 'Unit1/Shifts'
    Assert-True ('Unit1/Intervals' -in $shifts.Dependencies -and 'Unit1/Demand' -notin $shifts.Dependencies) 'A2 waits for Intervals, not all D2'
    $settings = $all.Jobs | Where-Object Id -eq 'Unit1/Settings'
    Assert-True ('Unit1/AllocationInput' -in $settings.Dependencies) 'allocation before settings'

    $sharedMaster = @($defaultRun.Jobs | Where-Object Id -eq 'Date/DemandMaster')[0]
    Assert-True ($sharedMaster.RelativePath -eq '1. Input/Demand-MasterRoster Manual Read.xlsx') 'shared master resolves at the Date input folder'
    Assert-True (@($defaultRun.Jobs | Where-Object Path -eq $sharedMaster.Path).Count -eq 1) 'default run refreshes the shared master only once'
    Assert-True (@($defaultRun.Jobs | Where-Object Id -in @('BD/DemandMaster','TE/DemandMaster','JH-RY/DemandMaster')).Count -eq 0) 'obsolete facility master targets are absent'
    Assert-True (($sharedMaster.Dependencies -join ',') -eq 'BD/Settings,TE/Settings,JH-RY/Settings') 'shared producer retains the existing declared prerequisites'
    foreach ($unit in @('BD','TE','JH-RY')) {
        $demandInput = @($defaultRun.Jobs | Where-Object Id -eq "$unit/DemandInput")[0]
        Assert-True ('Date/DemandMaster' -in $demandInput.Dependencies -and $sharedMaster.Path -in $demandInput.Reads) "$unit demand extraction waits for and reads the shared producer"
        Assert-True (@($demandInput.Reads | Where-Object { $_ -like "*/UNITS/$unit/1. Input/Demand-MasterRoster*" -or $_ -like "*\UNITS\$unit\1. Input\Demand-MasterRoster*" }).Count -eq 0) "$unit no longer checks the obsolete master input"
        $demandBatch = Select-BatchPlan $catalogue -Units $unit -Batches D1
        Assert-True ($demandBatch.Jobs.Count -eq 2 -and 'Date/DemandMaster' -in $demandBatch.Jobs.Id) "$unit D1 includes its shared producer"
    }
    $sharedSelection = Select-BatchPlan $catalogue -Workbooks 'BD/DemandMaster','TE/DemandMaster','1. Input/Demand-MasterRoster Manual Read.xlsx'
    Assert-True ($sharedSelection.Jobs.Count -eq 1 -and $sharedSelection.Jobs[0].Id -eq 'Date/DemandMaster') 'former job IDs and exact shared path select the same single producer'
    $savedDemand = Select-BatchPlan $catalogue -Workbooks 'BD/DemandInput'
    Assert-True ($savedDemand.Jobs.Count -eq 1 -and $sharedMaster.Path -in $savedDemand.Jobs[0].Reads) 'exact extraction selection validates the shared saved input without silently selecting a refresh'
    $sharedDeps = Select-BatchPlan $catalogue -Units BD -Batches A2 -IncludeDependencies
    Assert-True ('Date/DemandMaster' -in $sharedDeps.Jobs.Id -and 'BD/DemandMaster' -notin $sharedDeps.Jobs.Id) 'dependency expansion follows the relocated producer'
    $parallelCatalogue = Get-BatchCatalogue $catalogueRoot 'ParallelInputs-Unit1Priority'
    $parallelPlan = Select-BatchPlan $parallelCatalogue -RunAll
    $parallelInput = @($parallelPlan.Jobs | Where-Object Id -eq 'TE/DemandInput')[0]
    Assert-True ('Date/DemandMaster' -in $parallelInput.Dependencies -and $sharedMaster.Path -in $parallelInput.Reads) 'alternate sequence profile redirects explicit read jobs too'
    Assert-True (($parallelPlan.Jobs | Select-Object -Last 3 | ForEach-Object Id) -join ',' -eq 'Org/LeaveBalanceStats,Org/LeaveBalanceReduction,Org/LeaveBalanceReporting') 'alternate sequence profile includes the final Leave Balance batch'

    # Synthetic final-chain jobs verify save ordering and failure propagation.
    $syntheticRelease = New-FakeJob 'Org/Reporting' 'Org/O5'
    $syntheticStats = New-FakeJob $leaveBalanceStats.Id $leaveBalanceStats.BatchKey $leaveBalanceStats.Dependencies
    $syntheticReduction = New-FakeJob $leaveBalanceReduction.Id $leaveBalanceReduction.BatchKey $leaveBalanceReduction.Dependencies
    $syntheticReporting = New-FakeJob $leaveBalanceReporting.Id $leaveBalanceReporting.BatchKey $leaveBalanceReporting.Dependencies
    $leaveBalanceResult = Invoke-FakeRun @($syntheticRelease, $syntheticStats, $syntheticReduction, $syntheticReporting) -Slow 'Org/Reporting'
    Assert-True ($leaveBalanceResult.Code -eq 0 -and $leaveBalanceResult.Events.IndexOf('end:Org/Reporting') -lt $leaveBalanceResult.Events.IndexOf('start:Org/LeaveBalanceStats')) 'final batch starts after the preceding reporting save'
    Assert-True ($leaveBalanceResult.Events.IndexOf('end:Org/LeaveBalanceStats') -lt $leaveBalanceResult.Events.IndexOf('start:Org/LeaveBalanceReduction') -and $leaveBalanceResult.Events.IndexOf('end:Org/LeaveBalanceReduction') -lt $leaveBalanceResult.Events.IndexOf('start:Org/LeaveBalanceReporting')) 'reduction starts after stats saves and reporting starts after reduction saves'
    Assert-True (@($leaveBalanceResult.Events | Where-Object { $_ -eq 'start:Org/LeaveBalanceStats' }).Count -eq 1 -and @($leaveBalanceResult.Events | Where-Object { $_ -eq 'start:Org/LeaveBalanceReduction' }).Count -eq 1 -and @($leaveBalanceResult.Events | Where-Object { $_ -eq 'start:Org/LeaveBalanceReporting' }).Count -eq 1) 'each final workbook dispatches once'
    $leaveBalanceFailed = Invoke-FakeRun @($syntheticRelease, $syntheticStats, $syntheticReduction, $syntheticReporting) -Fail 'Org/Reporting'
    Assert-True ($leaveBalanceFailed.Code -eq 1 -and $leaveBalanceFailed.State.Jobs[$syntheticStats.Id].Status -eq 'Blocked' -and $leaveBalanceFailed.State.Jobs[$syntheticReporting.Id].Status -eq 'Blocked') 'failed preceding reporting blocks the final batch'
    $leaveBalanceFailed = Invoke-FakeRun @($syntheticRelease, $syntheticStats, $syntheticReduction, $syntheticReporting) -Fail $syntheticStats.Id
    Assert-True ($leaveBalanceFailed.Code -eq 1 -and $leaveBalanceFailed.State.Jobs[$syntheticReporting.Id].Status -eq 'Blocked') 'failed Leave Balance stats blocks its reporting workbook'
    $leaveBalanceFailed = Invoke-FakeRun @($syntheticRelease, $syntheticStats, $syntheticReduction, $syntheticReporting) -Fail $syntheticReduction.Id
    Assert-True ($leaveBalanceFailed.Code -eq 1 -and $leaveBalanceFailed.State.Jobs[$syntheticReporting.Id].Status -eq 'Blocked') 'failed Leave Balance reduction blocks its reporting workbook'

    # Exercise the resolved graph with synthetic files: one shared write must
    # finish before any consumer, and its failure blocks every selected consumer.
    $syntheticMaster = New-FakeJob 'Date/DemandMaster' 'Date/D1'
    $syntheticConsumers = @(foreach ($unit in @('BD','TE','JH-RY')) {
        $inputJob = @($defaultRun.Jobs | Where-Object Id -eq "$unit/DemandInput")[0]
        New-FakeJob $inputJob.Id $inputJob.BatchKey @($inputJob.Dependencies | Where-Object { $_ -eq $syntheticMaster.Id }) @($syntheticMaster.Path)
    })
    $sharedResult = Invoke-FakeRun (@($syntheticMaster) + $syntheticConsumers)
    Assert-True ($sharedResult.Code -eq 0 -and @($sharedResult.Events | Where-Object { $_ -eq 'start:Date/DemandMaster' }).Count -eq 1) 'shared producer dispatches once'
    foreach ($consumer in $syntheticConsumers) {
        Assert-True ($sharedResult.Events.IndexOf('end:Date/DemandMaster') -lt $sharedResult.Events.IndexOf("start:$($consumer.Id)")) 'consumer starts after shared save completion'
    }
    $sharedFailed = Invoke-FakeRun (@($syntheticMaster) + $syntheticConsumers) -Fail 'Date/DemandMaster'
    Assert-True ($sharedFailed.Code -eq 1 -and @($syntheticConsumers | Where-Object { $sharedFailed.State.Jobs[$_.Id].Status -eq 'Blocked' }).Count -eq 3) 'shared producer failure blocks all three consumers'
    $sharedLock = Join-Path $testRoot ('~$' + (Split-Path $syntheticMaster.Path -Leaf))
    [IO.File]::WriteAllText($sharedLock, 'synthetic lock')
    $lockedSharedIssues = @(Test-BatchPlan $sharedResult.Plan)
    Assert-True (@($lockedSharedIssues | Where-Object { $_ -match 'Excel lock file' }).Count -gt 0) 'shared target and consumer input locks remain enforced'
    Remove-Item -LiteralPath $sharedLock
    Remove-Item -LiteralPath $syntheticMaster.Path
    Assert-True (@(Test-BatchPlan $sharedResult.Plan | Where-Object { $_ -match 'Missing file' }).Count -gt 0) 'missing shared source remains a validation error'

    # A disposable configuration fixture tests dynamic discovery without touching
    # any workbook or the user's saved production role profile.
    $fixtureRoot = Join-Path $testRoot 'catalogue'
    $fixtureDate = Join-Path $fixtureRoot 'CLIENT/DATExx-Whiddon'
    [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureDate 'runner'))
    foreach ($unit in @('Unit1','Unit2')) { [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureDate "UNITS/$unit")) }
    Copy-Item -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Destination $fixtureRoot
    foreach ($file in @('ResidentialCare-BatchCatalogue.psd1','ResidentialCare-BatchApproval.psd1','ResidentialCare-UnitWorkbookSequence.psd1','ResidentialCare-OrgWorkbookSequence.psd1')) {
        Copy-Item -LiteralPath (Join-Path $repoRoot "CLIENT/DATExx-Whiddon/runner/$file") -Destination (Join-Path $fixtureDate 'runner')
    }
    $fixtureSettingsPath = Join-Path $fixtureDate 'runner/ResidentialCare-BatchApproval.psd1'
    $fixtureSettingsText = [IO.File]::ReadAllText($fixtureSettingsPath)
    $fixtureSettingsText = [regex]::Replace($fixtureSettingsText, '(?ms)^    AdditionalUnits = @\(.*?^    \)', '    AdditionalUnits = @()')
    $fixtureSettingsText = [regex]::Replace($fixtureSettingsText, '(?ms)^    SharedUnitJobs = @\(.*?^    \)', '    SharedUnitJobs = @()')
    $fixtureSettingsText = [regex]::Replace($fixtureSettingsText, '(?m)^    OrganisationUnits = .*$', "    OrganisationUnits = @('Unit1', 'Unit2')")
    $fixtureSettingsText = [regex]::Replace($fixtureSettingsText, '(?m)^    RunAllUnits = .*$', "    RunAllUnits = @('Unit1', 'Unit2')")
    $fixtureSettingsText = [regex]::Replace($fixtureSettingsText, '(?m)^    RunAllIncludeOrg = .*$', '    RunAllIncludeOrg = $true')
    [IO.File]::WriteAllText($fixtureSettingsPath, $fixtureSettingsText)
    $profilePath = Join-Path $fixtureDate 'runner/ResidentialCare-ClientRunProfile.psd1'
    foreach ($count in @(1, 2, 5)) {
        $entries = @(1..$count | ForEach-Object { "@{ Folder = 'Role$_'; Enabled = `$true }" }) -join ','
        $content = "@{ SchemaVersion = 1; Roles = @($entries, @{ Folder = 'Disabled'; Enabled = `$false }); RoleWorkbookOrder = @('CapacityDistrib(A.1)-shifts.xlsx','CapacityDistrib(A.2)-shifts.xlsx','CapacityDistrib(B)-shifts.xlsx') }"
        [IO.File]::WriteAllText($profilePath, $content)
        $dynamic = Get-BatchCatalogue $fixtureRoot
        Assert-True ($dynamic.Jobs.Count -eq (2 * (16 + 3 * $count) + 11)) "dynamic expansion for $count roles"
        Assert-True (@($dynamic.Jobs | Where-Object Role -eq 'Disabled').Count -eq 0) 'disabled roles excluded'
        $dynamicPlan = Select-BatchPlan $dynamic -Units Unit1 -Batches "C2.$count"
        Assert-True ($dynamicPlan.Jobs.Count -eq 3 -and $dynamicPlan.Jobs[0].Role -eq "Role$count") 'last dynamic role identity'
        $dynamicPicks = @(Get-BatchPickList $dynamic)
        Assert-True (@($dynamicPicks | Where-Object { $_ -like "C2.$count - Role capacity - Role$count *" }).Count -eq 1) 'batch picker expands current enabled roles'
        Assert-True (@($dynamicPicks | Where-Object { $_ -match 'Disabled|TestRole' }).Count -eq 0) 'batch picker excludes disabled or unapproved roles'
    }
    $missing = @(Test-BatchPlan (Select-BatchPlan $dynamic -Units Unit1 -Batches A1))
    Assert-True ($missing.Count -gt 0) 'missing files fail validation without Excel'
    $settingsPath = Join-Path $fixtureDate 'runner/ResidentialCare-BatchApproval.psd1'
    $settingsText = [IO.File]::ReadAllText($settingsPath)
    [IO.File]::WriteAllText($settingsPath, ($settingsText -replace 'AdditionalInputs = @\{\}', "AdditionalInputs = @{ 'Unit1/Workers' = @('mock-extra-input.data') }"))
    $changedSettings = Get-BatchCatalogue $fixtureRoot
    Assert-True ($changedSettings.Fingerprint -ne $dynamic.Fingerprint) 'operational setting changes still invalidate resume identity'
    $workers = $changedSettings.Jobs | Where-Object Id -eq 'Unit1/Workers'
    Assert-True ((Join-Path $fixtureRoot 'mock-extra-input.data') -in $workers.Reads) 'additional input protection is retained without the approval gate'

    foreach ($unit in @('Unit10', 'Unit3')) { [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureDate "UNITS/$unit")) }
    $expandedUnits = Get-BatchCatalogue $fixtureRoot
    $expandedUnitPicks = @(Get-UnitPickList $expandedUnits | Where-Object { $_ -match '^Unit[0-9]+$' })
    Assert-True (($expandedUnitPicks -join ',') -eq 'Unit1,Unit2,Unit3,Unit10') 'Unit picker follows current discovery in numeric order, not a hard-coded list'

    $savedPlan = Join-Path $testRoot 'roundtrip.json'
    Write-BatchJson $savedPlan $all
    $roundtrip = Read-BatchJson $savedPlan
    $reconstructed = Select-BatchPlan $catalogue -Workbooks @($roundtrip.Jobs | ForEach-Object Id)
    Assert-True (($roundtrip.Jobs | ConvertTo-Json -Depth 30 -Compress) -eq ($reconstructed.Jobs | ConvertTo-Json -Depth 30 -Compress)) 'immutable plan survives JSON roundtrip'

    # Synthetic files are plain text with .data extensions: no workbook is opened.
    $parallel = @(1..12 | ForEach-Object { New-FakeJob "parallel$_" })
    $result = Invoke-FakeRun $parallel
    Assert-True ($result.Code -eq 0 -and $result.Max -eq 12) 'twelve independent batches run concurrently'
    $result = Invoke-FakeRun $parallel -Max 1
    Assert-True ($result.Code -eq 0 -and $result.Max -eq 1) 'serial override'
    Assert-Throws { Invoke-FakeRun $parallel -Max 13 } 'hard maximum twelve'
    $first = New-FakeJob 'U1/A1' 'U1/A'
    $second = New-FakeJob 'U1/A2' 'U1/A' @($first.Id) @($first.Path)
    $otherUnit = New-FakeJob 'U2/D1' 'U2/D'
    $org = New-FakeJob 'Org/O3' 'Org/O3'
    $result = Invoke-FakeRun @($first, $second, $otherUnit, $org)
    Assert-True ($result.Max -eq 3) 'parallel work is not limited to C2 or one Unit'
    Assert-True ($result.Events.IndexOf('end:U1/A1') -lt $result.Events.IndexOf('start:U1/A2')) 'within-batch sequence'
    $intervals = New-FakeJob 'Intervals' 'D2'
    $demand = New-FakeJob 'Demand' 'D2' @('Intervals') @($intervals.Path)
    $shiftsJob = New-FakeJob 'Shifts' 'A2' @('Intervals') @($intervals.Path)
    $result = Invoke-FakeRun @($intervals, $demand, $shiftsJob) -Slow Demand
    Assert-True ($result.Events.IndexOf('start:Shifts') -lt $result.Events.IndexOf('end:Demand')) 'file-level release before producing batch completes'
    $a = New-FakeJob 'pause1' 'paused'
    $b = New-FakeJob 'pause2' 'paused' @('unblock')
    $unblock = New-FakeJob 'unblock'
    $result = Invoke-FakeRun @($a, $b, $unblock) -Max 1
    Assert-True ($result.Code -eq 0 -and $result.Events.IndexOf('start:unblock') -lt $result.Events.IndexOf('start:pause2')) 'waiting batch releases its slot'
    $producer = New-FakeJob 'producer'
    $consumer = New-FakeJob 'consumer' 'consumer' @() @($producer.Path)
    Assert-True (Test-BatchConflict $producer @($consumer)) 'producer cannot overwrite an active reader input'
    $result = Invoke-FakeRun @($consumer, $producer)
    Assert-True ($result.Max -eq 1) 'read/write exclusion applies even without explicit dependency'
    $reader2 = New-FakeJob 'reader2' 'reader2' @() @($producer.Path)
    $result = Invoke-FakeRun @($consumer, $reader2)
    Assert-True ($result.Max -eq 2) 'stable shared readers can overlap'
    $reader2.Exclusive = $true
    $result = Invoke-FakeRun @($consumer, $reader2)
    Assert-True ($result.Max -eq 1) 'exclusive side-effect jobs run alone'
    $consumer.Dependencies = @('producer')
    $unrelated = New-FakeJob 'independent'
    $result = Invoke-FakeRun @($producer, $consumer, $unrelated) -Fail producer
    Assert-True ($result.Code -eq 1 -and $result.State.Jobs.consumer.Status -eq 'Blocked' -and $result.State.Jobs.independent.Status -eq 'Completed') 'failure blocks descendants, not independent work'
    $result = Invoke-FakeRun @($consumer)
    Assert-True ($result.Code -eq 0) 'unselected saved producer may be reused'
    $result = Invoke-FakeRun @($producer) -Fail producer -FailureCode 124
    Assert-True ($result.Code -eq 124) 'timeout exit retained'
    $result = Invoke-FakeRun @($first, $second, $otherUnit) -Max 1 -Stop AfterCurrent
    Assert-True ($result.Code -eq 3 -and $result.State.Jobs[$first.Id].Status -eq 'Completed' -and $result.State.Jobs[$second.Id].Status -eq 'Stopped') 'stop after current file'
    $result = Invoke-FakeRun @($first, $second, $otherUnit) -Max 1 -Stop AfterBatch
    Assert-True ($result.Code -eq 3 -and $result.State.Jobs[$second.Id].Status -eq 'Completed' -and $result.State.Jobs[$otherUnit.Id].Status -eq 'Stopped') 'stop after active batch'
    $result = Invoke-FakeRun @($first, $second) -Max 1 -Stop Now -Uncertain
    Assert-True ($result.Code -eq 3 -and $result.State.Jobs[$first.Id].NeedsInspection) 'stop now records uncertain work'
    Assert-Throws { Reset-BatchResume $result.Plan $result.State } 'uncertain save cannot automatically resume'
    $result = Invoke-FakeRun @($first, $second, $otherUnit)
    Reset-BatchResume $result.Plan $result.State
    Assert-True (@($result.State.Jobs.Values | Where-Object Status -ne 'Completed').Count -eq 0) 'unchanged completed work preserved on resume'
    [IO.File]::WriteAllText($first.Path, 'changed upstream output')
    Reset-BatchResume $result.Plan $result.State
    Assert-True ($result.State.Jobs[$first.Id].Status -eq 'Pending' -and $result.State.Jobs[$second.Id].Status -eq 'Pending' -and $result.State.Jobs[$otherUnit.Id].Status -eq 'Completed') 'resume invalidates affected descendants only'
    $result.Plan.Fingerprint = 'changed'
    Assert-Throws { Reset-BatchResume $result.Plan $result.State } 'changed configuration rejects resume'

    # Gate leases cover both directions and orphan-worker recovery.
    $legacyLease = Enter-RefreshRunGate $testRoot
    $nested = Enter-RefreshRunGate $testRoot
    Assert-Throws { Enter-RefreshRunGate $testRoot -Exclusive } 'legacy blocks batch'
    $nested.Dispose(); $legacyLease.Dispose()
    $batchLease = Enter-RefreshRunGate $testRoot -Exclusive
    Assert-Throws { Enter-RefreshRunGate $testRoot } 'batch blocks legacy/test'
    Assert-Throws { Enter-RefreshRunGate $testRoot -Exclusive } 'second coordinator blocked'
    $workerLease = Enter-RefreshRunGate $testRoot -BatchWorker
    $batchLease.Dispose()
    Assert-Throws { Enter-RefreshRunGate $testRoot -Exclusive } 'orphan worker blocks new batch run'
    Assert-Throws { Enter-RefreshRunGate $testRoot } 'orphan worker blocks legacy run'
    $workerLease.Dispose()
    $batchLease = Enter-RefreshRunGate $testRoot -Exclusive; $batchLease.Dispose()
    Assert-RefreshRunGateAvailable $testRoot

    # Exercise the actual adapter and process receipt path with a test-only
    # supervisor in a disposable directory. The real Excel engine is not copied.
    $adapterRoot = Join-Path $testRoot 'adapter'
    [void] [IO.Directory]::CreateDirectory($adapterRoot)
    foreach ($name in @('Invoke-BatchWorkbook.ps1','RefreshRunGate.ps1')) {
        Copy-Item -LiteralPath (Join-Path $repoRoot "CLIENT/DATExx-Whiddon/runner/src/$name") -Destination $adapterRoot
    }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'fixtures/BatchMockWorkbookWorker.ps1') -Destination (Join-Path $adapterRoot 'Invoke-AllUnitsExcelWorkbookRefresh.ps1')
    $mockTarget = New-FakeJob 'process-adapter'
    $mockInput = New-FakeJob 'process-adapter-input'
    $inputSnapshotPath = Join-Path $adapterRoot 'input-snapshot.json'
    Write-BatchJson $inputSnapshotPath @{ SchemaVersion = 1; Target = $mockTarget.Path; Inputs = @{$mockInput.Path=(Get-BatchFileStamp $mockInput.Path)} }
    $requestPath = Join-Path $adapterRoot 'request.json'
    Write-BatchJson $requestPath @{
        Id = $mockTarget.Id; Path = $mockTarget.Path; DateRoot = $testRoot; ParentId = $PID
        ParentStarted = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks
        LogPath = (Join-Path $adapterRoot 'refresh.log'); StatusPath = (Join-Path $adapterRoot 'status.txt')
        StopPath = (Join-Path $adapterRoot 'stop.txt'); Visible = 'false'; TimeoutMinutes = 1
        InputSnapshotPath = $inputSnapshotPath
    }
    $info = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh).Source)
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    foreach ($arg in @('-NoProfile','-File',(Join-Path $adapterRoot 'Invoke-BatchWorkbook.ps1'),'-RequestPath',$requestPath)) { $info.ArgumentList.Add($arg) }
    $process = [Diagnostics.Process]::Start($info)
    $handle = @{ Process = $process; Stdout = $process.StandardOutput.ReadToEndAsync(); Stderr = $process.StandardError.ReadToEndAsync(); Directory = $adapterRoot }
    do { Start-Sleep -Milliseconds 100; $receipt = Receive-BatchWorkbookProcess $handle } while ($null -eq $receipt)
    Assert-True ($receipt.ExitCode -eq 0 -and -not $receipt.NeedsInspection) "process adapter receipt: $($receipt.Message)"
    Assert-True ((Read-BatchJson (Join-Path $adapterRoot 'refresh-worker-mock.json')).verifiedInputs -eq 1) 'actual adapter forwards the input snapshot to its worker'

    # A detected policy block and an uncertain Save must retain their original
    # explanations when converted into durable batch completion records.
    foreach ($case in @(
        @{ Phase = 'WaitingToSave'; Message = 'External-user policy blocked access; no save attempted.'; Uncertain = $false },
        @{ Phase = 'Saving'; Message = 'Declared input changed since dispatch; refresh cannot be accepted.'; Uncertain = $true }
    )) {
        $receiptDirectory = Join-Path $testRoot $case.Phase
        [void] [IO.Directory]::CreateDirectory($receiptDirectory)
        Write-BatchJson (Join-Path $receiptDirectory 'refresh-worker-failed.json') @{
            phase = $case.Phase; saved = $false; error = @{ Phase = $case.Phase; Message = $case.Message }
        }
        $finishedProcess = [pscustomobject]@{ HasExited = $true; ExitCode = 1 }
        $finishedProcess | Add-Member ScriptMethod Dispose {}
        $receipt = Receive-BatchWorkbookProcess @{
            Process = $finishedProcess; Directory = $receiptDirectory
            Stdout = [Threading.Tasks.Task]::FromResult([string] ''); Stderr = [Threading.Tasks.Task]::FromResult([string] '')
        }
        Assert-True ($receipt.Message.Contains($case.Message)) "original worker explanation preserved: $($case.Phase)"
        Assert-True ($receipt.ExitCode -eq 1 -and $receipt.NeedsInspection -eq $case.Uncertain) "save uncertainty preserved: $($case.Phase)"
    }

    $errors = @()
    $scripts = @(Get-ChildItem (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner') -Recurse -Filter '*.ps1')
    $scripts += @(Get-ChildItem (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS') -Recurse -Filter '*WorkflowRefresh*.ps1')
    foreach ($script in $scripts) {
        $tokens = $null; $parseErrors = $null
        [void] [Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref] $tokens, [ref] $parseErrors)
        $errors += @($parseErrors)
    }
    Assert-True ($errors.Count -eq 0) "PowerShell parse checks: $(($errors | ForEach-Object Message) -join '; ')"
    Assert-True (@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object Id -notin $beforeExcel).Count -eq 0) 'no Excel process was started'
    Write-Host "PASS: $script:assertions assertions; selection, twelve-batch scheduling, stop/resume, locks and syntax. No Excel opened."
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]] '\/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-batch-*') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
