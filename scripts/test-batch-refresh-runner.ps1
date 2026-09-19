param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchScheduling.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-batch-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
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
    $path = Join-Path $testRoot ($Id.Replace('/', '-') + '.data')
    [IO.File]::WriteAllText($path, 'original')
    return [pscustomobject]@{ Id = $Id; BatchKey = $Batch; Batch = $Batch; Path = $path; Reads = $Reads; Dependencies = $Depends; Exclusive = $false }
}
function Invoke-FakeRun {
    param([object[]] $Jobs, [int] $Max = 7, [string[]] $Fail = @(), [string] $Stop = '',
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
    $catalogue = Get-BatchCatalogue $repoRoot
    $all = Select-BatchPlan $catalogue -RunAll
    Assert-True ($all.Jobs.Count -eq 60) 'two Units plus organisation should resolve 60 files'
    Assert-True (@($all.Jobs | Group-Object BatchKey).Count -eq 27) 'expected 27 batches'
    Assert-True ($all.Jobs[0].Id -eq 'Unit1/AllocationInput') 'numeric Unit and catalogue order'
    Assert-True (@($all.Jobs | Where-Object RelativePath -match 'TestRole').Count -eq 0) 'test roles excluded'
    Assert-True (-not $catalogue.Settings.ContainsKey('Approved')) 'operational settings do not require approval metadata'
    Assert-True (($catalogue.OrgUnits -join ',') -eq 'Unit1,Unit2') 'organisation consumer scope is retained'
    $unitPicks = @(Get-UnitPickList $catalogue)
    Assert-True ('Unit1' -in $unitPicks -and 'Unit2' -in $unitPicks) 'Unit picker lists exact discovered IDs'
    Assert-True ('All - all listed Units' -in $unitPicks) 'Unit picker offers All'
    Assert-True ('Blank = cancel this selection; nothing will run.' -in $unitPicks) 'dedicated Unit selection explains empty-input cancellation'
    Assert-True ('Blank = all listed Units.' -in @(Get-UnitPickList $catalogue -BlankMeansAll)) 'batch and role Unit qualifiers explain blank means all'
    $pickList = @(Get-BatchPickList $catalogue)
    Assert-True ('D1 - Demand transformation [Demand-MasterRoster Manual Read.xlsx, 2-DemandExtract.xlsx]' -in $pickList) 'batch choices show ID, title and ordered comma-separated filenames'
    Assert-True (@($pickList | Where-Object { $_ -like 'D1 -*' }).Count -eq 1) 'shared Unit batch choices are not duplicated'
    Assert-True ('C2.2 - Role capacity - AINC4 [CapacityDistrib(A.1)-shifts.xlsx, CapacityDistrib(A.2)-shifts.xlsx, CapacityDistrib(B)-shifts.xlsx]' -in $pickList) 'role choices show actual role mapping and all files'
    Assert-True (@($pickList | Where-Object { $_ -like 'C2 - All enabled*' }).Count -eq 1) 'all-role shorthand is offered once'
    Assert-True ('O1 - Cross-unit assembly [StafMasterList-All.xlsx, Effort-All.xlsx]' -in $pickList) 'organisation choices include files'
    $savedTitles = $catalogue.BatchTitles
    try {
        $catalogue.BatchTitles = @{}
        Assert-True ('A1 [1-AllocationExtracted.xlsx]' -in @(Get-BatchPickList $catalogue)) 'untitled batches still show their ID and files'
    } finally { $catalogue.BatchTitles = $savedTitles }
    $one = Select-BatchPlan $catalogue -RunAll -Units Unit1
    Assert-True ($one.Jobs.Count -eq 26 -and @($one.Jobs | Where-Object Unit -eq 'Org').Count -eq 0) 'Unit scope excludes org'
    $subset = Select-BatchPlan $catalogue -Units Unit1 -Batches 'A1,A3'
    Assert-True ($subset.Jobs.Count -eq 4 -and @($subset.Jobs | Where-Object Batch -notin @('A1','A3')).Count -eq 0) 'disjoint batches do not become ranges'
    $role = Select-BatchPlan $catalogue -Units Unit1 -Batches C2.2
    Assert-True ($role.Jobs.Count -eq 3 -and $role.Jobs[0].Role -eq 'AINC4') 'numbered role mapping'
    $role = Select-BatchPlan $catalogue -Units Unit2 -Roles RN
    Assert-True ($role.Jobs.Count -eq 3 -and $role.Jobs[0].BatchKey -eq 'Unit2/C2:RN') 'named role selection'
    $exact = Select-BatchPlan $catalogue -Workbooks 'Unit1/DemandMaster', 'Unit2/Workers'
    Assert-True ($exact.Jobs.Count -eq 2) 'new files selectable without legacy renumbering'
    $range = Select-BatchPlan $catalogue -StartAtSequence 24 -EndAtSequence 26
    Assert-True (($range.Jobs.Id | Sort-Object) -join ',' -eq 'Unit1/Effort,Unit2/AllocationInput,Unit2/DemandInput') 'legacy cross-Unit numbers preserved'
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
    $shifts = $all.Jobs | Where-Object Id -eq 'Unit1/Shifts'
    Assert-True ('Unit1/Intervals' -in $shifts.Dependencies -and 'Unit1/Demand' -notin $shifts.Dependencies) 'A2 waits for Intervals, not all D2'
    $settings = $all.Jobs | Where-Object Id -eq 'Unit1/Settings'
    Assert-True ('Unit1/AllocationInput' -in $settings.Dependencies) 'allocation before settings'
    $capacity = $all.Jobs | Where-Object Id -eq 'Unit1/Capacity'
    Assert-True (@('Unit1/Role:AIN/3','Unit1/Role:AINC4/3','Unit1/Role:RN/3' | Where-Object { $_ -notin $capacity.Dependencies }).Count -eq 0) 'production RoleOutputs still use each role terminal workbook'

    $ain2Catalogue = Get-BatchCatalogue $repoRoot Current AIN2
    $ain2Plan = Select-BatchPlan $ain2Catalogue -Units Unit1 -Roles AIN2
    Assert-True (($ain2Plan.Jobs.Id -join ',') -eq 'Unit1/Role:AIN2/1,Unit1/Role:AIN2/2') 'isolated AIN2 profile resolves only its ordered A1 and B workbooks'
    Assert-True ((Split-Path $ain2Plan.Jobs[1].RelativePath -Leaf) -eq 'CapacityDistrib(B)-shifts.xlsx') 'AIN2 terminal workbook is B'
    $ain2Capacity = $ain2Catalogue.Jobs | Where-Object Id -eq 'Unit1/Capacity'
    Assert-True ('Unit1/Role:AIN2/2' -in $ain2Capacity.Dependencies -and 'Unit1/Role:AIN2/3' -notin $ain2Capacity.Dependencies) 'RoleOutputs dependency follows the configured AIN2 terminal index'
    Assert-Throws { Select-BatchPlan $ain2Catalogue -RunAll } 'isolated AIN2 profile rejects RunAll'
    Assert-Throws { Select-BatchPlan $ain2Catalogue -Units Unit2 -Roles AIN2 } 'isolated AIN2 profile rejects other Units'
    Assert-Throws { Select-BatchPlan $ain2Catalogue -Units Unit1 -Batches U4 } 'isolated AIN2 profile rejects production Capacity jobs'
    Assert-Throws { Select-BatchPlan $ain2Catalogue -Units Unit1 -Roles AIN2 -IncludeDependencies } 'isolated AIN2 profile cannot schedule production dependencies'
    Assert-Throws { Select-BatchPlan $ain2Catalogue -Units Unit1 -Roles AIN2 -IncludeOrg } 'isolated AIN2 profile rejects organisation jobs'
    $preparedProfile = Import-PowerShellDataFile (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/ResidentialCare-ClientRunProfile-AINTwoFile-Prepared.psd1')
    $preparedAIN = @($preparedProfile.Roles | Where-Object Folder -eq 'AIN')[0]
    $preparedThreeFileRoles = @($preparedProfile.Roles | Where-Object Folder -in @('AINC4', 'RN'))
    Assert-True (@($preparedAIN.WorkbookOrder).Count -eq 2 -and $preparedAIN.WorkbookOrder[-1] -eq 'CapacityDistrib(B)-shifts.xlsx') 'prepared promotion profile makes B the AIN terminal workbook'
    Assert-True (@($preparedThreeFileRoles | Where-Object { @($_.WorkbookOrder).Count -ne 3 }).Count -eq 0) 'prepared promotion profile leaves AINC4 and RN on three workbooks'

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
    $profilePath = Join-Path $fixtureDate 'runner/ResidentialCare-ClientRunProfile.psd1'
    foreach ($count in @(1, 2, 5)) {
        $entries = @(1..$count | ForEach-Object { "@{ Folder = 'Role$_'; Enabled = `$true }" }) -join ','
        $content = "@{ SchemaVersion = 1; Roles = @($entries, @{ Folder = 'Disabled'; Enabled = `$false }); RoleWorkbookOrder = @('CapacityDistrib(A.1)-shifts.xlsx','CapacityDistrib(A.2)-shifts.xlsx','CapacityDistrib(B)-shifts.xlsx') }"
        [IO.File]::WriteAllText($profilePath, $content)
        $dynamic = Get-BatchCatalogue $fixtureRoot
        Assert-True ($dynamic.Jobs.Count -eq (2 * (17 + 3 * $count) + 8)) "dynamic expansion for $count roles"
        Assert-True (@($dynamic.Jobs | Where-Object Role -eq 'Disabled').Count -eq 0) 'disabled roles excluded'
        $dynamicPlan = Select-BatchPlan $dynamic -Units Unit1 -Batches "C2.$count"
        Assert-True ($dynamicPlan.Jobs.Count -eq 3 -and $dynamicPlan.Jobs[0].Role -eq "Role$count") 'last dynamic role identity'
        $dynamicPicks = @(Get-BatchPickList $dynamic)
        Assert-True (@($dynamicPicks | Where-Object { $_ -like "C2.$count - Role capacity - Role$count *" }).Count -eq 1) 'batch picker expands current enabled roles'
        Assert-True (@($dynamicPicks | Where-Object { $_ -match 'Disabled|TestRole' }).Count -eq 0) 'batch picker excludes disabled or unapproved roles'
    }
    $content = "@{ SchemaVersion = 1; Roles = @(@{ Folder = 'Role1'; Enabled = `$true; WorkbookOrder = @('CapacityDistrib(A.1)-shifts.xlsx','CapacityDistrib(B)-shifts.xlsx') }, @{ Folder = 'Role2'; Enabled = `$true }); RoleWorkbookOrder = @('CapacityDistrib(A.1)-shifts.xlsx','CapacityDistrib(A.2)-shifts.xlsx','CapacityDistrib(B)-shifts.xlsx') }"
    [IO.File]::WriteAllText($profilePath, $content)
    $perRole = Get-BatchCatalogue $fixtureRoot
    $role1Jobs = @($perRole.Jobs | Where-Object Role -eq 'Role1')
    $role2Jobs = @($perRole.Jobs | Where-Object Role -eq 'Role2')
    Assert-True ($role1Jobs.Count -eq 4 -and $role2Jobs.Count -eq 6) 'per-role workbook order overrides the shared default across two Units'
    $fixtureCapacity = $perRole.Jobs | Where-Object Id -eq 'Unit1/Capacity'
    Assert-True ('Unit1/Role:Role1/2' -in $fixtureCapacity.Dependencies -and 'Unit1/Role:Role2/3' -in $fixtureCapacity.Dependencies) 'mixed role lengths resolve their own terminal dependencies'
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
    Assert-True ($result.Code -eq 0 -and $result.Max -eq 7) 'seven independent batches run concurrently'
    $result = Invoke-FakeRun $parallel -Max 1
    Assert-True ($result.Code -eq 0 -and $result.Max -eq 1) 'serial override'
    Assert-Throws { Invoke-FakeRun $parallel -Max 8 } 'hard maximum seven'
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
    Write-Host "PASS: $script:assertions assertions; selection, seven-batch scheduling, stop/resume, locks and syntax. No Excel opened."
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]] '\/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-batch-*') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
