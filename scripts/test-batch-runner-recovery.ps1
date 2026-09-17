# Real Windows file locks plus synthetic workers: never opens an Excel workbook.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchScheduling.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-recovery-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
$script:checks = 0
$helpers = [Collections.Generic.List[object]]::new()
function Assert-Recovery {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}
function Start-StatusLock {
    param([string] $Path, [switch] $ReadOnly)
    $ready = Join-Path $testRoot ([guid]::NewGuid().ToString('N') + '.ready')
    $info = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh).Source)
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    foreach ($arg in @('-NoProfile', '-File', (Join-Path $PSScriptRoot 'fixtures/BatchStatusLockHolder.ps1'), '-Path', $Path, '-ReadyPath', $ready)) { $info.ArgumentList.Add($arg) }
    if ($ReadOnly) { $info.ArgumentList.Add('-ReadOnly') }
    $process = [Diagnostics.Process]::Start($info)
    $helpers.Add($process)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath $ready)) {
        if ($process.HasExited -or $timer.Elapsed.TotalSeconds -gt 10) { throw 'Status lock helper failed to become ready.' }
        Start-Sleep -Milliseconds 10
    }
    return $process
}
function Invoke-RecoveryFixture {
    param([ValidateSet('Transient', 'Permanent', 'BeforeDispatch', 'Slow', 'Heartbeat', 'StopDuringPublish')] [string] $Mode)
    $directory = Join-Path $testRoot $Mode
    [void] [IO.Directory]::CreateDirectory($directory)
    $jobs = @(1..3 | ForEach-Object {
        $path = Join-Path $directory "file$_.data"
        [IO.File]::WriteAllText($path, 'original')
        [pscustomobject]@{ Id = "file$_"; BatchKey = "batch$_"; Path = $path; Reads = @(); Dependencies = @(); Exclusive = $false }
    })
    $jobs[2].Dependencies = @('file1'); $jobs[2].Reads = @($jobs[0].Path)
    $plan = [pscustomobject]@{ SchemaVersion = 1; Fingerprint = 'fixture'; DateRoot = $directory; Jobs = $jobs }
    $state = New-BatchState $plan $Mode
    $sim = @{ Writes = 0; Starts = [Collections.Generic.List[string]]::new(); Lock = $null; Polls = 0 }
    $persist = {
        param($path, $value)
        $sim.Writes++
        if ($Mode -eq 'StopDuringPublish' -and $sim.Writes -eq 1) { Write-BatchJson (Join-Path $directory 'stop.json') @{ Mode = 'AfterCurrent' } }
        if ($Mode -eq 'BeforeDispatch' -and $sim.Writes -eq 1) { throw 'Original failure: state publication unavailable before dispatch.' }
        if ($sim.Writes -eq 3) {
            if ($Mode -eq 'Transient') { $null = Start-StatusLock $path }
            if ($Mode -eq 'Permanent') { $sim.Lock = [IO.File]::Open($path, 'Open', 'Read', [IO.FileShare]::ReadWrite) }
        }
        Write-RefreshJson $path $value -RetryMilliseconds $(if ($Mode -eq 'Permanent') { 120 } else { 5000 })
    }.GetNewClosure()
    $start = {
        param($job, $jobDirectory, $options)
        $published = Read-BatchJson (Join-Path $directory 'state.json') -AsHashtable
        Assert-Recovery ($published.Jobs[$job.Id].Status -eq 'Running') 'running intent is durable BEFORE dispatch'
        $sim.Starts.Add($job.Id)
        return @{ Job = $job; Directory = $jobDirectory; Polls = 0 }
    }.GetNewClosure()
    $poll = {
        param($handle)
        $sim.Polls++; $handle.Polls++
        $stop = Test-Path -LiteralPath (Join-Path $handle.Directory 'stop.txt')
        if ($Mode -eq 'Permanent' -and $handle.Job.Id -eq 'file1') {
            # The worker finished saving just before the coordinator's failure.
            [IO.File]::WriteAllText($handle.Job.Path, 'verified saved output')
            return @{ ExitCode = 0; NeedsInspection = $false; Message = 'Confirmed save and close before stop arrived.' }
        }
        if ($stop) { return @{ ExitCode = 3; NeedsInspection = $false; Message = 'Stopped before saving.' } }
        $pollLimit = if ($Mode -in @('Slow', 'Heartbeat')) { 100 } else { 2 }
        if ($handle.Polls -lt $pollLimit) { return $null }
        [IO.File]::WriteAllText($handle.Job.Path, 'saved output')
        return @{ ExitCode = 0; NeedsInspection = $false; Message = 'Confirmed save and close.' }
    }.GetNewClosure()
    try {
        $code = Invoke-BatchSchedule $plan $state $directory -MaxParallelBatches 2 -StartWorker $start -PollWorker $poll -PersistState $persist -ValidateJob {} -PollMilliseconds $(if ($Mode -eq 'Heartbeat') { 15 } else { 0 }) -HeartbeatSeconds $(if ($Mode -eq 'Heartbeat') { 1 } else { 15 })
        return @{ Code = $code; State = $state; Plan = $plan; Directory = $directory; Sim = $sim }
    } finally { if ($null -ne $sim.Lock) { $sim.Lock.Dispose() } }
}
function Test-ReceiptFixture {
    param([string] $Phase, [bool] $Saved, [int] $ExitCode, [switch] $BrokenConsoleLog)
    $directory = Join-Path $testRoot ([guid]::NewGuid().ToString('N'))
    [void] [IO.Directory]::CreateDirectory($directory)
    Write-BatchJson (Join-Path $directory 'refresh-worker-fixture.json') @{ phase = $Phase; saved = $Saved }
    if ($BrokenConsoleLog) { [void] [IO.Directory]::CreateDirectory((Join-Path $directory 'stdout.log')) }
    $info = [Diagnostics.ProcessStartInfo]::new($env:ComSpec)
    $info.Arguments = "/d /c exit $ExitCode"
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    $handle = @{ Process = $process; Stdout = $process.StandardOutput.ReadToEndAsync(); Stderr = $process.StandardError.ReadToEndAsync(); Directory = $directory }
    try {
        if (-not $process.WaitForExit(3000)) { throw 'Synthetic receipt process did not exit.' }
        $receipt = Receive-BatchWorkbookProcess $handle
        $again = Receive-BatchWorkbookProcess $handle
        Assert-Recovery ([object]::ReferenceEquals($receipt, $again)) 'completed receipt is cached before disposing its process'
        return $receipt
    } finally { $process.Dispose() }
}
try {
    # Reproduce the original un-retried Move failure on actual Windows handles,
    # then prove retry succeeds after > the old 200 ms window for two denial causes.
    foreach ($readOnly in @($false, $true)) {
        $path = Join-Path $testRoot "atomic-$readOnly.json"
        Write-BatchJson $path @{ Version = 'old' }
        $process = Start-StatusLock $path -ReadOnly:$readOnly
        $temp = "$path.original.tmp"
        [IO.File]::WriteAllText($temp, '{"Version":"unpublished"}')
        $denied = $null
        try { [IO.File]::Move($temp, $path, $true) } catch { $denied = $_.Exception.GetBaseException() }
        Assert-Recovery ($null -ne $denied) 'original atomic replacement fails while the target is held'
        Write-Host "Reproduced: $($denied.GetType().Name), Windows code $($denied.HResult -band 0xffff)."
        Assert-Recovery ((Read-BatchJson $path).Version -eq 'old') 'original published JSON is intact after denial'
        $timer = [Diagnostics.Stopwatch]::StartNew()
        Write-RefreshJson $path @{ Version = 'new' } -RetryMilliseconds 5000
        Assert-Recovery ($timer.ElapsedMilliseconds -gt 200) 'test exercised a lock longer than the old retry window'
        Assert-Recovery ((Read-BatchJson $path).Version -eq 'new') 'atomic publication succeeds after transient denial clears'
        Assert-Recovery ($process.WaitForExit(2000) -and $process.ExitCode -eq 0) 'lock holder exits normally'
    }

    # Persistent lock: no truncation, no access bypass, bounded failure, useful evidence.
    $path = Join-Path $testRoot 'persistent.json'
    Write-BatchJson $path @{ Version = 'old' }
    $lock = [IO.File]::Open($path, 'Open', 'Read', [IO.FileShare]::ReadWrite)
    $failure = $null; $timer = [Diagnostics.Stopwatch]::StartNew()
    try { Write-RefreshJson $path @{ Version = 'new' } -RetryMilliseconds 200 } catch { $failure = $_ }
    finally { $lock.Dispose() }
    Assert-Recovery ($null -ne $failure -and $timer.Elapsed.TotalSeconds -lt 3) 'permanent lock fails within its configured deadline'
    Assert-Recovery ($failure.Exception.Message.Contains($path) -and $failure.Exception.Message -match 'Windows code|recovery snapshot') 'error identifies the exact path and recovery evidence'
    Assert-Recovery ((Read-BatchJson $path).Version -eq 'old') 'permanent failure preserves published JSON'
    Assert-Recovery ((Read-BatchJson $failure.Exception.Data['RecoveryPath']).Version -eq 'new') 'prepared state survives for inspection'

    # The shared supervised worker now uses the same hardened publication path.
    $WorkerStatePath = Join-Path $testRoot 'worker.json'
    $script:RefreshState = @{ phase = 'Refreshing'; saved = $false; updatedAt = '' }
    Write-RefreshWorkerState
    $null = Start-StatusLock $WorkerStatePath
    $script:RefreshState.phase = 'Complete'; $script:RefreshState.saved = $true
    Write-RefreshWorkerState
    Assert-Recovery ((Read-RefreshWorkerState $WorkerStatePath).phase -eq 'Complete') 'worker receipt tolerates a transient reader too'

    $result = Invoke-RecoveryFixture Transient
    Assert-Recovery ($result.Code -eq 0 -and @($result.State.Jobs.Values | Where-Object Status -ne 'Completed').Count -eq 0) 'live batch scheduling survives transient status locks'
    Assert-Recovery (($result.Sim.Starts -join ',') -eq 'file1,file2,file3') 'every selected synthetic file dispatches exactly once in deterministic order'
    Assert-Recovery (@(Get-ChildItem $result.Directory -Filter stop.txt -Recurse).Count -eq 0) 'transient state locks do not stop workers'
    Assert-Recovery ((Get-Content (Join-Path $result.Directory 'coordinator.log') -Raw) -match 'file3 \| Completed') 'readable live log records progress and completion'

    $result = Invoke-RecoveryFixture Permanent
    Assert-Recovery ($result.Code -eq 1 -and $result.State.Status -eq 'PartialOrFailed') 'coordinator failure is not falsely reported as operator stop'
    Assert-Recovery ($result.State.Jobs.file1.Status -eq 'Completed' -and $result.State.Jobs.file1.ExitCode -eq 0 -and $result.State.Jobs.file1.OutputStamp -and $result.State.Jobs.file1.Finished) 'completion racing coordinator failure retains save metadata and exit code'
    Assert-Recovery ($result.State.Jobs.file2.Status -eq 'Stopped' -and $result.State.Jobs.file2.ExitCode -eq 3 -and $result.State.Jobs.file2.Finished) 'unsaved worker is stopped with its actual receipt'
    Assert-Recovery ($result.State.Jobs.file3.Status -eq 'Blocked' -and 'file3' -notin $result.Sim.Starts) 'coordinator failure never dispatches pending descendants'
    $errors = @(Get-ChildItem $result.Directory -Filter 'coordinator-error-*.json')
    Assert-Recovery ($errors.Count -eq 1 -and (Read-BatchJson $errors[0].FullName).Message -eq $result.State.CoordinatorError.Message) 'original exception survives final publication failure'
    $recovery = @(Get-ChildItem $result.Directory -Filter 'state-recovery-*.json')
    $recovered = Read-BatchJson $recovery[0].FullName -AsHashtable
    Assert-Recovery ($recovery.Count -eq 1 -and $recovered.ExitCode -eq 1 -and $recovered.Jobs.file1.Status -eq 'Completed') 'separate final recovery record preserves accurate terminal state'
    Reset-BatchResume $result.Plan $result.State
    Assert-Recovery ($result.State.Jobs.file1.Status -eq 'Completed' -and $result.State.Jobs.file2.Status -eq 'Pending' -and -not $result.State.ContainsKey('CoordinatorError')) 'metadata-based reset preserves confirmed work and clears previous error only in memory'

    $result = Invoke-RecoveryFixture BeforeDispatch
    Assert-Recovery ($result.Sim.Starts.Count -eq 0 -and $result.Code -eq 1) 'failed initial publication starts NO worker'
    Assert-Recovery (-not $result.State.Jobs.file1.NeedsInspection -and $result.State.Jobs.file1.Status -eq 'Blocked') 'never-started job is not mislabelled as an uncertain save'
    Assert-Recovery ($result.State.CoordinatorError.Message -like 'Original failure:*') 'coordinator cause is persisted when a later state write succeeds'

    $result = Invoke-RecoveryFixture StopDuringPublish
    Assert-Recovery ($result.Code -eq 3 -and $result.Sim.Starts.Count -eq 0) 'stop arriving during status publication prevents the next dispatch'

    $result = Invoke-RecoveryFixture Slow
    Assert-Recovery ($result.Code -eq 0 -and $result.Sim.Polls -ge 300 -and $result.Sim.Writes -lt 12) 'unchanged polling no longer rewrites the ledger every 100 ms'
    $quietWrites = $result.Sim.Writes
    $result = Invoke-RecoveryFixture Heartbeat
    Assert-Recovery ($result.Code -eq 0 -and $result.Sim.Writes -gt $quietWrites) 'quiet active runs still publish bounded heartbeats'

    # Missing/ambiguous save confirmation remains unsafe even if a process exits 0.
    $job = $result.Plan.Jobs[0]; $record = $result.State.Jobs[$job.Id]
    Complete-BatchJob $job $record @{ ExitCode = 0; NeedsInspection = $true; Message = 'Uncertain save' }
    Assert-Recovery ($record.Status -eq 'Failed' -and $record.ExitCode -eq 1) 'uncertain save never becomes successful completion'
    $rejected = $false
    try { Reset-BatchResume $result.Plan $result.State } catch { $rejected = $true }
    Assert-Recovery $rejected 'uncertain save cannot be automatically retried'

    $receipt = Test-ReceiptFixture -Phase Complete -Saved $true -ExitCode 0 -BrokenConsoleLog
    Assert-Recovery ($receipt.ExitCode -eq 0 -and -not $receipt.NeedsInspection) 'failed console-log copy does not erase a verified save/close receipt'
    $receipt = Test-ReceiptFixture -Phase Quitting -Saved $true -ExitCode 1
    Assert-Recovery ($receipt.NeedsInspection -and $receipt.ExitCode -eq 1) 'saved but unconfirmed close/cleanup requires inspection'
    $receipt = Test-ReceiptFixture -Phase Saving -Saved $false -ExitCode 3
    Assert-Recovery ($receipt.NeedsInspection -and $receipt.ExitCode -eq 3) 'stop during save remains uncertain'
    $receipt = Test-ReceiptFixture -Phase CheckingReadiness -Saved $false -ExitCode 3
    Assert-Recovery (-not $receipt.NeedsInspection -and $receipt.ExitCode -eq 3) 'confirmed pre-save stop remains distinct from uncertain save'
    $receipt = Test-ReceiptFixture -Phase Refreshing -Saved $false -ExitCode 0
    Assert-Recovery ($receipt.NeedsInspection -and $receipt.ExitCode -eq 1) 'exit zero without a complete save/close receipt is not success'
    Write-Host "PASS: $script:checks recovery checks; real Windows lock reproduction, safe retries, completion recovery and no Excel execution."
} finally {
    foreach ($process in $helpers) {
        if (-not $process.WaitForExit(3000)) { $process.Kill($true); $process.WaitForExit() }
        $process.Dispose()
    }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-recovery-*') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
