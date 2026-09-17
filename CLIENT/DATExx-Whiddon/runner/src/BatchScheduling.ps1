Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'RefreshJson.ps1')

function Read-BatchJson {
    param([string] $Path, [switch] $AsHashtable)
    # Readers permit atomic replacement by the coordinator/stop requester.
    $stream = [IO.File]::Open($Path, 'Open', 'Read', ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    $reader = [IO.StreamReader]::new($stream)
    try { return ($reader.ReadToEnd() | ConvertFrom-Json -AsHashtable:$AsHashtable) }
    finally { $reader.Dispose() }
}

function Write-BatchJson {
    param([string] $Path, $Value)
    Write-RefreshJson $Path $Value
}

function Write-BatchProgress {
    param([string] $RunDirectory, [string] $Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line
    # A text-log viewer must not become another reason to interrupt a refresh.
    try {
        $stream = [IO.File]::Open((Join-Path $RunDirectory 'coordinator.log'), 'Append', 'Write',
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $writer = [IO.StreamWriter]::new($stream)
        try { $writer.WriteLine($line) } finally { $writer.Dispose() }
    } catch { Write-Warning "Could not append coordinator.log: $($_.Exception.Message)" }
}

function Write-BatchCoordinatorError {
    param([string] $RunDirectory, $ErrorRecord)
    $diagnostic = @{
        Time = [datetime]::UtcNow.ToString('o'); Message = $ErrorRecord.Exception.Message
        Exception = $ErrorRecord.Exception.ToString(); ScriptStack = $ErrorRecord.ScriptStackTrace
        Location = $ErrorRecord.InvocationInfo.PositionMessage
    }
    $name = 'coordinator-error-' + [guid]::NewGuid().ToString('N') + '.json'
    $diagnostic.Record = $name
    try { Write-RefreshJson (Join-Path $RunDirectory $name) $diagnostic -RetryMilliseconds 1000 }
    catch { Write-Warning "Could not persist coordinator error: $($_.Exception.Message)" }
    Write-BatchProgress $RunDirectory "COORDINATOR ERROR: $($diagnostic.Message) | Details: $name"
    return $diagnostic
}

function New-BatchState {
    param($Plan, [string] $RunId)
    $records = @{}
    foreach ($job in $Plan.Jobs) {
        $records[$job.Id] = @{
            Status = 'Pending'; Message = ''; Attempt = 0; Started = ''; Finished = ''; ExitCode = $null
            OutputStamp = ''; Inputs = @{}; NeedsInspection = $false; Directory = ''
        }
    }
    return @{
        SchemaVersion = 1; RunId = $RunId; Fingerprint = $Plan.Fingerprint; Status = 'Pending'
        Started = [datetime]::UtcNow.ToString('o'); Updated = ''; Jobs = $records; Batches = @{}
        MaxObserved = 0; ExitCode = $null; BusinessReconciliation = 'Not evaluated'
    }
}

function Update-BatchSummary {
    param($Plan, [hashtable] $State)
    $State.Batches = @{}
    foreach ($group in ($Plan.Jobs | Group-Object BatchKey)) {
        $statuses = @($group.Group | ForEach-Object { $State.Jobs[$_.Id].Status })
        $status = if ($statuses -contains 'Running') { 'Running' }
            elseif ($statuses -contains 'Failed') { 'Failed' }
            elseif ($statuses -contains 'Blocked') { 'Blocked' }
            elseif ($statuses -contains 'Stopped') { 'Stopped' }
            elseif (@($statuses | Where-Object { $_ -ne 'Completed' }).Count -eq 0) { 'Completed' }
            else { 'Pending' }
        $State.Batches[$group.Name] = @{ Status = $status; Completed = @($statuses | Where-Object { $_ -eq 'Completed' }).Count; Total = $statuses.Count }
    }
    $State.Updated = [datetime]::UtcNow.ToString('o')
}

function Reset-BatchResume {
    param($Plan, [hashtable] $State)
    if ($State.Fingerprint -ne $Plan.Fingerprint -or $State.SchemaVersion -ne 1) { throw 'Configuration changed: create a new plan instead of resuming.' }
    if ($State.Jobs.Count -ne $Plan.Jobs.Count) { throw 'Saved selection does not match the run state.' }
    foreach ($job in $Plan.Jobs) {
        if (-not $State.Jobs.ContainsKey($job.Id)) { throw "Saved state is missing $($job.Id)." }
        $record = $State.Jobs[$job.Id]
        if ($record.NeedsInspection -or $record.Status -eq 'Running') {
            throw "Inspection required for $($job.Id): interrupted or uncertain save. Do not automatically retry; verify the original and create a new targeted run."
        }
        if ($record.Status -eq 'Completed') {
            $changed = (Get-BatchFileStamp $job.Path) -ne $record.OutputStamp
            foreach ($path in $record.Inputs.Keys) {
                if ((Get-BatchFileStamp $path) -ne $record.Inputs[$path]) { $changed = $true }
            }
            if ($changed) { $record.Status = 'Pending'; $record.Message = 'Invalidated: input/output metadata changed.' }
        } else { $record.Status = 'Pending'; $record.Message = '' }
    }
    do {
        $changed = $false
        foreach ($job in $Plan.Jobs) {
            if ($State.Jobs[$job.Id].Status -ne 'Completed') { continue }
            foreach ($dep in $job.Dependencies) {
                if ($State.Jobs.ContainsKey($dep) -and $State.Jobs[$dep].Status -ne 'Completed') {
                    $State.Jobs[$job.Id].Status = 'Pending'; $State.Jobs[$job.Id].Message = "Invalidated by $dep"; $changed = $true; break
                }
            }
        }
    } while ($changed)
    $State.Status = 'Pending'; $State.ExitCode = $null
    $State.Remove('CoordinatorError'); $State.Remove('FinalStateError')
}

function Test-BatchConflict {
    param($Job, [object[]] $ActiveJobs)
    foreach ($active in $ActiveJobs) {
        if ($active.BatchKey -eq $Job.BatchKey -or $active.Exclusive -or $Job.Exclusive -or
            $active.Path -eq $Job.Path -or $active.Path -in $Job.Reads -or $Job.Path -in $active.Reads) { return $true }
    }
    return $false
}

function Get-BatchStopMode {
    param([string] $Path, [string] $CurrentMode)
    if (-not (Test-Path -LiteralPath $Path)) { return $CurrentMode }
    $request = Read-BatchJson $Path
    if ($request.Mode -notin @('AfterCurrent', 'AfterBatch', 'Now')) { throw 'Invalid stop request.' }
    $rank = @{ '' = 0; AfterBatch = 1; AfterCurrent = 2; Now = 3 }
    if ($rank[$request.Mode] -gt $rank[$CurrentMode]) { return $request.Mode }
    return $CurrentMode
}

function Start-BatchWorkbookProcess {
    param($Job, [string] $Directory, $Options)
    $inputStamps = @{}
    if ($Options.ContainsKey('InputStamps')) { $inputStamps = $Options.InputStamps }
    else { foreach ($path in $Job.Reads) { $inputStamps[$path] = Get-BatchFileStamp $path } }
    $inputSnapshotPath = Join-Path $Directory 'input-snapshot.json'
    Write-BatchJson $inputSnapshotPath @{ SchemaVersion = 1; Target = $Job.Path; Inputs = $inputStamps }
    $request = @{
        Id = $Job.Id; Path = $Job.Path; DateRoot = $Options.DateRoot
        ParentId = $PID; ParentStarted = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks
        LogPath = (Join-Path $Directory 'refresh.log'); StatusPath = (Join-Path $Directory 'status.txt')
        StopPath = (Join-Path $Directory 'stop.txt'); Visible = $Options.Visible; TimeoutMinutes = $Options.TimeoutMinutes
        InputSnapshotPath = $inputSnapshotPath
    }
    $requestPath = Join-Path $Directory 'request.json'
    Write-BatchJson $requestPath $request
    $info = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh).Source)
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    foreach ($arg in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'Invoke-BatchWorkbook.ps1'), '-RequestPath', $requestPath)) { $info.ArgumentList.Add($arg) }
    $process = [Diagnostics.Process]::Start($info)
    return @{ Process = $process; Stdout = $process.StandardOutput.ReadToEndAsync(); Stderr = $process.StandardError.ReadToEndAsync(); Directory = $Directory }
}

function Receive-BatchWorkbookProcess {
    param($Handle)
    if ($Handle.ContainsKey('Receipt')) { return $Handle.Receipt }
    if (-not $Handle.Process.HasExited) { return $null }
    $code = $Handle.Process.ExitCode
    $confirmed = $false; $uncertain = $true; $message = "Worker exit $code"
    try {
        $states = @(Get-ChildItem -LiteralPath $Handle.Directory -Filter 'refresh-worker-*.json' -File)
        if ($states.Count -eq 1) {
            $worker = Read-BatchJson $states[0].FullName
            $confirmed = $worker.phase -eq 'Complete' -and $worker.saved
            $uncertain = ($worker.saved -and ($code -ne 0 -or -not $confirmed)) -or
                (-not $worker.saved -and $worker.phase -in @('Saving', 'RestoringBackgroundRefresh', 'Closing', 'Quitting'))
            $message += "; phase $($worker.phase); save confirmed: $($worker.saved)"
            if ($null -ne $worker.PSObject.Properties['error']) { $message += "; $($worker.error.Message)" }
        }
    } catch { $message += "; receipt unreadable: $($_.Exception.Message)" }
    if ($code -eq 0 -and -not $confirmed) { $code = 1; $uncertain = $true; $message += '; no confirmed save/close receipt.' }
    foreach ($entry in @(@('Stdout', 'stdout.log'), @('Stderr', 'stderr.log'))) {
        try { [IO.File]::WriteAllText((Join-Path $Handle.Directory $entry[1]), $Handle[$entry[0]].GetAwaiter().GetResult()) }
        catch { Write-Warning "Worker console log unavailable: $($_.Exception.Message)" }
    }
    $Handle.Receipt = @{ ExitCode = $code; NeedsInspection = $uncertain; Message = $message }
    $Handle.Process.Dispose()
    return $Handle.Receipt
}

function Complete-BatchJob {
    param($Job, [hashtable] $Record, $Receipt)
    $Record.ExitCode = $Receipt.ExitCode; $Record.NeedsInspection = $Receipt.NeedsInspection
    $Record.Message = $Receipt.Message; $Record.Finished = [datetime]::UtcNow.ToString('o')
    $Record.Status = if ($Receipt.ExitCode -eq 0 -and -not $Receipt.NeedsInspection) { 'Completed' }
        elseif ($Receipt.ExitCode -eq 3) { 'Stopped' } else { 'Failed' }
    if ($Record.Status -eq 'Failed' -and $Record.ExitCode -eq 0) { $Record.ExitCode = 1 }
    if ($Record.Status -eq 'Completed') {
        $Record.OutputStamp = Get-BatchFileStamp $Job.Path
        foreach ($path in $Record.Inputs.Keys) {
            if ((Get-BatchFileStamp $path) -ne $Record.Inputs[$path]) {
                $Record.Status = 'Failed'; $Record.ExitCode = 1; $Record.Message = "Input changed during refresh: $path"; break
            }
        }
    }
}

function Invoke-BatchSchedule {
    param($Plan, [hashtable] $State, [string] $RunDirectory,
        [ValidateRange(1, 7)] [int] $MaxParallelBatches = 7,
        [string] $Visible = 'false', [double] $TimeoutMinutes = 30,
        [scriptblock] $StartWorker = ${function:Start-BatchWorkbookProcess},
        [scriptblock] $PollWorker = ${function:Receive-BatchWorkbookProcess},
        [scriptblock] $ValidateJob = { param($job)
            Assert-BatchFileAccessible $job.Path -OutputFile -InspectFileUsers -FileUserWaitSeconds 15
            foreach ($path in $job.Reads) { Assert-BatchFileAccessible $path }
        },
        [int] $PollMilliseconds = 100,
        [scriptblock] $PersistState = { param($path, $value) Write-BatchJson $path $value },
        [ValidateRange(1, 3600)] [int] $HeartbeatSeconds = 15)
    $active = @{}; $jobs = @{}; $admitted = @{}; $drainBatches = $null; $stopMode = ''
    foreach ($job in $Plan.Jobs) { $jobs[$job.Id] = $job }
    $statePath = Join-Path $RunDirectory 'state.json'
    $stopPath = Join-Path $RunDirectory 'stop.json'
    $State.Status = 'Running'; $State.MaxParallelBatches = $MaxParallelBatches
    $State.Coordinator = @{ Id = $PID; Started = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks }
    $options = @{ DateRoot = $Plan.DateRoot; Visible = $Visible; TimeoutMinutes = $TimeoutMinutes }
    $coordinatorFailure = $null
    $persistence = @{ Snapshot = ''; Timer = [Diagnostics.Stopwatch]::StartNew() }
    $saveState = {
        param([switch] $Force)
        $snapshot = $State | ConvertTo-Json -Depth 40 -Compress
        if ($Force -or $snapshot -ne $persistence.Snapshot -or $persistence.Timer.Elapsed.TotalSeconds -ge $HeartbeatSeconds) {
            Update-BatchSummary $Plan $State
            & $PersistState $statePath $State
            $persistence.Snapshot = $State | ConvertTo-Json -Depth 40 -Compress
            $persistence.Timer.Restart()
        }
    }
    Write-BatchProgress $RunDirectory "Run $($State.RunId) started: $($Plan.Jobs.Count) files; up to $MaxParallelBatches active batches."
    try {
        while ($true) {
            $nextStopMode = Get-BatchStopMode $stopPath $stopMode
            if ($nextStopMode -ne $stopMode) {
                $stopMode = $nextStopMode
                Write-BatchProgress $RunDirectory "Operator stop requested: $stopMode"
            }
            if ($stopMode -eq 'AfterBatch' -and $null -eq $drainBatches) {
                $drainBatches = @{}; foreach ($key in $admitted.Keys) { $drainBatches[$key] = $true }
            }
            foreach ($id in @($active.Keys)) {
                if ($stopMode -eq 'Now') { [IO.File]::WriteAllText((Join-Path $active[$id].Directory 'stop.txt'), "mode=Now`n") }
                $receipt = & $PollWorker $active[$id].Handle
                if ($null -eq $receipt) { continue }
                $record = $State.Jobs[$id]
                Complete-BatchJob $jobs[$id] $record $receipt
                Write-BatchProgress $RunDirectory "$id | $($record.Status) | $($record.Message)"
                $active.Remove($id)
            }
            # Failed/stopped selected prerequisites never fall back to old saved files.
            do {
                $blocked = $false
                foreach ($job in $Plan.Jobs) {
                    $record = $State.Jobs[$job.Id]
                    if ($record.Status -ne 'Pending') { continue }
                    foreach ($dep in $job.Dependencies) {
                        if ($jobs.ContainsKey($dep) -and $State.Jobs[$dep].Status -in @('Failed', 'Blocked', 'Stopped')) {
                            $record.Status = 'Blocked'; $record.Message = "Prerequisite $dep did not complete."; $blocked = $true
                            Write-BatchProgress $RunDirectory "$($job.Id) | Blocked | $($record.Message)"
                            break
                        }
                    }
                }
            } while ($blocked)
            foreach ($key in @($admitted.Keys)) {
                $remaining = @($Plan.Jobs | Where-Object { $_.BatchKey -eq $key -and $State.Jobs[$_.Id].Status -in @('Pending', 'Running') })
                if (-not $remaining.Count) { $admitted.Remove($key) }
            }
            foreach ($job in $Plan.Jobs) {
                if ($active.Count -ge $MaxParallelBatches -or $stopMode -in @('AfterCurrent', 'Now')) { break }
                $record = $State.Jobs[$job.Id]
                if ($record.Status -ne 'Pending') { continue }
                if ($null -ne $drainBatches -and -not $drainBatches.ContainsKey($job.BatchKey)) { continue }
                $waiting = @($job.Dependencies | Where-Object { $jobs.ContainsKey($_) -and $State.Jobs[$_].Status -ne 'Completed' })
                if ($waiting.Count) { $record.Message = "Waiting for $($waiting -join ', ')"; continue }
                if (Test-BatchConflict $job @($active.Keys | ForEach-Object { $jobs[$_] })) { continue }
                $publishingStart = $false
                try {
                    foreach ($dep in $job.Dependencies) {
                        if ($jobs.ContainsKey($dep) -and $State.Jobs[$dep].Status -eq 'Completed' -and
                            (Get-BatchFileStamp $jobs[$dep].Path) -ne $State.Jobs[$dep].OutputStamp) {
                            $State.Jobs[$dep].Status = 'Failed'; $State.Jobs[$dep].ExitCode = 1
                            $State.Jobs[$dep].Message = 'Output changed after confirmed completion; new inspection/selection required.'
                            throw "Selected prerequisite $dep changed after completion."
                        }
                    }
                    & $ValidateJob $job
                    $record.Inputs = @{}; foreach ($path in $job.Reads) { $record.Inputs[$path] = Get-BatchFileStamp $path }
                    $record.Attempt++; $record.Started = [datetime]::UtcNow.ToString('o'); $record.Status = 'Running'; $record.Message = ''
                    $index = [array]::IndexOf($Plan.Jobs, $job)
                    $directory = Join-Path $RunDirectory ('job-{0:d3}/attempt-{1}' -f $index, $record.Attempt)
                    [void] [IO.Directory]::CreateDirectory($directory)
                    $record.Directory = [IO.Path]::GetRelativePath($RunDirectory, $directory)
                    # Commit the running record before process launch: a coordinator crash
                    # cannot leave an apparently safe-to-retry job that may have saved.
                    $publishingStart = $true
                    & $saveState -Force
                    # Publication can wait for a reader. Honor a stop that arrived
                    # during that wait BEFORE launching the next workbook.
                    $nextStopMode = Get-BatchStopMode $stopPath $stopMode
                    if ($nextStopMode -ne $stopMode) {
                        $stopMode = $nextStopMode
                        Write-BatchProgress $RunDirectory "Operator stop requested: $stopMode"
                    }
                    if ($stopMode -eq 'AfterBatch' -and $null -eq $drainBatches) {
                        $drainBatches = @{}; foreach ($key in $admitted.Keys) { $drainBatches[$key] = $true }
                    }
                    if ($stopMode -in @('AfterCurrent', 'Now') -or
                        ($null -ne $drainBatches -and -not $drainBatches.ContainsKey($job.BatchKey))) {
                        $record.Status = 'Pending'; $record.Started = ''; $record.Message = "Operator stop: $stopMode"
                        continue
                    }
                    $publishingStart = $false
                    $jobOptions = @{} + $options
                    $jobOptions.InputStamps = @{} + $record.Inputs
                    $handle = & $StartWorker $job $directory $jobOptions
                    $active[$job.Id] = @{ Handle = $handle; Directory = $directory }
                    $admitted[$job.BatchKey] = $true
                    $State.MaxObserved = [Math]::Max($State.MaxObserved, $active.Count)
                    Write-BatchProgress $RunDirectory "$($job.Id) | Running | $($job.Path)"
                }
                catch {
                    if ($publishingStart) {
                        # No worker was dispatched. This is a coordinator storage
                        # failure, not a workbook failure or an uncertain save.
                        $record.Status = 'Pending'; $record.Started = ''; $record.NeedsInspection = $false
                        throw
                    }
                    $record.NeedsInspection = $record.Status -eq 'Running'
                    $record.Status = if ($record.NeedsInspection) { 'Failed' } else { 'Blocked' }
                    $record.ExitCode = 1; $record.Message = $_.Exception.Message
                    Write-BatchProgress $RunDirectory "$($job.Id) | $($record.Status) | $($record.Message)"
                }
            }
            & $saveState
            if ($active.Count -eq 0) {
                foreach ($record in $State.Jobs.Values) {
                    if ($record.Status -eq 'Pending') {
                        $record.Status = if ($stopMode) { 'Stopped' } else { 'Blocked' }
                        if (-not $record.Message) { $record.Message = if ($stopMode) { "Operator stop: $stopMode" } else { 'No eligible work remains.' } }
                    }
                }
                break
            }
            if ($PollMilliseconds -gt 0) { Start-Sleep -Milliseconds $PollMilliseconds }
        }
    }
    catch {
        $coordinatorFailure = $_
        $State.CoordinatorError = Write-BatchCoordinatorError $RunDirectory $_
    }
    finally {
        if ($active.Count) {
            # Keep the Date lease until each independent supervisor has stopped.
            foreach ($item in $active.Values) { [IO.File]::WriteAllText((Join-Path $item.Directory 'stop.txt'), "mode=Now`n") }
            while ($active.Count) {
                foreach ($id in @($active.Keys)) {
                    $receipt = & $PollWorker $active[$id].Handle
                    if ($null -eq $receipt) { continue }
                    # A save/close may already have completed when interruption
                    # begins. Preserve its receipt instead of overwriting success.
                    $record = $State.Jobs[$id]
                    Complete-BatchJob $jobs[$id] $record $receipt
                    if ($record.Status -eq 'Stopped') { $record.Message = "Coordinator interrupted; $($record.Message)" }
                    Write-BatchProgress $RunDirectory "$id | $($record.Status) | $($record.Message)"
                    $active.Remove($id)
                }
                if ($active.Count) { Start-Sleep -Milliseconds 100 }
            }
        }
        if ($null -ne $coordinatorFailure) {
            foreach ($record in $State.Jobs.Values) {
                if ($record.Status -eq 'Pending') {
                    $record.Status = 'Blocked'; $record.Message = "Coordinator failed: $($State.CoordinatorError.Message)"
                }
            }
        }
        $states = @($State.Jobs.Values | ForEach-Object Status)
        $State.ExitCode = if ($null -ne $coordinatorFailure) { 1 } elseif ($states -contains 'Stopped') { 3 } elseif ($states -contains 'Failed' -or $states -contains 'Blocked') { 1 } elseif ($states -contains 'Pending' -or $states -contains 'Running') { 1 } else { 0 }
        if ($null -eq $coordinatorFailure -and $State.ExitCode -eq 1 -and @($State.Jobs.Values | Where-Object ExitCode -eq 124).Count) { $State.ExitCode = 124 }
        $State.Status = if ($State.ExitCode -eq 0) { 'Completed' } elseif ($State.ExitCode -eq 3) { 'Stopped' } else { 'PartialOrFailed' }
        try { & $saveState -Force }
        catch {
            # Preserve the FIRST error even if the final state is still locked.
            $State.FinalStateError = $_.Exception.Message
            if ($null -eq $coordinatorFailure) { $State.CoordinatorError = Write-BatchCoordinatorError $RunDirectory $_ }
            $State.ExitCode = 1; $State.Status = 'PartialOrFailed'
            $recoveryPath = Join-Path $RunDirectory ('state-recovery-' + [guid]::NewGuid().ToString('N') + '.json')
            try {
                Update-BatchSummary $Plan $State
                Write-RefreshJson $recoveryPath $State -RetryMilliseconds 1000
                Write-BatchProgress $RunDirectory "Final state could not be published. Inspect $recoveryPath; do not automatically resume the stale state.json."
            } catch { Write-Warning "Final recovery snapshot could not be written: $($_.Exception.Message)" }
        }
        Write-BatchProgress $RunDirectory "Run $($State.RunId): $($State.Status); exit $($State.ExitCode)."
    }
    return $State.ExitCode
}
