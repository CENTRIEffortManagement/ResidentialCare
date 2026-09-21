Set-StrictMode -Version Latest

function Get-BatchLocationIdentity {
    param([string] $DateRoot)
    $location = [Environment]::MachineName + '|' + [IO.Path]::GetFullPath($DateRoot).TrimEnd('\','/').ToLowerInvariant()
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($location)))
}

function ConvertTo-PortableBatchPlan {
    param($Plan)
    # Paths used by workers are reconstructed locally; the witness prevents
    # accidentally resuming an old attempt on a different machine or checkout.
    $portable = $Plan | ConvertTo-Json -Depth 40 | ConvertFrom-Json
    $portable | Add-Member LocationIdentity (Get-BatchLocationIdentity $Plan.DateRoot) -Force
    foreach ($job in $portable.Jobs) {
        $job | Add-Member ReadFiles @($job.Reads | ForEach-Object {
            $relative = [IO.Path]::GetRelativePath($Plan.DateRoot, $_)
            if ($relative -ne '..' -and -not $relative.StartsWith('..' + [IO.Path]::DirectorySeparatorChar) -and -not [IO.Path]::IsPathRooted($relative)) {
                [ordered]@{ Root = 'Date'; Path = $relative.Replace('\','/') }
            } else {
                # Known environment roots are portable. Unknown external roots
                # keep an opaque identity and are never rebound on another host.
                $publicRoot = [Environment]::GetEnvironmentVariable('PUBLIC')
                if ($publicRoot -and $_.StartsWith($publicRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
                    [ordered]@{ Root = 'PUBLIC'; Path = [IO.Path]::GetRelativePath($publicRoot, $_).Replace('\','/') }
                } else { [ordered]@{ Root = 'External'; Identity = Get-BatchLocationIdentity $_ } }
            }
        }) -Force
        $job.PSObject.Properties.Remove('Path')
        $job.PSObject.Properties.Remove('Reads')
    }
    $portable.PSObject.Properties.Remove('DateRoot')
    return $portable
}

function Get-BatchDurationHistory {
    param([string] $LogRoot, $Plan = $null)
    $durations = @{}
    if (-not (Test-Path -LiteralPath $LogRoot)) { return $durations }
    foreach ($directory in (Get-ChildItem -LiteralPath $LogRoot -Directory)) {
        $statePath = Join-Path $directory.FullName 'state.json'
        if (-not (Test-Path -LiteralPath $statePath) -or @(Get-ChildItem -LiteralPath $directory.FullName -Filter 'state-recovery-*.json').Count) { continue }
        try {
            $state = Read-BatchJson $statePath -AsHashtable
            $historicalPlanPath = Join-Path $directory.FullName 'plan.json'
            if (-not (Test-Path -LiteralPath $historicalPlanPath)) { continue }
            $historicalPlan = Read-BatchJson $historicalPlanPath
            foreach ($id in $state.Jobs.Keys) {
                if ($null -ne $Plan) {
                    $target = @($Plan.Jobs | Where-Object Id -eq $id)
                    $source = @($historicalPlan.Jobs | Where-Object Id -eq $id)
                    if ($target.Count -ne 1 -or $source.Count -ne 1 -or $target[0].RelativePath -ne $source[0].RelativePath) { continue }
                }
                $record = $state.Jobs[$id]
                if ($record.Status -ne 'Completed' -or $record.NeedsInspection -or -not $record.Started -or -not $record.Finished) { continue }
                $start = [datetimeoffset]$record.Started
                $finish = [datetimeoffset]$record.Finished
                $seconds = ($finish - $start).TotalSeconds
                if ($seconds -le 0) { continue }
                if (-not $durations.ContainsKey($id) -or $finish -gt [datetimeoffset]::Parse($durations[$id].Observed)) {
                    $durations[$id] = @{ Seconds = $seconds; SourceRun = $directory.Name; Observed = $finish.ToString('o') }
                }
            }
        } catch { Write-Warning "History unavailable for $($directory.Name): $($_.Exception.Message)" }
    }
    return $durations
}

function Get-BatchForecast {
    param($Plan, [hashtable] $Durations, [ValidateRange(1,12)] [int] $Limit = 12)
    $status = @{}; $active = @{}; $ends = @{}; $rows = [Collections.Generic.List[object]]::new()
    foreach ($job in $Plan.Jobs) { $status[$job.Id] = 'Pending' }
    $time = 0.0; $peak = 0
    while ($true) {
        foreach ($id in @($active.Keys)) {
            if ($ends[$id] -le $time) { $active.Remove($id); $status[$id] = 'Completed' }
        }
        foreach ($job in $Plan.Jobs) {
            if (-not (Test-BatchAdmission $job $status @($active.Values) $Limit)) { continue }
            if (-not $Durations.ContainsKey($job.Id)) {
                # Stop at the first unknown admission: its slot occupancy could
                # alter all later starts. Keep the known prefix trustworthy.
                return @{ Rows = @($rows); Complete = $false; TotalSeconds = $null; Peak = $peak
                    UnknownAdmission = $job.Id; Durations = $Durations }
            }
            $seconds = [double]$Durations[$job.Id].Seconds
            if ($seconds -le 0) { throw "Invalid duration for $($job.Id)" }
            $ends[$job.Id] = $time + $seconds
            $active[$job.Id] = $job; $status[$job.Id] = 'Running'
            $peak = [Math]::Max($peak, $active.Count)
            $rows.Add(@{ Id = $job.Id; Start = $time; Finish = $ends[$job.Id]; Status = 'Forecast'; Attempt = 1 })
        }
        if (-not $active.Count) { break }
        $time = [double](@($active.Keys | ForEach-Object { $ends[$_] } | Sort-Object)[0])
    }
    return @{ Rows = @($rows); Complete = (@($status.Values | Where-Object { $_ -ne 'Completed' }).Count -eq 0)
        TotalSeconds = $time; Peak = $peak; UnknownAdmission = ''; Durations = $Durations }
}

function ConvertTo-BatchChartLabel {
    param([string] $Text)
    # Mermaid treats colons/commas as metadata delimiters.
    return ($Text -replace '[:;,\r\n#<>]', ' ' -replace '\s+', ' ').Trim()
}

function Write-BatchGantt {
    param($Plan, [object[]] $Rows, [string] $Path, [string] $Title)
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add('gantt')
    $lines.Add('    title ' + (ConvertTo-BatchChartLabel $Title))
    $lines.Add('    dateFormat YYYY-MM-DD HH:mm:ss.SSS')
    $lines.Add('    axisFormat %H:%M:%S')
    $lines.Add('    todayMarker off')
    $epoch = [datetime]::new(2000,1,1,0,0,0)
    $index = 0; $section = ''
    foreach ($job in $Plan.Jobs) {
        if ($section -ne $job.BatchKey) {
            $section = $job.BatchKey
            $lines.Add('    section ' + (ConvertTo-BatchChartLabel $section))
        }
        $matches = @($Rows | Where-Object { $_.Id -eq $job.Id })
        if (-not $matches.Count) { $lines.Add("    %% Untimed $($job.Id); predecessors: $($job.Dependencies -join ', ')"); continue }
        foreach ($row in $matches) {
            $index++
            $label = ConvertTo-BatchChartLabel "$($job.Id) attempt $($row.Attempt) $($row.Status)"
            $flag = if ($row.Status -eq 'Completed') { 'done, ' } elseif ($row.Status -in @('Failed','Stopped','Interrupted')) { 'crit, ' } else { '' }
            $start = $epoch.AddSeconds([double]$row.Start).ToString('yyyy-MM-dd HH:mm:ss.fff')
            $end = $epoch.AddSeconds([double]$row.Finish).ToString('yyyy-MM-dd HH:mm:ss.fff')
            $lines.Add(('    %% Predecessors: ' + ($job.Dependencies -join ', ')).TrimEnd())
            $lines.Add("    $label :${flag}task$index, $start, $end")
        }
    }
    # At least one real milestone keeps an empty/unknown forecast renderable.
    if (-not $index) { $lines.Add('    Report origin :milestone, origin, 2000-01-01 00:00:00.000, 0s') }
    [IO.File]::WriteAllLines($Path, $lines, [Text.UTF8Encoding]::new($false))
}

function Write-BatchForecastReport {
    param($Plan, [string] $LogRoot, [string] $Directory, [int] $Limit = 7)
    $forecast = Get-BatchForecast $Plan (Get-BatchDurationHistory $LogRoot $Plan) $Limit
    $forecast.Untimed = @($Plan.Jobs | Where-Object { $_.Id -notin @($forecast.Rows | ForEach-Object { $_.Id }) } | ForEach-Object {
        @{ Id = $_.Id; Predecessors = $_.Dependencies; Reason = 'Unknown duration or start affected by unknown slot occupancy' }
    })
    Write-BatchJson (Join-Path $Directory 'forecast.json') $forecast
    $total = if ($forecast.Complete) { ([timespan]::FromSeconds($forecast.TotalSeconds)).ToString('hh\:mm\:ss') } else { 'unknown - incomplete timing history' }
    Write-BatchGantt $Plan $forecast.Rows (Join-Path $Directory 'planned.mmd') "$($Plan.SequenceProfile) forecast - elapsed $total"
    $notes = @("# Forecast: $($Plan.SequenceProfile)", '', "Estimated elapsed: $total. Peak scheduled workers: $($forecast.Peak).",
        '', 'Historical coordinator timings include process overhead. Input sizes and concurrency may change actual durations.',
        '', 'Original 17m 23s chart retained separately as an unverified illustrative baseline.', '', '## Untimed work', '')
    foreach ($item in $forecast.Untimed) { $notes += "- $($item.Id) — predecessors: $($item.Predecessors -join ', '); $($item.Reason)." }
    if ($notes[-1] -eq '') { $notes = @($notes | Select-Object -SkipLast 1) }
    [IO.File]::WriteAllLines((Join-Path $Directory 'planned-notes.md'), $notes)
    return $forecast
}

function Write-BatchActualReport {
    param($Plan, [hashtable] $State, [string] $Directory)
    $origin = [datetimeoffset]$State.Started
    $rows = [Collections.Generic.List[object]]::new()
    $incomplete = [Collections.Generic.List[object]]::new()
    $attemptFiles = @(Get-ChildItem -LiteralPath $Directory -Recurse -Filter 'attempt.json' -File)
    foreach ($job in $Plan.Jobs) {
        $records = @()
        foreach ($file in $attemptFiles) {
            $item = Read-BatchJson $file.FullName -AsHashtable
            if ($item.Id -eq $job.Id) { $records += $item }
        }
        # Older records may lack per-attempt history. Do not fabricate prior attempts.
        if (-not $records.Count) { $records = @($State.Jobs[$job.Id]) }
        foreach ($record in $records) {
            if (-not $record.Started) { continue }
            $start = ([datetimeoffset]$record.Started - $origin).TotalSeconds
            $finish = if ($record.Finished) { ([datetimeoffset]$record.Finished - $origin).TotalSeconds } else { $null }
            if ($null -eq $finish -or $finish -lt $start) {
                $incomplete.Add(@{ Id=$job.Id; Attempt=$record.Attempt; Status=$record.Status; Reason='No reliable finish timestamp' })
                continue
            }
            $rows.Add(@{ Id = $job.Id; Attempt = $record.Attempt; Status = $record.Status; Start = $start; Finish = $finish })
        }
    }
    $events = @($rows | ForEach-Object { @{ Time = $_.Start; Delta = 1 }; @{ Time = $_.Finish; Delta = -1 } } | Sort-Object Time,Delta)
    $active = 0; $peak = 0; $busy = 0.0; $previous = 0.0
    foreach ($event in $events) {
        if ($active -gt 0) { $busy += $event.Time - $previous }
        $active += $event.Delta; $peak = [Math]::Max($peak, $active); $previous = $event.Time
    }
    $elapsed = ([datetimeoffset]$State.Updated - $origin).TotalSeconds
    $untimed = @($Plan.Jobs | Where-Object { $_.Id -notin @($rows | ForEach-Object { $_.Id }) } | ForEach-Object {
        @{ Id = $_.Id; Status = $State.Jobs[$_.Id].Status; Reason = $State.Jobs[$_.Id].Message; Predecessors = $_.Dependencies }
    })
    $summary = @{ RunId = $State.RunId; Status = $State.Status; Rows = @($rows); Untimed = $untimed; IncompleteAttempts = @($incomplete)
        ElapsedSeconds = $elapsed; ActiveIntervalSeconds = $busy; GapSeconds = [Math]::Max(0,$elapsed-$busy)
        PeakConcurrency = $peak; OriginalIllustrativeBaselineSeconds = 1043
        Comparison = 'Historical illustrative baseline only; not evidence of performance improvement.' }
    Write-BatchJson (Join-Path $Directory 'actual.json') $summary
    $label = ([timespan]::FromSeconds($elapsed)).ToString('hh\:mm\:ss')
    Write-BatchGantt $Plan $rows (Join-Path $Directory 'actual.mmd') "$($Plan.SequenceProfile) $($State.RunId) $($State.Status) - elapsed $label - peak $peak"
    $notes = @("# Actual run: $($State.RunId)", '', "Status: $($State.Status). Elapsed: $label. Peak recorded concurrency: $peak.",
        "Union of observed active intervals: $([math]::Round($busy,3)) seconds. Gaps including resume pauses: $([math]::Round($summary.GapSeconds,3)) seconds.",
        '', 'Technical refresh only; business reconciliation not evaluated.',
        'Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.',
        '', '## Untimed or incomplete work', '')
    foreach ($item in @($untimed) + @($incomplete)) { $notes += "- $($item.Id): $($item.Status); $($item.Reason)" }
    if ($notes[-1] -eq '') { $notes = @($notes | Select-Object -SkipLast 1) }
    [IO.File]::WriteAllLines((Join-Path $Directory 'actual-notes.md'), $notes)
    return $summary
}

function Publish-BatchReports {
    param([string] $Directory, [string] $ReportRoot, [string] $Profile, [string] $RunId)
    [void][IO.Directory]::CreateDirectory($ReportRoot)
    foreach ($name in @('planned.mmd','forecast.json','planned-notes.md','actual.mmd','actual.json','actual-notes.md')) {
        $source = Join-Path $Directory $name
        if (Test-Path -LiteralPath $source) {
            $target = Join-Path $ReportRoot "$Profile-$RunId-$name"
            [IO.File]::Copy($source, $target, $true)
            Write-Host "Report: $target"
        }
    }
}
