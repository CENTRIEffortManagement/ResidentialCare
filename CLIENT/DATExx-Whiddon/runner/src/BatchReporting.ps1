Set-StrictMode -Version Latest

function Get-BatchOutputStatus {
    param($Plan, [hashtable] $State, [string] $DateRoot, [string[]] $AllowedRoots)
    # This assessment reads filesystem metadata only. It never adopts a manual
    # save as a runner receipt or changes the historical run ledger.
    $property = { param($value, $name, $fallback = $null)
        if ($null -eq $value) { return $fallback }
        if ($value -is [Collections.IDictionary]) {
            if ($value.Contains($name)) { return $value[$name] }
        } elseif ($value.PSObject.Properties[$name]) { return $value.$name }
        return $fallback
    }
    if (-not $DateRoot) { $DateRoot = & $property $Plan 'DateRoot' }
    $rootError = ''
    try { if ($DateRoot) { $DateRoot = [IO.Path]::GetFullPath($DateRoot) } else { $rootError = 'The run location is unavailable.' } }
    catch { $rootError = 'The run location is invalid.' }
    $roots = @()
    foreach ($root in @($(if ($AllowedRoots) { $AllowedRoots } else { $DateRoot }))) {
        if (-not $root) { continue }
        try { $roots += [IO.Path]::GetFullPath($root) } catch { $rootError = 'An approved filesystem root is invalid.' }
    }
    $publicRoot = [Environment]::GetEnvironmentVariable('PUBLIC')
    $resolve = { param($base, $path, [string[]] $approved, [bool] $relativeOnly = $false)
        if ([string]::IsNullOrWhiteSpace($path)) { throw 'A declared path is unavailable.' }
        if ($relativeOnly -and [IO.Path]::IsPathRooted($path)) { throw 'A portable path must be relative.' }
        $full = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($path)) { $path } else { Join-Path $base $path }))
        foreach ($root in $approved) {
            $boundary = $root.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
            if ($full.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase)) {
                # Match the runner's existing canonical-path boundary: a junction
                # inside an approved root must not alias an outside workbook.
                $part=$full; $base=$root.TrimEnd('\','/')
                while ($part -and $part.Length -ge $base.Length) {
                    if ((Test-Path -LiteralPath $part -ErrorAction Stop) -and
                        ((Get-Item -LiteralPath $part -Force -ErrorAction Stop).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                        throw 'Reparse-point paths require an explicit canonical configuration.'
                    }
                    $part=Split-Path $part -Parent
                }
                return $full
            }
        }
        throw 'A declared path is outside the approved filesystem roots.'
    }
    $metadata = { param($path)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop)) { return $null }
        $item = Get-Item -LiteralPath $path -ErrorAction Stop
        return @{ Stamp = "$($item.Length):$($item.LastWriteTimeUtc.Ticks)"; Saved = $item.LastWriteTimeUtc }
    }
    $identity = & $property $Plan 'LocationIdentity'
    if (-not $rootError -and $identity -and $DateRoot -and $identity -ne (Get-BatchLocationIdentity $DateRoot)) { $rootError = 'The saved run belongs to a different machine or location.' }
    $planFingerprint=& $property $Plan 'Fingerprint'
    $stateFingerprint=& $property $State 'Fingerprint'
    if (-not $rootError -and (-not $planFingerprint -or -not $stateFingerprint -or $planFingerprint -ne $stateFingerprint)) {
        $rootError='The saved plan and run state do not have matching configuration fingerprints.'
    }
    $jobs = @(& $property $Plan 'Jobs' @())
    $stateJobs=& $property $State 'Jobs' @{}
    if ($stateJobs -isnot [Collections.IDictionary]) { $stateJobs=@{} }
    $byId = @{}; $rows = [Collections.Generic.List[object]]::new(); $rowById = @{}; $pathsById=@{}; $readsById=@{}
    foreach ($job in $jobs) { $id = & $property $job 'Id'; if ($id) { $byId[$id] = $job } }
    $updated = $null
    try { if ($State.Updated) { $updated = ([datetimeoffset]$State.Updated).UtcDateTime } } catch { }
    foreach ($job in $jobs) {
        $id = [string](& $property $job 'Id' '(unknown job)')
        $record = if ($stateJobs.Contains($id)) { $stateJobs[$id] } else { $null }
        $runStatus = [string](& $property $record 'Status' 'Unknown')
        $relative = [string](& $property $job 'RelativePath' '')
        $row = [pscustomobject]@{ Id=$id; RelativePath=$relative; RunStatus=$runStatus; OutputStatus='Unverifiable'
            LastSavedUtc=$null; Reason='Run evidence is incomplete.'; RootCauses=@(); ChangedInputs=@(); SavedAfterDeclaredInputs=$null }
        $rows.Add($row); $rowById[$id] = $row
        try {
            if ($rootError) { throw $rootError }
            $runtimePath = & $property $job 'Path'
            $path = if ($runtimePath) { & $resolve $DateRoot $runtimePath $roots } else { & $resolve $DateRoot $relative $roots $true }
            $pathsById[$id]=$path
            if (-not $row.RelativePath) { $row.RelativePath = [IO.Path]::GetRelativePath($DateRoot, $path).Replace('\','/') }
            $target = & $metadata $path
            if ($null -eq $target) { $row.OutputStatus='Missing'; $row.Reason='The output file is missing.'; continue }
            $row.LastSavedUtc = $target.Saved.ToString('o')
            $declared = [Collections.Generic.List[string]]::new()
            if ($job -is [Collections.IDictionary]) { $hasReads=$job.Contains('Reads'); $hasReadFiles=$job.Contains('ReadFiles') }
            else { $hasReads=$null -ne $job.PSObject.Properties['Reads']; $hasReadFiles=$null -ne $job.PSObject.Properties['ReadFiles'] }
            $inputError = ''
            if ($hasReads) {
                foreach ($read in @(& $property $job 'Reads' @())) {
                    try { $declared.Add((& $resolve $DateRoot $read $roots)) } catch { $inputError=$_.Exception.Message }
                }
            } elseif ($hasReadFiles) {
                foreach ($read in @(& $property $job 'ReadFiles' @())) {
                    try {
                        switch ([string](& $property $read 'Root')) {
                            'Date' { $declared.Add((& $resolve $DateRoot (& $property $read 'Path') $roots $true)); break }
                            'PUBLIC' {
                                if (-not $publicRoot) { throw 'The PUBLIC input root is unavailable.' }
                                $declared.Add((& $resolve $publicRoot (& $property $read 'Path') @([IO.Path]::GetFullPath($publicRoot)) $true)); break
                            }
                            default { throw 'An external input has an opaque or unknown identity.' }
                        }
                    } catch { $inputError=$_.Exception.Message }
                }
            } else { $inputError='Declared input paths are unavailable in this run record.' }
            $readsById[$id]=@($declared | Select-Object -Unique)
            $captured = & $property $record 'Inputs'
            $coverage = $captured -is [Collections.IDictionary]
            $changed = @(); $latestInput = $null; $allInputsObserved=-not [bool]$inputError
            foreach ($read in @($declared | Select-Object -Unique)) {
                $input = & $metadata $read
                if ($null -eq $input) { $allInputsObserved=$false; $changed += [IO.Path]::GetRelativePath($DateRoot,$read).Replace('\','/'); continue }
                if ($null -eq $latestInput -or $input.Saved -gt $latestInput) { $latestInput=$input.Saved }
                $key = @(if ($coverage) { $captured.Keys | Where-Object { [string]::Equals([string]$_,$read,[StringComparison]::OrdinalIgnoreCase) } })
                if ($key.Count -ne 1) { $coverage=$false }
                elseif ([string]$captured[$key[0]] -ne $input.Stamp) { $changed += [IO.Path]::GetRelativePath($DateRoot,$read).Replace('\','/') }
            }
            $row.ChangedInputs = @($changed | Select-Object -Unique)
            if ($latestInput -and $allInputsObserved) { $row.SavedAfterDeclaredInputs = $target.Saved -ge $latestInput }
            if (& $property $record 'NeedsInspection' $false) { $row.OutputStatus='NeedsInspection'; $row.Reason='The run recorded an uncertain save requiring inspection.' }
            elseif ($runStatus -ne 'Completed') {
                $row.OutputStatus='NotUpdated'; $row.Reason="Not updated by this run ($runStatus). " + [string](& $property $record 'Message' '')
                if ($updated -and $target.Saved -gt $updated) { $row.OutputStatus='ChangedAfterRun'; $row.Reason="File modified after this run; its refresh is unverified. Historical run status remains $runStatus." }
                elseif ($row.SavedAfterDeclaredInputs -eq $false) { $row.Reason += ' The output predates a current declared input.' }
            } elseif ($null -eq (& $property $record 'ExitCode') -or (& $property $record 'ExitCode') -ne 0 -or -not (& $property $record 'Finished')) {
                $row.Reason='The completed job lacks a successful exit and finished receipt.'
            } elseif (-not (& $property $record 'OutputStamp')) { $row.Reason='The completed job has no recorded output stamp.' }
            elseif ([string](& $property $record 'OutputStamp') -ne $target.Stamp) { $row.OutputStatus='ChangedAfterRun'; $row.Reason='The output differs from the confirmed run receipt; later changes are unverified.' }
            elseif ($inputError) { $row.Reason=$inputError }
            elseif (-not $coverage) { $row.Reason='Captured input stamps do not cover every declared input.' }
            elseif ($row.ChangedInputs.Count) { $row.OutputStatus='InputsChanged'; $row.Reason='Declared input files changed or are missing after the confirmed refresh.' }
            else { $row.OutputStatus='Verified'; $row.Reason='Confirmed run output and all recorded declared input metadata still match.' }
        } catch { $row.OutputStatus='Unverifiable'; $row.Reason=$_.Exception.Message }
    }
    # Propagate changed producer evidence through actual reads. Ordering-only
    # dependencies still gate execution, but a later save of one is not a data read.
    do {
        $propagated = $false
        foreach ($row in $rows) {
            if ($row.OutputStatus -ne 'Verified') { continue }
            foreach ($dep in @(& $property $byId[$row.Id] 'Dependencies' @())) {
                if ($rowById.ContainsKey($dep) -and ($rowById[$dep].RunStatus -ne 'Completed' -or $rowById[$dep].OutputStatus -eq 'NeedsInspection')) {
                    $row.OutputStatus='InputsChanged'; $row.Reason="Selected prerequisite $dep is not verified: $($rowById[$dep].Reason)"
                    $row.ChangedInputs=@($row.ChangedInputs) + $rowById[$dep].RelativePath; $propagated=$true; break
                }
            }
            if ($row.OutputStatus -ne 'Verified') { continue }
            foreach ($producer in $rows) {
                if ($producer.Id -eq $row.Id -or -not $pathsById.ContainsKey($producer.Id)) { continue }
                if ($pathsById[$producer.Id] -in @($readsById[$row.Id]) -and $producer.OutputStatus -ne 'Verified') {
                    $row.OutputStatus='InputsChanged'; $row.Reason="Declared producer $($producer.Id) is not verified: $($producer.Reason)"
                    $row.ChangedInputs=@($row.ChangedInputs) + $producer.RelativePath; $propagated=$true; break
                }
            }
        }
    } while ($propagated)
    $findRoots = { param($id, [hashtable]$seen)
        if ($seen.ContainsKey($id) -or -not $byId.ContainsKey($id) -or -not $stateJobs.Contains($id)) { return }
        $seen[$id]=$true; $record=$stateJobs[$id]; $status=[string](& $property $record 'Status')
        if ($status -in @('Failed','Stopped') -or (& $property $record 'NeedsInspection' $false)) { "$id`: $([string](& $property $record 'Message' $status))"; return }
        $found=@()
        foreach ($dep in @(& $property $byId[$id] 'Dependencies' @())) { $found += @(& $findRoots $dep $seen) }
        if ($found.Count) { $found | Select-Object -Unique }
        elseif ($status -eq 'Blocked') { "$id`: $([string](& $property $record 'Message' 'Blocked without a recorded failed prerequisite.'))" }
    }
    foreach ($row in $rows) { $row.RootCauses=@(& $findRoots $row.Id @{} | Select-Object -Unique) }
    return [pscustomobject]@{ RunId=[string](& $property $State 'RunId' 'Unknown'); RunStatus=[string](& $property $State 'Status' 'Unknown')
        CheckedAtUtc=[datetime]::UtcNow.ToString('o'); AllVerified=((& $property $State 'Status') -eq 'Completed' -and $rows.Count -gt 0 -and @($rows | Where-Object OutputStatus -ne 'Verified').Count -eq 0); Rows=@($rows) }
}

function Show-BatchOutputStatus {
    param($Assessment)
    $unverified=@($Assessment.Rows | Where-Object OutputStatus -ne 'Verified')
    if ($Assessment.AllVerified) { Write-Host "All $($Assessment.Rows.Count) outputs match their technical run receipts and declared input metadata." }
    elseif ($unverified.Count) {
        Write-Host "OUTPUTS NOT VERIFIED/UPDATED: $($unverified.Count) of $($Assessment.Rows.Count)" -ForegroundColor Yellow
        foreach ($row in $unverified) {
            Write-Host "$($row.Id) | $($row.OutputStatus) | $($row.RelativePath)"
            Write-Host "  $($row.Reason) Last saved UTC: $($row.LastSavedUtc)"
            foreach ($cause in $row.RootCauses) { Write-Host "  Root cause: $cause" }
        }
    } else { Write-Host "Run status remains $($Assessment.RunStatus); matching output metadata does not complete the historical run." -ForegroundColor Yellow }
    Write-Host 'Metadata check only; no workbook was opened or refreshed. Business results were not validated.'
}

function Write-BatchOutputStatusReport {
    param($Plan, [hashtable] $State, [string] $Directory, [string] $DateRoot, [string[]] $AllowedRoots)
    $assessment=Get-BatchOutputStatus -Plan $Plan -State $State -DateRoot $DateRoot -AllowedRoots $AllowedRoots
    [void][IO.Directory]::CreateDirectory($Directory)
    Write-BatchJson (Join-Path $Directory 'output-status.json') $assessment
    $bad=@($assessment.Rows | Where-Object OutputStatus -ne 'Verified')
    $headline=if ($assessment.AllVerified) { "**All $($assessment.Rows.Count) outputs match their technical receipts and declared input metadata.**" }
        elseif ($bad.Count) { "**OUTPUTS NOT VERIFIED/UPDATED: $($bad.Count) of $($assessment.Rows.Count).**" }
        else { "**Historical run status remains $($assessment.RunStatus); matching output metadata does not complete the run.**" }
    $lines=@("# Output status: $($assessment.RunId)", '', "Historical run status: $($assessment.RunStatus). Checked UTC: $($assessment.CheckedAtUtc).",
        '', $headline, '',
        'Filesystem metadata only. This checker did not open, refresh, save or change a workbook. A later manual save is not a runner completion receipt. Business results were not validated.', '')
    foreach ($row in $assessment.Rows) {
        $lines += "- $($row.Id): **$($row.OutputStatus)**; $($row.RelativePath); last saved UTC: $($row.LastSavedUtc). $($row.Reason)"
        foreach ($cause in $row.RootCauses) { $lines += "  Root cause: $cause" }
    }
    [IO.File]::WriteAllLines((Join-Path $Directory 'output-status.md'),$lines,[Text.UTF8Encoding]::new($false))
    return $assessment
}

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
    $savedInputNotice = @(Get-BatchSavedIntermediateNotice $Plan)
    if ($savedInputNotice.Count) { $notes += @('', '## Saved intermediate outputs', '') + $savedInputNotice }
    if ($notes[-1] -eq '') { $notes = @($notes | Select-Object -SkipLast 1) }
    [IO.File]::WriteAllLines((Join-Path $Directory 'planned-notes.md'), $notes)
    return $forecast
}

function Write-BatchActualReport {
    param($Plan, [hashtable] $State, [string] $Directory, [string] $DateRoot, [string[]] $AllowedRoots)
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
            $rows.Add(@{ Id = $job.Id; Attempt = $record.Attempt; Status = $record.Status; Start = $start; Finish = $finish
                Message = $(if ($record.ContainsKey('Message')) { $record.Message } else { '' }) })
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
    if ($Plan.PSObject.Properties['SavedIntermediates']) { $summary.SavedIntermediates = @($Plan.SavedIntermediates) }
    $assessment=$null
    if ($DateRoot -or $Plan.PSObject.Properties['DateRoot']) {
        $assessment=Write-BatchOutputStatusReport -Plan $Plan -State $State -Directory $Directory -DateRoot $DateRoot -AllowedRoots $AllowedRoots
        $summary.OutputStatus=$assessment
    }
    Write-BatchJson (Join-Path $Directory 'actual.json') $summary
    $label = ([timespan]::FromSeconds($elapsed)).ToString('hh\:mm\:ss')
    Write-BatchGantt $Plan $rows (Join-Path $Directory 'actual.mmd') "$($Plan.SequenceProfile) $($State.RunId) $($State.Status) - elapsed $label - peak $peak"
    $notes = @("# Actual run: $($State.RunId)", '', "Status: $($State.Status). Elapsed: $label. Peak recorded concurrency: $peak.",
        "Union of observed active intervals: $([math]::Round($busy,3)) seconds. Gaps including resume pauses: $([math]::Round($summary.GapSeconds,3)) seconds.",
        '', 'Technical refresh only; business reconciliation not evaluated.',
        'Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.')
    if ($null -ne $assessment) {
        $bad=@($assessment.Rows | Where-Object OutputStatus -ne 'Verified')
        $headline=if ($assessment.AllVerified) { "**All $($assessment.Rows.Count) outputs match their technical receipts and declared input metadata.**" }
            elseif ($bad.Count) { "**OUTPUTS NOT VERIFIED/UPDATED: $($bad.Count) of $($assessment.Rows.Count).**" }
            else { "**Historical run status remains $($assessment.RunStatus); matching output metadata does not complete the run.**" }
        $notes += @('', $headline, '', 'Metadata assessment only; later manual saves are not runner receipts. See output-status.md for every output.')
        foreach ($item in $bad) {
            $notes += "- $($item.Id): $($item.OutputStatus); $($item.Reason)"
            foreach ($cause in $item.RootCauses) { $notes += "  Root cause: $cause" }
        }
    }
    $failed=@($rows | Where-Object Status -in @('Failed','Stopped','Interrupted'))
    if ($failed.Count) { $notes += @('', '## Failed or stopped attempts', ''); foreach ($item in $failed) { $notes += "- $($item.Id): $($item.Status); $($item.Message)" } }
    $notes += @('', '## Untimed or incomplete work', '')
    foreach ($item in @($untimed) + @($incomplete)) { $notes += "- $($item.Id): $($item.Status); $($item.Reason)" }
    $savedInputNotice = @(Get-BatchSavedIntermediateNotice $Plan)
    if ($savedInputNotice.Count) { $notes += @('', '## Saved intermediate outputs', '') + $savedInputNotice }
    if ($notes[-1] -eq '') { $notes = @($notes | Select-Object -SkipLast 1) }
    [IO.File]::WriteAllLines((Join-Path $Directory 'actual-notes.md'), $notes)
    return $summary
}

function Publish-BatchReports {
    param([string] $Directory, [string] $ReportRoot, [string] $Profile, [string] $RunId)
    [void][IO.Directory]::CreateDirectory($ReportRoot)
    foreach ($name in @('planned.mmd','forecast.json','planned-notes.md','actual.mmd','actual.json','actual-notes.md','output-status.json','output-status.md')) {
        $source = Join-Path $Directory $name
        if (Test-Path -LiteralPath $source) {
            $target = Join-Path $ReportRoot "$Profile-$RunId-$name"
            [IO.File]::Copy($source, $target, $true)
            Write-Host "Report: $target"
        }
    }
}
