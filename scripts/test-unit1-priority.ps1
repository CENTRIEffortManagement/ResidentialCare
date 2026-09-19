param([string] $MermaidBundle, [string] $Chromium)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$src = Join-Path $repo 'CLIENT/DATExx-Whiddon/runner/src'
. (Join-Path $src 'BatchPlanning.ps1')
. (Join-Path $src 'BatchScheduling.ps1')
. (Join-Path $src 'BatchReporting.ps1')
$checks = 0
function Check([bool] $condition, [string] $message) {
    if (-not $condition) { throw "FAIL: $message" }
    $script:checks++
}
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('priority test ' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temporary)
try {
    $old = Get-BatchCatalogue $repo
    $new = Get-BatchCatalogue $repo 'ParallelInputs-Unit1Priority'
    Check ($old.Fingerprint -ne $new.Fingerprint) 'separate catalogue fingerprints'
    Check ((Get-FileHash (Join-Path $repo 'docs/batch-duration-gantt.mmd')).Hash -eq 'F29199C84B672FD8394203D73CB650E9358A3D09186AA1706256F735FE659DCD') 'original 17m23s baseline unchanged'
    Check ('Unit1/AllocationInput' -in @($old.Jobs | Where-Object Id -eq 'Unit1/Settings')[0].Dependencies) 'Current retains its edges'
    $roots = @($new.Jobs | Where-Object { -not $_.Dependencies.Count } | ForEach-Object Id)
    Check (($roots -join ',') -eq 'Unit1/AllocationInput,Unit1/Settings,Unit1/DemandMaster,Unit2/AllocationInput,Unit2/Settings,Unit2/DemandMaster') 'six roots in Unit priority order'
    $plan = Select-BatchPlan $new -RunAll
    $times = @{}; foreach ($job in $plan.Jobs) { $times[$job.Id] = @{ Seconds = 10; SourceRun = 'synthetic'; Observed = '2026-01-01T00:00:00Z' } }
    foreach ($limit in 1..7) {
        $forecast = Get-BatchForecast $plan $times $limit
        Check ($forecast.Complete -and $forecast.Rows.Count -eq $plan.Jobs.Count -and $forecast.Peak -le $limit) "complete bounded forecast $limit"
        $byId = @{}; foreach ($row in $forecast.Rows) { $byId[$row.Id] = $row }
        foreach ($job in $plan.Jobs) {
            foreach ($dep in $job.Dependencies) { Check ($byId[$job.Id].Start -ge $byId[$dep].Finish) "edge $dep -> $($job.Id)" }
        }
    }
    $forecast = Get-BatchForecast $plan $times 7
    Check (($forecast.Rows | Select-Object -First 6 | ForEach-Object Id) -join ',' -eq ($roots -join ',')) 'startup dispatch order'
    $status=@{}; foreach($job in $plan.Jobs){$status[$job.Id]='Pending'}
    $first=$plan.Jobs[0]; $other=$plan.Jobs | Where-Object Id -eq 'Unit2/AllocationInput'
    $other.Exclusive=$true
    Check (-not (Test-BatchAdmission $first $status @($other) 7)) 'exclusive worker blocks other admissions'
    $other.Exclusive=$false
    $savedReads=$first.Reads; $first.Reads=@($other.Path)
    Check (-not (Test-BatchAdmission $first $status @($other) 7)) 'active writer excludes readers'
    $first.Reads=$savedReads
    $times.Remove('Org/Cost')
    $unknown = Get-BatchForecast $plan $times 7
    Check (-not $unknown.Complete -and $null -eq $unknown.TotalSeconds) 'unknown timings cannot produce a total'
    Check ('Org/Reporting' -notin @($unknown.Rows | ForEach-Object Id)) 'unknown successor is not assigned a bar'
    $portable = ConvertTo-PortableBatchPlan $plan
    $json = $portable | ConvertTo-Json -Depth 40
    Check (-not $json.Contains($repo.Replace('\','\\'))) 'portable plan has no repository absolute paths'
    Check ($portable.LocationIdentity -ne (Get-BatchLocationIdentity (Join-Path $temporary 'relocated'))) 'relocation blocks resume'

    # Relocate only scripts/configuration, never copy or inspect real workbooks.
    $fixture = Join-Path $temporary 'relocated project'
    $fixtureDate = Join-Path $fixture 'CLIENT/DATExx-Whiddon'
    [void][IO.Directory]::CreateDirectory($fixtureDate)
    Copy-Item -LiteralPath (Join-Path $repo 'pq.project.json') -Destination $fixtureDate
    Move-Item -LiteralPath (Join-Path $fixtureDate 'pq.project.json') -Destination $fixture
    Copy-Item -LiteralPath (Join-Path $repo 'CLIENT/DATExx-Whiddon/runner') -Destination $fixtureDate -Recurse
    foreach ($unit in @('Unit1','Unit2')) { [void][IO.Directory]::CreateDirectory((Join-Path $fixtureDate "UNITS/$unit")) }
    Push-Location ([IO.Path]::GetTempPath())
    try {
        $relocated = Get-BatchCatalogue $fixture 'ParallelInputs-Unit1Priority'
        Check ($relocated.Jobs[0].Path.StartsWith($fixture)) 'roots resolve at relocated project'
        $output = & pwsh -NoProfile -File (Join-Path $fixtureDate 'runner/Invoke-BatchRefreshRunner.ps1') -SequenceProfile ParallelInputs-Unit1Priority -ShowPlan -Units Unit1 2>&1
        Check ($LASTEXITCODE -eq 0 -and ($output -join ' ') -match '26 files') 'launcher independent of working directory'
        $reportReject = & pwsh -NoProfile -File (Join-Path $fixtureDate 'runner/Invoke-BatchRefreshRunner.ps1') -ReportRun fixture -RefreshSelected 2>&1
        Check ($LASTEXITCODE -ne 0 -and ($reportReject -join ' ') -match 'report-only') 'report regeneration cannot execute'
        $resumeDirectory=Join-Path $fixtureDate 'RunLogs/BatchRefresh/fixture'
        [void][IO.Directory]::CreateDirectory($resumeDirectory)
        Write-BatchJson (Join-Path $resumeDirectory 'plan.json') (ConvertTo-PortableBatchPlan $plan)
        $mismatch = & pwsh -NoProfile -File (Join-Path $fixtureDate 'runner/Invoke-BatchRefreshRunner.ps1') -ResumeRun fixture -SequenceProfile Current -RefreshSelected 2>&1
        Check ($LASTEXITCODE -ne 0 -and ($mismatch -join ' ') -match 'conflicts with the saved') 'profile mismatch rejected before execution'
        $localPlan = Select-BatchPlan $relocated -Units Unit1
        Write-BatchJson (Join-Path $resumeDirectory 'plan.json') (ConvertTo-PortableBatchPlan $localPlan)
        Write-BatchJson (Join-Path $resumeDirectory 'state.json') (New-BatchState $localPlan 'fixture')
        $resumePreview = & pwsh -NoProfile -File (Join-Path $fixtureDate 'runner/Invoke-BatchRefreshRunner.ps1') -ResumeRun fixture -ShowPlan 2>&1
        Check ($LASTEXITCODE -eq 0 -and ($resumePreview -join ' ') -match 'Sequence profile: ParallelInputs-Unit1Priority') 'resume reloads saved profile across processes'
        Write-BatchJson (Join-Path $resumeDirectory 'plan.json') (ConvertTo-PortableBatchPlan $plan)
        $relocation = & pwsh -NoProfile -File (Join-Path $fixtureDate 'runner/Invoke-BatchRefreshRunner.ps1') -ResumeRun fixture -ShowPlan 2>&1
        Check ($LASTEXITCODE -ne 0 -and ($relocation -join ' ') -match 'start a fresh run') 'relocated resume rejected'
    } finally { Pop-Location }

    # Compare the live coordinator (mock workers) with virtual-clock admission.
    $jobs = @()
    foreach ($spec in @(
        @{ Id='Unit1/a'; Depends=@(); Seconds=0.05 },
        @{ Id='Unit1/b'; Depends=@('Unit1/a'); Seconds=0.02 },
        @{ Id='Unit2/a'; Depends=@(); Seconds=0.15 },
        @{ Id='Unit2/b'; Depends=@(); Seconds=0.02 }
    )) {
        $path = Join-Path $temporary ($spec.Id.Replace('/','-') + '.data')
        [IO.File]::WriteAllText($path, 'synthetic')
        $jobs += [pscustomobject]@{ Id=$spec.Id; BatchKey=$spec.Id; Path=$path; Reads=@(); Dependencies=$spec.Depends; Exclusive=$false; Seconds=$spec.Seconds }
    }
    $syntheticPlan = [pscustomobject]@{ SchemaVersion=1; Fingerprint='fixture'; DateRoot=$temporary; SequenceProfile='ParallelInputs-Unit1Priority'; Jobs=$jobs }
    $duration = @{}; foreach ($job in $jobs) { $duration[$job.Id]=@{Seconds=$job.Seconds} }
    $predicted = Get-BatchForecast $syntheticPlan $duration 2
    $directory=Join-Path $temporary 'run'
    [void][IO.Directory]::CreateDirectory($directory)
    $state=New-BatchState $syntheticPlan 'fixture'
    $events=[Collections.Generic.List[string]]::new()
    $start = { param($job,$dir,$options) $events.Add($job.Id); return @{ Job=$job; Timer=[Diagnostics.Stopwatch]::StartNew() } }.GetNewClosure()
    $poll = { param($handle) if($handle.Timer.Elapsed.TotalSeconds -lt $handle.Job.Seconds){return $null}; return @{ExitCode=0;NeedsInspection=$false;Message='Synthetic confirmed save'} }
    $result=Invoke-BatchSchedule $syntheticPlan $state $directory -MaxParallelBatches 2 -StartWorker $start -PollWorker $poll -ValidateJob {} -PollMilliseconds 1
    Check ($result -eq 0) 'mock coordinator succeeds'
    Check (($events -join ',') -eq (($predicted.Rows | ForEach-Object Id) -join ',')) 'virtual and live admission match'
    Check ($events[2] -eq 'Unit1/b') 'newly ready Unit1 wins next slot'
    Check (@(Get-ChildItem $directory -Recurse -Filter attempt.json).Count -eq 4) 'durable attempt per dispatch'
    $actual=Write-BatchActualReport $syntheticPlan $state $directory
    Check ($actual.Rows.Count -eq 4 -and $actual.PeakConcurrency -eq 2) 'actual interval accounting'
    $history=Get-Content (Join-Path $directory 'job-000/attempt-1/attempt.json') -Raw
    $state.Jobs['Unit1/a'].Status='Failed'
    Reset-BatchResume $syntheticPlan $state
    $result=Invoke-BatchSchedule $syntheticPlan $state $directory -MaxParallelBatches 2 -StartWorker $start -PollWorker $poll -ValidateJob {} -PollMilliseconds 1
    Check ((Get-Content (Join-Path $directory 'job-000/attempt-1/attempt.json') -Raw) -eq $history) 'resume preserves old attempt'
    $actual=Write-BatchActualReport $syntheticPlan $state $directory
    Check ($actual.Rows.Count -gt 4) 'actual report includes retry history'
    if ($MermaidBundle -and $Chromium) {
        & node (Join-Path $PSScriptRoot 'test-gantt-render.cjs') $MermaidBundle $Chromium (Join-Path $directory 'actual.mmd')
        Check ($LASTEXITCODE -eq 0) 'actual retry chart renders with Mermaid'
    }
    Write-Host "PASS: $checks priority, forecast, attempt-history and portability checks. No Excel opened."
} finally {
    $resolved=[IO.Path]::GetFullPath($temporary)
    if (-not $resolved.StartsWith([IO.Path]::GetTempPath(),[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe cleanup target' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
