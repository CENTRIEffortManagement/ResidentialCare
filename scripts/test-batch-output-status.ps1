# Metadata-only output readiness regression tests. All data files are disposable synthetic fixtures.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchScheduling.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchReporting.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-output-status-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$beforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$originalPublic = [Environment]::GetEnvironmentVariable('PUBLIC', 'Process')
$script:checks = 0
$script:lastOutputRow = $null
$junctions = [Collections.Generic.List[string]]::new()

function Assert-OutputStatus {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) {
        $actual = if ($null -eq $script:lastOutputRow) { '' } else { '; actual row: ' + ($script:lastOutputRow | ConvertTo-Json -Depth 10 -Compress) }
        throw "FAIL: $Message$actual"
    }
    $script:checks++
}

function Set-SyntheticFile {
    param([string] $Path, [datetime] $SavedUtc, [string] $Text = 'synthetic data')
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path, $Text)
    [IO.File]::SetLastWriteTimeUtc($Path, $SavedUtc)
}

function New-OutputFixture {
    param([string] $Name)
    $directory = Join-Path $testRoot $Name
    $dateRoot = Join-Path $directory 'client'
    [void][IO.Directory]::CreateDirectory($dateRoot)
    $time = [datetime]::UtcNow.AddHours(-4)
    $inputPath = Join-Path $dateRoot 'Capacity.data'
    $outputPath = Join-Path $dateRoot 'Effort.data'
    Set-SyntheticFile $inputPath $time.AddMinutes(-10) 'saved capacity fixture'
    Set-SyntheticFile $outputPath $time.AddMinutes(10) 'saved effort fixture'
    $job = [pscustomobject]@{
        Id = 'BD/Effort'; Unit = 'BD'; Role = ''; Batch = 'U4'; BatchKey = 'BD/U4'
        RelativePath = 'Effort.data'; Path = $outputPath; Reads = @($inputPath); Dependencies = @(); Exclusive = $false
    }
    $plan = [pscustomobject]@{
        SchemaVersion = 1; Fingerprint = 'metadata-fixture'; SequenceProfile = 'Current'; DateRoot = $dateRoot; Jobs = @($job)
    }
    $state = New-BatchState $plan $Name
    $state.Status = 'Completed'; $state.ExitCode = 0
    $state.Started = $time.ToString('o'); $state.Updated = $time.AddMinutes(60).ToString('o')
    $record = $state.Jobs[$job.Id]
    $record.Status = 'Completed'; $record.ExitCode = 0; $record.Attempt = 1
    $record.Started = $time.ToString('o'); $record.Finished = $time.AddMinutes(20).ToString('o')
    $record.OutputStamp = Get-BatchFileStamp $outputPath
    $record.Inputs[$inputPath] = Get-BatchFileStamp $inputPath
    return [pscustomobject]@{
        Directory = $directory; DateRoot = $dateRoot; Time = $time; Input = $inputPath; Output = $outputPath
        Job = $job; Plan = $plan; State = $state; Record = $record
    }
}

function Get-FixtureAssessment {
    param($Fixture)
    return Get-BatchOutputStatus -Plan $Fixture.Plan -State $Fixture.State -DateRoot $Fixture.DateRoot -AllowedRoots @($Fixture.DateRoot)
}

function Get-EffortRow {
    param($Assessment)
    $rows = @($Assessment.Rows | Where-Object Id -eq 'BD/Effort')
    Assert-OutputStatus ($rows.Count -eq 1) 'assessment contains exactly one Effort row'
    $script:lastOutputRow = $rows[0]
    return $rows[0]
}

function Test-ChangedInputReported {
    param($Row, [string] $Path, [string] $DateRoot)
    $reported = @($Row.ChangedInputs | ForEach-Object { $_.Replace('\', '/') })
    $relative = [IO.Path]::GetRelativePath($DateRoot, $Path).Replace('\', '/')
    return $reported -contains $Path.Replace('\', '/') -or $reported -contains $relative
}

function Add-FixtureProducer {
    param($Fixture, [string] $Id, [string] $Status = 'Completed', [switch] $ReadByEffort)
    $relative = ($Id -replace '[^A-Za-z0-9_.-]', '-') + '.data'
    $path = Join-Path $Fixture.DateRoot $relative
    Set-SyntheticFile $path $Fixture.Time.AddMinutes(-5) "producer $Id"
    $job = [pscustomobject]@{
        Id = $Id; Unit = 'BD'; Role = ''; Batch = 'A3'; BatchKey = 'BD/A3'; RelativePath = $relative
        Path = $path; Reads = @(); Dependencies = @(); Exclusive = $false
    }
    $Fixture.Plan.Jobs = @($job) + @($Fixture.Plan.Jobs)
    $Fixture.Job.Dependencies = @($Fixture.Job.Dependencies) + @($Id)
    if ($ReadByEffort) {
        $Fixture.Job.Reads = @($Fixture.Job.Reads) + @($path)
        $Fixture.Record.Inputs[$path] = Get-BatchFileStamp $path
    }
    $Fixture.State.Jobs[$Id] = @{
        Status = $Status; Message = ''; Attempt = 1; Started = $Fixture.Time.AddMinutes(-10).ToString('o')
        Finished = $Fixture.Time.ToString('o'); ExitCode = $(if ($Status -eq 'Completed') { 0 } else { 1 })
        OutputStamp = (Get-BatchFileStamp $path); Inputs = @{}; NeedsInspection = $false; Directory = ''
    }
    return $job
}

try {
    $fixture = New-OutputFixture 'confirmed'
    $stateBefore = $fixture.State | ConvertTo-Json -Depth 30 -Compress
    $assessment = Get-FixtureAssessment $fixture
    $row = Get-EffortRow $assessment
    Assert-OutputStatus ($row.OutputStatus -eq 'Verified' -and $assessment.AllVerified) ("matching completed receipt and complete inputs are verified; actual: " + ($assessment | ConvertTo-Json -Depth 10 -Compress))
    Assert-OutputStatus ($row.LastSavedUtc -and $assessment.CheckedAtUtc) 'assessment reports current output save and assessment times'
    Assert-OutputStatus ($row.SavedAfterDeclaredInputs -eq $true) 'saved output follows its declared input timestamps'
    Assert-OutputStatus (($fixture.State | ConvertTo-Json -Depth 30 -Compress) -ceq $stateBefore) 'assessment leaves the run ledger byte-for-byte unchanged'

    $blocked = New-OutputFixture 'blocked-effort'
    $producer = Add-FixtureProducer $blocked 'BD/MultiRole' 'Failed' -ReadByEffort
    $blocked.State.Jobs[$producer.Id].Message = "Expression.Error: The key didn't match any rows in the table."
    $blocked.State.Status = 'PartialOrFailed'; $blocked.State.ExitCode = 1
    $blocked.Record.Status = 'Blocked'; $blocked.Record.ExitCode = $null
    $blocked.Record.Started = ''; $blocked.Record.Finished = ''; $blocked.Record.OutputStamp = ''
    $blocked.Record.Message = 'Prerequisite BD/MultiRole did not complete.'
    Set-SyntheticFile $blocked.Input $blocked.Time.AddMinutes(5) 'new capacity from current run'
    Set-SyntheticFile $blocked.Output $blocked.Time.AddMinutes(-20) 'old effort before current run'
    $beforeBlocked = $blocked.State | ConvertTo-Json -Depth 30 -Compress
    $blockedAssessment = Get-FixtureAssessment $blocked
    $row = Get-EffortRow $blockedAssessment
    Assert-OutputStatus ($row.OutputStatus -eq 'NotUpdated' -and -not $blockedAssessment.AllVerified) 'blocked Effort is prominently not updated'
    Assert-OutputStatus ($row.RunStatus -eq 'Blocked' -and $row.SavedAfterDeclaredInputs -eq $false) 'old output is distinguished from newly saved Capacity'
    Assert-OutputStatus (($row.RootCauses -join ' ') -match 'BD/MultiRole' -and ($row.RootCauses -join ' ') -match 'key.*match') 'blocked output identifies the failed producer and original key error'
    Assert-OutputStatus (($blocked.State | ConvertTo-Json -Depth 30 -Compress) -ceq $beforeBlocked) 'blocked assessment does not repair the historical ledger'

    Set-SyntheticFile $blocked.Output $blocked.Time.AddMinutes(80) 'later manual Effort save after all inputs'
    $manualAssessment = Get-FixtureAssessment $blocked
    $row = Get-EffortRow $manualAssessment
    Assert-OutputStatus ($row.OutputStatus -eq 'ChangedAfterRun' -and -not $manualAssessment.AllVerified) 'later manual save is changed and unverified rather than silently accepted'
    Assert-OutputStatus ($row.RunStatus -eq 'Blocked' -and $row.SavedAfterDeclaredInputs -eq $true) 'manual save ordering is updated while historical blocked status remains'
    Assert-OutputStatus (($blocked.State | ConvertTo-Json -Depth 30 -Compress) -ceq $beforeBlocked) 'manual-save assessment leaves status, exit code and input receipts unchanged'

    $changedInput = New-OutputFixture 'changed-input'
    Set-SyntheticFile $changedInput.Input $changedInput.Time.AddMinutes(30) 'changed capacity input'
    $row = Get-EffortRow (Get-FixtureAssessment $changedInput)
    Assert-OutputStatus ($row.OutputStatus -eq 'InputsChanged' -and (Test-ChangedInputReported $row $changedInput.Input $changedInput.DateRoot)) 'changed captured input prevents verification and names the file'
    Assert-OutputStatus ($row.SavedAfterDeclaredInputs -eq $false) 'output older than a changed input has explicit ordering failure'

    $missingInput = New-OutputFixture 'missing-input'
    [IO.File]::Delete($missingInput.Input)
    $row = Get-EffortRow (Get-FixtureAssessment $missingInput)
    Assert-OutputStatus ($row.OutputStatus -ne 'Verified' -and (Test-ChangedInputReported $row $missingInput.Input $missingInput.DateRoot)) 'missing input is reported and cannot verify a completed receipt'
    $remainingInput = Join-Path $missingInput.DateRoot 'remaining-input.data'
    Set-SyntheticFile $remainingInput $missingInput.Time.AddMinutes(-10)
    $missingInput.Job.Reads = @($missingInput.Input, $remainingInput)
    $missingInput.Record.Inputs[$remainingInput] = Get-BatchFileStamp $remainingInput
    $row = Get-EffortRow (Get-FixtureAssessment $missingInput)
    Assert-OutputStatus ($null -eq $row.SavedAfterDeclaredInputs) 'save ordering remains unknown when any declared input is missing'

    $missingOutput = New-OutputFixture 'missing-output'
    [IO.File]::Delete($missingOutput.Output)
    $missingOutput.Record.OutputStamp = 'missing'
    $row = Get-EffortRow (Get-FixtureAssessment $missingOutput)
    Assert-OutputStatus ($row.OutputStatus -eq 'Missing') 'matching missing output stamps never count as verified'

    $legacy = New-OutputFixture 'incomplete-legacy-inputs'
    $legacy.Record.Inputs = @{}
    $row = Get-EffortRow (Get-FixtureAssessment $legacy)
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'missing legacy input snapshot cannot verify a consumer'

    $badReceipt = New-OutputFixture 'nonzero-completed-receipt'
    $badReceipt.Record.ExitCode = 1
    $row = Get-EffortRow (Get-FixtureAssessment $badReceipt)
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'a completed label without a successful receipt cannot verify an output'
    $badReceipt.Record.ExitCode = 0
    $badReceipt.Record.Finished = ''
    $row = Get-EffortRow (Get-FixtureAssessment $badReceipt)
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'an incomplete legacy receipt without a finish time is not verified'

    $wrongPlan = New-OutputFixture 'different-saved-plan'
    $wrongPlan.Plan.Fingerprint = 'different-plan-from-state'
    $row = Get-EffortRow (Get-FixtureAssessment $wrongPlan)
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'saved plan and state fingerprint mismatch cannot verify outputs'

    $uncertain = New-OutputFixture 'uncertain-save'
    $uncertain.Record.NeedsInspection = $true
    $row = Get-EffortRow (Get-FixtureAssessment $uncertain)
    Assert-OutputStatus ($row.OutputStatus -eq 'NeedsInspection') 'uncertain save retains its inspection requirement'

    $readProducer = New-OutputFixture 'changed-read-producer'
    $producer = Add-FixtureProducer $readProducer 'BD/Capacity' -ReadByEffort
    Set-SyntheticFile $producer.Path $readProducer.Time.AddMinutes(30) 'producer changed after confirmed run'
    $row = Get-EffortRow (Get-FixtureAssessment $readProducer)
    Assert-OutputStatus ($row.OutputStatus -ne 'Verified' -and (Test-ChangedInputReported $row $producer.Path $readProducer.DateRoot)) 'changed selected producer invalidates a consumer that reads it'

    $transitive = New-OutputFixture 'stale-producer-input'
    $producer = Add-FixtureProducer $transitive 'BD/Capacity' -ReadByEffort
    $upstreamPath = Join-Path $transitive.DateRoot 'upstream.data'
    Set-SyntheticFile $upstreamPath $transitive.Time.AddMinutes(-20) 'captured source input'
    $producer.Reads = @($upstreamPath)
    $transitive.State.Jobs[$producer.Id].Inputs[$upstreamPath] = Get-BatchFileStamp $upstreamPath
    Set-SyntheticFile $upstreamPath $transitive.Time.AddMinutes(30) 'source changed after Capacity saved'
    $row = Get-EffortRow (Get-FixtureAssessment $transitive)
    Assert-OutputStatus ($row.OutputStatus -eq 'InputsChanged' -and (Test-ChangedInputReported $row $producer.Path $transitive.DateRoot)) 'stale upstream evidence propagates through an unchanged consumed producer file'

    $orderingOnly = New-OutputFixture 'changed-order-only-producer'
    $producer = Add-FixtureProducer $orderingOnly 'BD/OrderingOnly'
    Set-SyntheticFile $producer.Path $orderingOnly.Time.AddMinutes(30) 'ordering-only file re-saved'
    $assessment = Get-FixtureAssessment $orderingOnly
    $row = Get-EffortRow $assessment
    Assert-OutputStatus ($row.OutputStatus -eq 'Verified' -and -not (Test-ChangedInputReported $row $producer.Path $orderingOnly.DateRoot)) 'changed completed ordering-only predecessor does not invent a consumed input'
    Assert-OutputStatus (-not $assessment.AllVerified) 'a verified consumer does not conceal the changed producer row'
    $orderingOnly.State.Jobs[$producer.Id].Status = 'Failed'
    $orderingOnly.State.Jobs[$producer.Id].Message = 'Selected prerequisite failed.'
    $row = Get-EffortRow (Get-FixtureAssessment $orderingOnly)
    Assert-OutputStatus ($row.OutputStatus -ne 'Verified') 'failed selected ordering predecessor still prevents run verification'
    $orderingOnly.State.Jobs[$producer.Id].Status = 'Completed'
    $orderingOnly.State.Jobs[$producer.Id].NeedsInspection = $true
    $row = Get-EffortRow (Get-FixtureAssessment $orderingOnly)
    Assert-OutputStatus ($row.OutputStatus -ne 'Verified') 'uncertain save in a selected ordering predecessor retains the dependency inspection guard'

    $portable = New-OutputFixture 'portable-date-plan'
    $portablePlan = ConvertTo-PortableBatchPlan $portable.Plan
    $assessment = Get-BatchOutputStatus -Plan $portablePlan -State $portable.State -DateRoot $portable.DateRoot -AllowedRoots @($portable.DateRoot)
    Assert-OutputStatus ((Get-EffortRow $assessment).OutputStatus -eq 'Verified') 'portable Date reads resolve under the declared current date root'
    $portablePlan.LocationIdentity = 'another-machine-or-checkout'
    $row = Get-EffortRow (Get-BatchOutputStatus -Plan $portablePlan -State $portable.State -DateRoot $portable.DateRoot -AllowedRoots @($portable.DateRoot))
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'portable location mismatch is not silently rebound'

    $escaping = New-OutputFixture 'escaping-portable-read'
    $portablePlan = ConvertTo-PortableBatchPlan $escaping.Plan
    $portablePlan.Jobs[0].ReadFiles[0].Path = '../outside.data'
    $row = Get-EffortRow (Get-BatchOutputStatus -Plan $portablePlan -State $escaping.State -DateRoot $escaping.DateRoot -AllowedRoots @($escaping.DateRoot))
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'portable Date path traversal is rejected'

    $outside = New-OutputFixture 'outside-runtime-target'
    $outsidePath = Join-Path $outside.Directory 'outside.data'
    Set-SyntheticFile $outsidePath $outside.Time.AddMinutes(10)
    $outside.Job.Path = $outsidePath
    $outside.Record.OutputStamp = Get-BatchFileStamp $outsidePath
    $row = Get-EffortRow (Get-FixtureAssessment $outside)
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'runtime targets outside explicitly allowed roots cannot verify'

    if ($IsWindows) {
        $alias = New-OutputFixture 'junction-target'
        $outsideDirectory = Join-Path $alias.Directory 'outside-approved-root'
        $aliasedOutput = Join-Path $outsideDirectory 'Effort.data'
        Set-SyntheticFile $aliasedOutput $alias.Time.AddMinutes(10)
        $junction = Join-Path $alias.DateRoot 'alias'
        $null = New-Item -ItemType Junction -Path $junction -Target $outsideDirectory
        $junctions.Add($junction)
        $alias.Job.Path = Join-Path $junction 'Effort.data'
        $alias.Record.OutputStamp = Get-BatchFileStamp $alias.Job.Path
        $row = Get-EffortRow (Get-FixtureAssessment $alias)
        Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'junction aliases cannot bypass the existing approved-root guard'
    }

    $public = New-OutputFixture 'portable-public-read'
    $publicRoot = Join-Path $public.Directory 'public'
    [void][IO.Directory]::CreateDirectory($publicRoot)
    [Environment]::SetEnvironmentVariable('PUBLIC', $publicRoot, 'Process')
    $publicInput = Join-Path $publicRoot 'controls.data'
    Set-SyntheticFile $publicInput $public.Time.AddMinutes(-10) 'synthetic public control'
    $public.Job.Reads = @($publicInput)
    $public.Record.Inputs = @{ $publicInput = (Get-BatchFileStamp $publicInput) }
    $portablePlan = ConvertTo-PortableBatchPlan $public.Plan
    $assessment = Get-BatchOutputStatus -Plan $portablePlan -State $public.State -DateRoot $public.DateRoot -AllowedRoots @($public.DateRoot, $publicRoot)
    Assert-OutputStatus ((Get-EffortRow $assessment).OutputStatus -eq 'Verified') 'portable PUBLIC reads use the declared public root'
    $portablePlan.Jobs[0].ReadFiles = @([pscustomobject]@{ Root = 'External'; Identity = 'opaque external input' })
    $row = Get-EffortRow (Get-BatchOutputStatus -Plan $portablePlan -State $public.State -DateRoot $public.DateRoot -AllowedRoots @($public.DateRoot, $publicRoot))
    Assert-OutputStatus ($row.OutputStatus -eq 'Unverifiable') 'opaque external input identities cannot silently verify freshness'

    foreach ($status in @('Pending', 'Running')) {
        $active = New-OutputFixture "active-$status"
        $active.Record.Status = $status; $active.State.Status = 'Running'
        $assessment = Get-FixtureAssessment $active
        Assert-OutputStatus ((Get-EffortRow $assessment).OutputStatus -ne 'Verified' -and -not $assessment.AllVerified) "$status job is never reported verified"
    }
    $partialRun = New-OutputFixture 'partial-run-valid-row'
    $partialRun.State.Status = 'PartialOrFailed'; $partialRun.State.ExitCode = 1
    $assessment = Get-FixtureAssessment $partialRun
    Assert-OutputStatus ((Get-EffortRow $assessment).OutputStatus -eq 'Verified' -and -not $assessment.AllVerified) 'individual valid output does not turn a partial run into fully verified success'

    $console = (& { Show-BatchOutputStatus -Assessment $blockedAssessment } 6>&1 | Out-String)
    Assert-OutputStatus ($console -match 'BD/Effort' -and $console -match 'NotUpdated') 'console visibly identifies stale Effort'
    Assert-OutputStatus ($console -match 'BD/MultiRole' -and $console -match 'key.*match') 'console exposes the root producer error'
    Assert-OutputStatus ($console -match '(?i)saved|utc') 'console includes saved-file evidence rather than status alone'
    $reportDirectory = Join-Path $blocked.Directory 'report'
    [void][IO.Directory]::CreateDirectory($reportDirectory)
    $reportStateBefore = $blocked.State | ConvertTo-Json -Depth 30 -Compress
    $null = Write-BatchOutputStatusReport -Plan $blocked.Plan -State $blocked.State -Directory $reportDirectory -DateRoot $blocked.DateRoot -AllowedRoots @($blocked.DateRoot)
    $jsonPath = Join-Path $reportDirectory 'output-status.json'
    $markdownPath = Join-Path $reportDirectory 'output-status.md'
    Assert-OutputStatus ((Test-Path -LiteralPath $jsonPath) -and (Test-Path -LiteralPath $markdownPath)) 'JSON and Markdown output-readiness reports are written'
    $jsonReport = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
    Assert-OutputStatus ($jsonReport.RunId -eq $blocked.State.RunId -and -not $jsonReport.AllVerified) 'JSON report retains run identity and does not claim full verification'
    $reportRow = Get-EffortRow $jsonReport
    Assert-OutputStatus ($reportRow.OutputStatus -eq 'ChangedAfterRun' -and $reportRow.RunStatus -eq 'Blocked') 'report recomputes current manual-save evidence without repairing history'
    $markdown = Get-Content -LiteralPath $markdownPath -Raw
    Assert-OutputStatus ($markdown -match 'BD/Effort' -and $markdown -match 'BD/MultiRole' -and $markdown -match '(?i)unverif|not.*verif|metadata') 'Markdown report exposes the output, root cause and evidence limitation'
    Assert-OutputStatus (($blocked.State | ConvertTo-Json -Depth 30 -Compress) -ceq $reportStateBefore) 'report writing never mutates the run ledger'

    $confirmedReportDirectory = Join-Path $fixture.Directory 'confirmed-report'
    $confirmedReport = Write-BatchOutputStatusReport -Plan $fixture.Plan -State $fixture.State -Directory $confirmedReportDirectory -DateRoot $fixture.DateRoot -AllowedRoots @($fixture.DateRoot)
    $confirmedMarkdown = Get-Content -LiteralPath (Join-Path $confirmedReportDirectory 'output-status.md') -Raw
    Assert-OutputStatus ($confirmedReport.AllVerified -and $confirmedMarkdown -notmatch 'OUTPUTS NOT VERIFIED/UPDATED') 'fully verified report uses a positive header without a stale-output false alarm'

    $afterExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Assert-OutputStatus (@($afterExcel | Where-Object { $_ -notin $beforeExcel }).Count -eq 0) 'metadata assessment and report tests start no Excel process'
    Write-Host "PASS: $script:checks batch output-status assertions; synthetic metadata only, no Excel or workbook-content reads."
}
finally {
    [Environment]::SetEnvironmentVariable('PUBLIC', $originalPublic, 'Process')
    # Delete only this test's verified disposable directory under the system temporary root.
    $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
    $temporaryParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]]'\/')
    if (-not $resolvedTestRoot.StartsWith($temporaryParent + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $resolvedTestRoot -Leaf) -notlike 'rc-output-status-*') {
        throw 'Refusing cleanup outside the generated output-status fixture directory.'
    }
    foreach ($junction in $junctions) {
        if (-not [IO.Path]::GetFullPath($junction).StartsWith($resolvedTestRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Refusing cleanup of an alias outside the generated fixture directory.'
        }
        if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction, $false) }
    }
    if (Test-Path -LiteralPath $resolvedTestRoot) { Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force }
}
