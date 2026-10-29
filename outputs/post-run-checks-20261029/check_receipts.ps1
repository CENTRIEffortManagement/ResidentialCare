$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchReporting.ps1')
$project = Get-Content -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Raw | ConvertFrom-Json
$dateRoot = Resolve-BatchPath $repoRoot $project.batchRunner.dateRoot
$roots = @($dateRoot)
foreach ($root in $project.batchRunner.workbookRoots.PSObject.Properties) {
    $roots += Resolve-BatchWorkbookRoot $repoRoot ([string]$root.Value)
}
$chosen = @{}
$summaries = @()
$jobIndex = @{}
$recordIndex = @{}
$basisPlan = $null
foreach ($runId in @('20261029-074441-a2fe3870', '20261029-085859-fbb9b367')) {
    $runDir = Join-Path $dateRoot "RunLogs/BatchRefresh/$runId"
    $plan = Get-Content -LiteralPath (Join-Path $runDir 'plan.json') -Raw | ConvertFrom-Json
    $state = Get-Content -LiteralPath (Join-Path $runDir 'state.json') -Raw | ConvertFrom-Json -AsHashtable
    if ($null -eq $basisPlan) { $basisPlan = $plan }
    if ($plan.Fingerprint -ne $basisPlan.Fingerprint -or $plan.LocationIdentity -ne $basisPlan.LocationIdentity) {
        throw 'Cannot combine receipts from different configurations or locations.'
    }
    $assessment = Get-BatchOutputStatus -Plan $plan -State $state -DateRoot $dateRoot -AllowedRoots $roots
    foreach ($job in $plan.Jobs) {
        if ($job.Id -in @('Org/Cost', 'Org/History', 'Org/HistoryRead')) { continue }
        $jobIndex[$job.Id] = $job
        $recordIndex[$job.Id] = $state.Jobs[$job.Id]
        $chosen[$job.Id] = [pscustomobject]@{
            Id = $job.Id; Path = $job.RelativePath; Run = $runId; Record = $state.Jobs[$job.Id]
            Metadata = @($assessment.Rows | Where-Object Id -eq $job.Id)[0]
        }
    }
    $summaries += [pscustomobject]@{
        Run = $runId; Status = $state.Status
        Rows = @($assessment.Rows | Where-Object Id -notin @('Org/Cost','Org/History','Org/HistoryRead'))
    }
}
# This is a derived read-only assessment, never a replacement run or ledger.
if (@($recordIndex.Values | Where-Object Status -ne 'Completed').Count) {
    throw 'The combined scope contains incomplete jobs.'
}
$derivedPlan = @{
    Fingerprint = $basisPlan.Fingerprint; LocationIdentity = $basisPlan.LocationIdentity
    Jobs = @($jobIndex.Values)
}
$derivedState = @{
    Fingerprint = $basisPlan.Fingerprint; Jobs = $recordIndex; Status = 'Completed'
    RunId = 'Derived audit of two runs; not a runner run'; Updated = [datetime]::UtcNow.ToString('o')
}
$combined = Get-BatchOutputStatus -Plan $derivedPlan -State $derivedState -DateRoot $dateRoot -AllowedRoots $roots
$combined | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'combined-metadata.json')
foreach ($row in $combined.Rows) { $chosen[$row.Id].Metadata = $row }
$chosen.Values | Sort-Object Id | ConvertTo-Json -Depth 12 |
    Set-Content -LiteralPath (Join-Path $PSScriptRoot 'selected-workbooks.json')
$summaries | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'run-metadata.json')
$chosen.Values | Group-Object { $_.Metadata.OutputStatus } | Select-Object Name,Count | ConvertTo-Json
if (@($chosen.Values | Where-Object { $_.Metadata.OutputStatus -ne 'Verified' }).Count) { exit 1 }
