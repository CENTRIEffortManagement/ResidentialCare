param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-unit-discovery-' + [guid]::NewGuid().ToString('N'))
$fixtureRepo = Join-Path $testRoot 'repo'
$fixtureRunner = Join-Path $fixtureRepo 'CLIENT/DATExx-Whiddon/runner'
$fixtureUnits = Join-Path $fixtureRepo 'CLIENT/DATExx-Whiddon/UNITS'
function Assert-Discovery {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    Write-Host "PASS: $Message"
}
try {
    [void] [IO.Directory]::CreateDirectory($fixtureRunner)
    Copy-Item -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Destination $fixtureRepo
    # Configuration-only fixtures: no workbooks are copied or opened.
    Get-ChildItem -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner') -Filter '*.psd1' -File |
        Copy-Item -Destination $fixtureRunner
    foreach ($unit in @('BD', 'TE', 'JH-RY')) {
        [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureUnits $unit))
    }
    foreach ($sequenceProfile in @('Current', 'ParallelInputs-Unit1Priority')) {
        $catalogue = Get-BatchCatalogue $fixtureRepo $sequenceProfile
        $plan = Select-BatchPlan $catalogue -RunAll
        Assert-Discovery (($catalogue.Units -join ',') -eq 'BD,TE,JH-RY') "$sequenceProfile discovers facilities without UnitN folders"
        Assert-Discovery ($plan.Jobs.Count -eq 90) "$sequenceProfile builds the complete default plan including CS without retired Workers jobs"
        $csPlan = Select-BatchPlan $catalogue -Roles CS
        Assert-Discovery ($csPlan.Jobs.Count -eq 9) "$sequenceProfile selects all nine facility CS workbooks by role"
        foreach ($unit in @('BD', 'TE', 'JH-RY')) {
            $c1 = Select-BatchPlan $catalogue -Units $unit -Batches C1
            Assert-Discovery (($c1.Jobs.Id -join ',') -eq "$unit/Availability,$unit/Staff") "$sequenceProfile runs $unit Availability then Staff without Worker Reconciliation"
            Assert-Discovery (@($plan.Jobs | Where-Object Id -eq "$unit/Workers").Count -eq 0 -and
                @($plan.Jobs | ForEach-Object Dependencies | Where-Object { $_ -eq "$unit/Workers" }).Count -eq 0 -and
                @($plan.Jobs | ForEach-Object Reads | Where-Object { $_ -like '*Worker Reconciliation.xlsx' }).Count -eq 0) "$sequenceProfile has no retired Workers job, dependency or file read"
            $allocation = @($plan.Jobs | Where-Object Id -eq "$unit/AllocationInput")[0]
            $settings = @($plan.Jobs | Where-Object Id -eq "$unit/Settings")[0]
            foreach ($job in $c1.Jobs) {
                Assert-Discovery ($allocation.Id -in $job.Dependencies -and $settings.Id -in $job.Dependencies -and
                    $allocation.Path -in $job.Reads -and $settings.Path -in $job.Reads) "$sequenceProfile retains Allocation and Settings ordering and input checks for $($job.Id)"
            }
            Assert-Discovery ("$unit/Availability" -in $c1.Jobs[1].Dependencies -and
                $c1.Jobs[0].Path -in $c1.Jobs[1].Reads) "$sequenceProfile releases $unit Staff only after Availability completes"
            $csJobs = @($plan.Jobs | Where-Object { $_.Unit -eq $unit -and $_.Role -eq 'CS' })
            Assert-Discovery (($csJobs.Id -join ',') -eq "$unit/Role:CS/1,$unit/Role:CS/2,$unit/Role:CS/3") "$sequenceProfile includes $unit CS in A.1, A.2, B order"
            Assert-Discovery ("$unit/Role:CS/1" -in $csJobs[1].Dependencies -and "$unit/Role:CS/2" -in $csJobs[2].Dependencies) "$sequenceProfile keeps $unit CS workbooks sequential"
            $capacity = @($plan.Jobs | Where-Object Id -eq "$unit/Capacity")[0]
            Assert-Discovery ("$unit/Role:CS/3" -in $capacity.Dependencies -and $csJobs[2].Path -in $capacity.Reads) "$sequenceProfile makes $unit consolidation wait for and read CS output"
        }
        Assert-Discovery (@($plan.Jobs | Where-Object { $_.Id -match '/MultiRole$' -or $_.RelativePath -match 'MutliRole|MultiRole' }).Count -eq 0) "$sequenceProfile excludes MultiRole from every discovered facility"
        Assert-Discovery (($plan.Jobs | Select-Object -Last 3 | ForEach-Object Id) -join ',' -eq 'Org/LeaveBalanceStats,Org/LeaveBalanceReduction,Org/LeaveBalanceReporting') "$sequenceProfile retains Leave Balance ordering"
    }
    [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureUnits 'Unit1'))
    $catalogue = Get-BatchCatalogue $fixtureRepo
    Assert-Discovery (($catalogue.Units -join ',') -eq 'Unit1,BD,TE,JH-RY') 'legacy Units remain available when present'
    Assert-Discovery ((Select-BatchPlan $catalogue -RunAll -Units Unit1).Jobs.Count -eq 25) 'explicit legacy Unit selection still works'
    foreach ($sequenceProfile in @('Current', 'ParallelInputs-Unit1Priority')) {
        $legacyCatalogue = Get-BatchCatalogue $fixtureRepo $sequenceProfile
        $legacyC1 = Select-BatchPlan $legacyCatalogue -Units Unit1 -Batches C1
        Assert-Discovery (($legacyC1.Jobs.Id -join ',') -eq 'Unit1/Workers,Unit1/Availability,Unit1/Staff' -and
            'Unit1/Workers' -in $legacyC1.Jobs[1].Dependencies -and
            $legacyC1.Jobs[0].Path -in $legacyC1.Jobs[1].Reads) "$sequenceProfile retains the legacy Workers prerequisite and input check"
    }
    $settingsPath = Join-Path $fixtureRunner 'ResidentialCare-BatchApproval.psd1'
    $settingsText = [IO.File]::ReadAllText($settingsPath)
    [IO.File]::WriteAllText($settingsPath, $settingsText.Replace('UseWorkerReconciliation = $false', "UseWorkerReconciliation = 'false'"))
    $message = ''
    try { Get-BatchCatalogue $fixtureRepo | Out-Null } catch { $message = $_.Exception.Message }
    Assert-Discovery ($message -eq 'UseWorkerReconciliation must be boolean: BD') 'text values cannot silently retire Workers jobs'
    [IO.File]::WriteAllText($settingsPath, $settingsText.Replace('; UseWorkerReconciliation = $false', ''))
    $defaultMembership = Get-BatchCatalogue $fixtureRepo
    Assert-Discovery (@($defaultMembership.Jobs | Where-Object Id -in @('BD/Workers','TE/Workers','JH-RY/Workers')).Count -eq 3) 'omitting the facility setting preserves existing Workers jobs'
    [IO.File]::WriteAllText($settingsPath, $settingsText.Replace("Folder = 'BD'", "Folder = 'MissingFacility'"))
    $message = ''
    try { Get-BatchCatalogue $fixtureRepo | Out-Null } catch { $message = $_.Exception.Message }
    Assert-Discovery ($message -eq 'Additional Unit folder is missing: MissingFacility') 'missing configured facilities still block planning'
} finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]] '\/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-unit-discovery-*') { throw 'Unsafe fixture cleanup target.' }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
