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
        Assert-Discovery ($plan.Jobs.Count -eq 84) "$sequenceProfile builds the complete default plan"
        Assert-Discovery (@($plan.Jobs | Where-Object { $_.Id -match '/MultiRole$' -or $_.RelativePath -match 'MutliRole|MultiRole' }).Count -eq 0) "$sequenceProfile excludes MultiRole from every discovered facility"
        Assert-Discovery (($plan.Jobs | Select-Object -Last 3 | ForEach-Object Id) -join ',' -eq 'Org/LeaveBalanceStats,Org/LeaveBalanceReduction,Org/LeaveBalanceReporting') "$sequenceProfile retains Leave Balance ordering"
    }
    [void] [IO.Directory]::CreateDirectory((Join-Path $fixtureUnits 'Unit1'))
    $catalogue = Get-BatchCatalogue $fixtureRepo
    Assert-Discovery (($catalogue.Units -join ',') -eq 'Unit1,BD,TE,JH-RY') 'legacy Units remain available when present'
    Assert-Discovery ((Select-BatchPlan $catalogue -RunAll -Units Unit1).Jobs.Count -eq 25) 'explicit legacy Unit selection still works'
    $settingsPath = Join-Path $fixtureRunner 'ResidentialCare-BatchApproval.psd1'
    $settingsText = [IO.File]::ReadAllText($settingsPath)
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
