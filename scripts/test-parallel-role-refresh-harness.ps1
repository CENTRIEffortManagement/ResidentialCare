param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path $PSScriptRoot -Parent
$coordinator = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/parallel-test/Invoke-ParallelRoleRefreshTest.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-parallel-test-' + [guid]::NewGuid().ToString('N'))
$runId = 'mock-full'
$beforeExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
try {
    # Isolated metadata-only fixtures: never require production TestRole workbooks.
    $fixtureDate = Join-Path $testRoot 'fixture Date'
    $fixtureUnits = Join-Path $fixtureDate 'UNITS'
    [void][IO.Directory]::CreateDirectory((Join-Path $fixtureUnits 'parallel-test'))
    [void][IO.Directory]::CreateDirectory((Join-Path $fixtureUnits 'runner/src'))
    [void][IO.Directory]::CreateDirectory((Join-Path $fixtureDate 'runner/src'))
    foreach ($name in @('Invoke-ParallelRoleRefreshTest.ps1','Invoke-ParallelRoleLane.ps1','ParallelRoleTest.psd1')) {
        Copy-Item -LiteralPath (Join-Path (Split-Path $coordinator -Parent) $name) -Destination (Join-Path $fixtureUnits 'parallel-test')
    }
    Copy-Item -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/runner/src/Invoke-ExcelWorkbookRefresh.ps1') -Destination (Join-Path $fixtureUnits 'runner/src')
    Copy-Item -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1') -Destination (Join-Path $fixtureDate 'runner/src')
    $fixtureConfig = Import-PowerShellDataFile (Join-Path $fixtureUnits 'parallel-test/ParallelRoleTest.psd1')
    foreach ($folder in $fixtureConfig.Roles.Values) {
        $roleDirectory = Join-Path $fixtureUnits "$($fixtureConfig.Unit)/$folder"
        [void][IO.Directory]::CreateDirectory($roleDirectory)
        foreach ($name in $fixtureConfig.WorkbookOrder) { [IO.File]::WriteAllText((Join-Path $roleDirectory $name), 'Synthetic mock metadata fixture; not an Excel workbook.') }
    }
    $coordinator = Join-Path $fixtureUnits 'parallel-test/Invoke-ParallelRoleRefreshTest.ps1'
    & (Get-Command pwsh -ErrorAction Stop).Source -NoProfile -ExecutionPolicy Bypass -File $coordinator `
        -Mode Full -MockWorkers -RunId $runId -RunRootOverride $testRoot
    if ($LASTEXITCODE -ne 0) { throw "Mock parallel coordinator exited with $LASTEXITCODE." }
    foreach ($role in @('AIN', 'AINC4', 'RN', 'TestRole1', 'TestRole2', 'TestRole3', 'TestRole4')) {
        $resultPath = Join-Path $testRoot "$runId/$role/result.json"
        if (-not (Test-Path -LiteralPath $resultPath)) { throw "Missing mock result: $resultPath" }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($result.ExitCode -ne 0 -or @($result.Workbooks).Count -ne 3) { throw "Unexpected mock result for $role." }
    }
    $newExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object Id -notin $beforeExcel)
    if ($newExcel.Count -gt 0) { throw 'Mock test unexpectedly started Excel.' }
    Write-Host 'PASS: seven isolated three-workbook lanes completed with mock workers and no Excel process.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        $resolved = (Resolve-Path -LiteralPath $testRoot).Path
        if (-not $resolved.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing to remove a mock test path outside the temp directory.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
