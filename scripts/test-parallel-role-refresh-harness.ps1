param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path $PSScriptRoot -Parent
$coordinator = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/parallel-test/Invoke-ParallelRoleRefreshTest.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-parallel-test-' + [guid]::NewGuid().ToString('N'))
$runId = 'mock-full'
$beforeExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
try {
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
