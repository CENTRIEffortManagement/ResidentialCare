param(
    [ValidateSet('Pair', 'Full')] [string] $Mode = 'Pair',
    [string] $RoleA = 'AIN',
    [string] $RoleB = 'AINC4',
    [string] $RoleC = 'RN',
    [string] $RoleD = 'TestRole1',
    [string] $RoleE = 'TestRole2',
    [string] $RoleF = 'TestRole3',
    [string] $RoleG = 'TestRole4',
    [switch] $ValidateOnly,
    [switch] $MockWorkers,
    [ValidateSet('Now', 'AfterCurrent')] [string] $StopMode,
    [string] $RunId,
    [string] $RunRootOverride
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$configPath = Join-Path $PSScriptRoot 'ParallelRoleTest.psd1'
$laneScript = Join-Path $PSScriptRoot 'Invoke-ParallelRoleLane.ps1'
$unitsRoot = Split-Path $PSScriptRoot -Parent
$workerScript = Join-Path $unitsRoot 'runner/src/Invoke-ExcelWorkbookRefresh.ps1'
$logRoot = if ($RunRootOverride) { $RunRootOverride } else { Join-Path $unitsRoot 'RunLogs/ParallelRoleTests' }
$currentRunPath = Join-Path $logRoot 'current-run.txt'
$parallelLockPath = Join-Path $logRoot 'ParallelRoleRefreshTest.lock'
Add-Type -AssemblyName Microsoft.VisualBasic
$computerInfo = [Microsoft.VisualBasic.Devices.ComputerInfo]::new()
$preExistingRefreshProcessIds = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    $_.ProcessName -eq 'EXCEL' -or $_.ProcessName -like 'Microsoft.Mashup.Container*'
} | ForEach-Object Id)
$laneRoles = @($RoleA, $RoleB, $RoleC, $RoleD, $RoleE, $RoleF, $RoleG)

function Write-MemorySample {
    param([string] $Path)
    $totalMB = $computerInfo.TotalPhysicalMemory / 1MB
    $availableMB = $computerInfo.AvailablePhysicalMemory / 1MB
    $usedPercent = 100 * ($totalMB - $availableMB) / $totalMB
    $newRefreshProcesses = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
        ($_.ProcessName -eq 'EXCEL' -or $_.ProcessName -like 'Microsoft.Mashup.Container*') -and
        $_.Id -notin $preExistingRefreshProcessIds
    })
    $newRefreshWorkingSetMB = if ($newRefreshProcesses.Count -eq 0) {
        0
    }
    else {
        ($newRefreshProcesses | Measure-Object -Property WorkingSet64 -Sum).Sum / 1MB
    }
    $invariant = [Globalization.CultureInfo]::InvariantCulture
    Add-Content -LiteralPath $Path -Value ('{0},{1},{2},{3},{4}' -f
        (Get-Date).ToString('o'),
        ([math]::Round($usedPercent, 2).ToString($invariant)),
        ([math]::Round($availableMB).ToString($invariant)),
        ([math]::Round($newRefreshWorkingSetMB, 1).ToString($invariant)),
        $newRefreshProcesses.Count)
}

function Write-StopFiles {
    param([string] $ActiveRunDirectory, [string] $ModeToWrite)
    foreach ($role in $laneRoles) {
        $laneDirectory = Join-Path $ActiveRunDirectory $role
        if (Test-Path -LiteralPath $laneDirectory -PathType Container) {
            Set-Content -LiteralPath (Join-Path $laneDirectory 'stop-request.txt') -Value @(
                "mode=$ModeToWrite", "requestedAt=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", 'source=Parallel role refresh test'
            )
        }
    }
}

[void] [IO.Directory]::CreateDirectory($logRoot)
if ($StopMode) {
    if (-not (Test-Path -LiteralPath $currentRunPath -PathType Leaf)) { throw 'No active parallel role test is registered.' }
    $activeRun = (Get-Content -LiteralPath $currentRunPath -Raw).Trim()
    Write-StopFiles -ActiveRunDirectory $activeRun -ModeToWrite $StopMode
    Write-Host "Parallel stop request written: $StopMode"
    exit 0
}

foreach ($required in @($configPath, $laneScript, $workerScript)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "Required file is missing: $required" }
}
$config = Import-PowerShellDataFile -LiteralPath $configPath
if (@($laneRoles | Sort-Object -Unique).Count -ne $laneRoles.Count) { throw 'The parallel roles must be different.' }
if ($config.MaxConcurrency -ne $laneRoles.Count) { throw "Parallel test configuration must limit concurrency to $($laneRoles.Count)." }
$unitRoot = Join-Path $unitsRoot $config.Unit
$productionLockPath = Join-Path $unitsRoot 'RunLogs/AllUnitsRefresh.lock'
if (Test-Path -LiteralPath $productionLockPath -PathType Leaf) { throw 'The production all-units runner is active. Parallel testing was not started.' }
$allPaths = [Collections.Generic.List[string]]::new()
$selectedNames = if ($Mode -eq 'Pair') { @($config.WorkbookOrder[0]) } else { @($config.WorkbookOrder) }
foreach ($role in $laneRoles) {
    if (-not $config.Roles.ContainsKey($role)) { throw "Role '$role' is not configured for the parallel test." }
    $roleFolder = [string] $config.Roles[$role]
    foreach ($fileName in $selectedNames) {
        $path = Join-Path (Join-Path $unitRoot $roleFolder) ([string] $fileName)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Test workbook is missing: $path" }
        $resolved = (Resolve-Path -LiteralPath $path).Path
        $lockPath = Join-Path (Split-Path $resolved -Parent) ('~$' + (Split-Path $resolved -Leaf))
        if (Test-Path -LiteralPath $lockPath -PathType Leaf) { throw "Workbook has an Excel lock file: $resolved" }
        $allPaths.Add($resolved)
    }
}
if (@($allPaths | Sort-Object -Unique).Count -ne $allPaths.Count) { throw 'The parallel selection contains a duplicate full workbook path.' }

Write-Host "Parallel role refresh test validation passed: $Mode"
Write-Host "Lane A: $RoleA"
Write-Host "Lane B: $RoleB"
Write-Host "Lane C: $RoleC"
Write-Host "Lane D: $RoleD"
Write-Host "Lane E: $RoleE"
Write-Host "Lane F: $RoleF"
Write-Host "Lane G: $RoleG"
foreach ($path in $allPaths) { Write-Host $path }
if ($ValidateOnly) { exit 0 }

$lockStream = $null
$processes = @{}
$runDirectory = $null
$coordinatorLog = $null
try {
    $lockStream = [IO.File]::Open($parallelLockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    if (-not $RunId) { $RunId = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6) }
    $runDirectory = Join-Path $logRoot $RunId
    if (Test-Path -LiteralPath $runDirectory) { throw "Parallel test run directory already exists: $runDirectory" }
    [void] [IO.Directory]::CreateDirectory($runDirectory)
    $coordinatorLog = Join-Path $runDirectory 'coordinator.log'
    $memoryLog = Join-Path $runDirectory 'memory.csv'
    Set-Content -LiteralPath $memoryLog -Value 'Timestamp,UsedPercent,AvailableMB,NewRefreshProcessWorkingSetMB,NewRefreshProcessCount'
    Write-MemorySample -Path $memoryLog
    Set-Content -LiteralPath $currentRunPath -Value $runDirectory
    Add-Content -LiteralPath $coordinatorLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting $Mode test: $($laneRoles -join ' + ')"
    foreach ($role in $laneRoles) {
        $startInfo = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh -ErrorAction Stop).Source)
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        foreach ($argument in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $laneScript,
            '-ConfigPath', $configPath, '-RunDirectory', $runDirectory, '-Role', $role,
            '-Mode', $Mode, '-WorkerScript', $workerScript)) { $startInfo.ArgumentList.Add([string] $argument) }
        if ($MockWorkers) { $startInfo.ArgumentList.Add('-MockWorkers') }
        $processes[$role] = [Diagnostics.Process]::Start($startInfo)
    }
    $failureStopSent = $false
    while (@($processes.Values | Where-Object { -not $_.HasExited }).Count -gt 0) {
        Write-MemorySample -Path $memoryLog
        $parts = foreach ($role in $laneRoles) {
            $statusPath = Join-Path (Join-Path $runDirectory $role) 'current-status.txt'
            $state = 'STARTING'
            $workbook = ''
            if (Test-Path -LiteralPath $statusPath) {
                foreach ($line in @(Get-Content -LiteralPath $statusPath -ErrorAction SilentlyContinue)) {
                    if ($line -match '^state=(.+)$') { $state = $matches[1] }
                    if ($line -match '^currentWorkbook=(.+)$') { $workbook = $matches[1] }
                }
            }
            "$role=$state$(if ($workbook) { " ($workbook)" })"
        }
        $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " + ($parts -join '; ')
        Write-Host $line
        Add-Content -LiteralPath $coordinatorLog -Value $line
        if (-not $failureStopSent) {
            $failed = @($processes.GetEnumerator() | Where-Object { $_.Value.HasExited -and $_.Value.ExitCode -ne 0 })
            if ($failed.Count -gt 0) {
                Write-StopFiles -ActiveRunDirectory $runDirectory -ModeToWrite 'AfterCurrent'
                $failureStopSent = $true
            }
        }
        Start-Sleep -Seconds 5
    }
    Write-MemorySample -Path $memoryLog
    $exitCode = 0
    foreach ($role in $laneRoles) {
        $process = $processes[$role]
        if ($process.ExitCode -ne 0) { $exitCode = $process.ExitCode }
        $resultPath = Join-Path (Join-Path $runDirectory $role) 'result.json'
        if (Test-Path -LiteralPath $resultPath) {
            $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
            foreach ($workbook in @($result.Workbooks)) {
                Write-Host ("{0} | {1} | exit {2} | {3:n1}s" -f $role, $workbook.Workbook, $workbook.ExitCode, $workbook.ElapsedSeconds)
            }
        }
    }
    $memorySamples = @(Import-Csv -LiteralPath $memoryLog)
    $baselineMemory = [double] $memorySamples[0].UsedPercent
    $peakMemory = [double] (($memorySamples | ForEach-Object { [double] $_.UsedPercent } | Measure-Object -Maximum).Maximum)
    $peakNewRefreshWorkingSet = [double] (($memorySamples | ForEach-Object { [double] $_.NewRefreshProcessWorkingSetMB } | Measure-Object -Maximum).Maximum)
    Write-Host ('Memory: baseline {0:n1}%; peak {1:n1}%; increase {2:n1} points; peak newly observed Excel/Power Query working set {3:n0} MB.' -f
        $baselineMemory, $peakMemory, ($peakMemory - $baselineMemory), $peakNewRefreshWorkingSet)
    if ($exitCode -eq 0) { Write-Host "Parallel role refresh test succeeded: $runDirectory" }
    else { Write-Host "Parallel role refresh test failed or stopped with exit code $exitCode`: $runDirectory" }
    exit $exitCode
}
finally {
    foreach ($process in @($processes.Values)) { if ($null -ne $process) { $process.Dispose() } }
    if ($null -ne $lockStream) { $lockStream.Dispose() }
    if ($runDirectory -and (Test-Path -LiteralPath $currentRunPath)) {
        $registered = (Get-Content -LiteralPath $currentRunPath -Raw -ErrorAction SilentlyContinue).Trim()
        if ($registered -eq $runDirectory) { Remove-Item -LiteralPath $currentRunPath -Force }
    }
}
