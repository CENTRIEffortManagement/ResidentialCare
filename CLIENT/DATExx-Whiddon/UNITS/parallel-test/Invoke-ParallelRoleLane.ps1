param(
    [Parameter(Mandatory = $true)] [string] $ConfigPath,
    [Parameter(Mandatory = $true)] [string] $RunDirectory,
    [Parameter(Mandatory = $true)] [string] $Role,
    [ValidateSet('Pair', 'Full')] [string] $Mode = 'Full',
    [Parameter(Mandatory = $true)] [string] $WorkerScript,
    [switch] $MockWorkers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$config = Import-PowerShellDataFile -LiteralPath $ConfigPath
$unitsRoot = Split-Path $PSScriptRoot -Parent
$unitRoot = Join-Path $unitsRoot $config.Unit
$laneDirectory = Join-Path $RunDirectory $Role
$laneLog = Join-Path $laneDirectory 'lane.log'
$laneStatus = Join-Path $laneDirectory 'current-status.txt'
$laneStop = Join-Path $laneDirectory 'stop-request.txt'
$resultPath = Join-Path $laneDirectory 'result.json'
[void] [IO.Directory]::CreateDirectory($laneDirectory)

function Write-LaneLog {
    param([string] $Message, [string] $Level = 'INFO')
    $line = '[{0}] [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Role, $Message
    Add-Content -LiteralPath $laneLog -Value $line
}

function Write-LaneStatus {
    param([string] $State, [string] $Workbook = '', [string] $Message = '')
    $lines = @("state=$State", "timestamp=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "role=$Role")
    if ($Workbook) { $lines += "currentWorkbook=$Workbook" }
    if ($Message) { $lines += "message=$Message" }
    Set-Content -LiteralPath $laneStatus -Value $lines
}

function Get-StopMode {
    if (-not (Test-Path -LiteralPath $laneStop -PathType Leaf)) { return $null }
    foreach ($line in @(Get-Content -LiteralPath $laneStop -ErrorAction SilentlyContinue)) {
        if ($line -match '^mode=(Now|AfterCurrent)$') { return $matches[1] }
    }
    return $null
}

function Test-SavedCompletionState {
    param([string] $StatusPath)
    if (-not (Test-Path -LiteralPath $StatusPath -PathType Leaf)) { return $false }
    $statusText = Get-Content -LiteralPath $StatusPath -Raw -ErrorAction SilentlyContinue
    if ($statusText -notmatch 'Supervisor exit 1 during Quitting; workbook save was confirmed\. State: (.+)') { return $false }
    $statePath = $matches[1].Trim()
    $completeStatePath = "$statePath.tmp"
    if (-not (Test-Path -LiteralPath $completeStatePath -PathType Leaf)) { return $false }
    try {
        $completeState = Get-Content -LiteralPath $completeStatePath -Raw | ConvertFrom-Json
        if ($completeState.phase -ne 'Complete' -or -not [bool] $completeState.saved) { return $false }
        $ownedProcess = Get-Process -Id ([int] $completeState.excel.id) -ErrorAction SilentlyContinue
        if ($null -ne $ownedProcess -and $ownedProcess.StartTime.Ticks -eq [long] $completeState.excel.startTicks) { return $false }
        return $true
    }
    catch {
        return $false
    }
}

function Test-RetryablePreOpenStateFailure {
    param([string] $LogPath)
    if (-not (Test-Path -LiteralPath $LogPath -PathType Leaf)) { return $false }
    $logText = Get-Content -LiteralPath $LogPath -Raw -ErrorAction SilentlyContinue
    return (
        $logText -match 'Exception calling "Move" with "3" argument\(s\): "Access to the path is denied\."' -and
        $logText -notmatch 'Owned Excel PID:' -and
        $logText -notmatch 'Workbook opened\.'
    )
}

if (-not $config.Roles.ContainsKey($Role)) { throw "Role '$Role' is not configured for the parallel test." }
$roleFolder = [string] $config.Roles[$Role]
$workbookNames = @($config.WorkbookOrder)
if ($Mode -eq 'Pair') { $workbookNames = @($workbookNames[0]) }
$results = [Collections.Generic.List[object]]::new()
$exitCode = 1

try {
    Write-LaneStatus -State 'RUNNING' -Message "$Mode lane starting."
    Write-LaneLog "$Mode lane starting with $($workbookNames.Count) workbook(s)."
    for ($index = 0; $index -lt $workbookNames.Count; $index++) {
        $stopMode = Get-StopMode
        if ($stopMode -eq 'Now' -or $stopMode -eq 'AfterCurrent') {
            Write-LaneLog "Stop request observed before the next workbook." 'WARN'
            $exitCode = 3
            break
        }
        $fileName = [string] $workbookNames[$index]
        $workbookPath = Join-Path (Join-Path $unitRoot $roleFolder) $fileName
        $displayName = "$Role / $fileName"
        $sequence = $index + 1
        $workbookLog = Join-Path $laneDirectory ('{0:D2}-{1}.log' -f $sequence, ([IO.Path]::GetFileNameWithoutExtension($fileName)))
        $workbookStatus = Join-Path $laneDirectory ('{0:D2}-status.txt' -f $sequence)
        Write-LaneStatus -State 'RUNNING' -Workbook $displayName -Message "Workbook $sequence/$($workbookNames.Count)."
        Write-LaneLog "Starting workbook $sequence/$($workbookNames.Count): $workbookPath"
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $completionNote = $null
        $retryCount = 0
        if ($MockWorkers) {
            Start-Sleep -Seconds 1
            $workerExit = 0
            Set-Content -LiteralPath $workbookLog -Value "MOCK SUCCESS: $workbookPath"
        }
        else {
            do {
                & (Get-Command pwsh -ErrorAction Stop).Source -NoProfile -ExecutionPolicy Bypass -File $WorkerScript `
                    -WorkbookPath $workbookPath -WorkbookName $displayName -VisibleOverride false `
                    -TimeoutMinutesOverride ([string] $config.TimeoutMinutes) -LogPath $workbookLog `
                    -StatusPath $workbookStatus -StatusPrefix "[$Role $sequence/$($workbookNames.Count)]" `
                    -StopRequestPath $laneStop
                $workerExit = $LASTEXITCODE
                if ($workerExit -eq 1 -and $retryCount -lt 2 -and (Test-RetryablePreOpenStateFailure -LogPath $workbookLog)) {
                    $retryCount++
                    Write-LaneLog "Retrying workbook after pre-open state-file access failure (retry $retryCount/2)." 'WARN'
                    Start-Sleep -Seconds 2
                    continue
                }
                break
            } while ($true)
            if ($workerExit -eq 1 -and (Test-SavedCompletionState -StatusPath $workbookStatus)) {
                $completionNote = 'Recovered supervisor state-file race after confirmed save, close, and owned Excel exit.'
                Write-LaneLog $completionNote 'WARN'
                $workerExit = 0
            }
        }
        $timer.Stop()
        $results.Add([pscustomobject] @{
            Role = $Role
            Sequence = $sequence
            Workbook = $fileName
            Path = $workbookPath
            ExitCode = $workerExit
            ElapsedSeconds = [math]::Round($timer.Elapsed.TotalSeconds, 2)
            LogPath = $workbookLog
            CompletionNote = $completionNote
            RetryCount = $retryCount
        })
        if ($workerExit -ne 0) {
            Write-LaneLog "Workbook failed or stopped with exit code ${workerExit}: $displayName" 'ERROR'
            $exitCode = $workerExit
            break
        }
        Write-LaneLog "Workbook completed in $([math]::Round($timer.Elapsed.TotalSeconds, 1)) second(s): $displayName"
        if ((Get-StopMode) -eq 'AfterCurrent') {
            Write-LaneLog 'Stop-after-current request observed.' 'WARN'
            $exitCode = 3
            break
        }
    }
    if ($results.Count -eq $workbookNames.Count -and @($results | Where-Object ExitCode -ne 0).Count -eq 0) {
        $exitCode = 0
        Write-LaneStatus -State 'SUCCEEDED' -Message "$($results.Count) workbook(s) completed."
        Write-LaneLog 'Lane succeeded.'
    }
    elseif ($exitCode -eq 3) {
        Write-LaneStatus -State 'STOPPED' -Message "$($results.Count) workbook(s) attempted."
    }
    else {
        Write-LaneStatus -State 'FAILED' -Message "$($results.Count) workbook(s) attempted."
    }
}
catch {
    Write-LaneLog $_.Exception.Message 'ERROR'
    Write-LaneStatus -State 'FAILED' -Message $_.Exception.Message
    $exitCode = 1
}
finally {
    [pscustomobject] @{ Role = $Role; Mode = $Mode; ExitCode = $exitCode; Workbooks = @($results) } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resultPath
}

exit $exitCode
