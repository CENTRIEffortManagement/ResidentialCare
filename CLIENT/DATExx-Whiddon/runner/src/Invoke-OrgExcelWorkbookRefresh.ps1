param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [string] $WorkbookName = $null,

    [object] $VisibleOverride = $null,

    [object] $TimeoutMinutesOverride = $null,

    [object] $SkipAsyncWait = $null,

    [object] $CleanupGhostExcelProcessesOverride = $null,

    [object] $ForceCloseExcelProcessesOverride = $null,

    [string] $LogPath = $null,

    [string] $StatusPath = $null,

    [string] $StatusPrefix = $null,

    [string] $StopRequestPath = $null
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$DefaultVisible = $false
$DefaultTimeoutMinutes = 30
$DefaultSkipAsyncWait = $true
$DefaultCleanupGhostExcelProcesses = $false
$DefaultForceCloseExcelProcesses = $false
$DefaultOpenSettleSeconds = 5

$script:LogFile = $null
$script:StatusFile = $null
$script:FinalStatus = "Error"
$script:ExitCode = 1
$script:PrimaryError = $null
$script:StopRequested = $false
$script:WorkbookClosed = $false
$script:ExcelQuit = $false

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string] $Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = if ([string]::IsNullOrWhiteSpace($StatusPrefix)) { "" } else { "$StatusPrefix " }
    $line = "[$timestamp] [$Level] $prefix$Message"
    Write-Host $line

    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line
    }
}

function Write-CurrentStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $State,

        [string] $Message = $null
    )

    if ([string]::IsNullOrWhiteSpace($script:StatusFile)) {
        return
    }

    $lines = @(
        "state=$State",
        "timestamp=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "workflow=ResidentialCare Date Refresh",
        "logFile=$script:LogFile"
    )

    if (-not [string]::IsNullOrWhiteSpace($WorkbookName)) {
        $lines += "currentWorkbook=$WorkbookName"
    }

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $lines += "message=$Message"
    }

    Set-Content -LiteralPath $script:StatusFile -Value $lines
}

function Resolve-OptionalBoolean {
    param(
        [object] $Value,
        [bool] $DefaultValue,
        [string] $Name
    )

    if ($null -eq $Value) {
        return $DefaultValue
    }

    if ($Value -is [bool]) {
        return [bool] $Value
    }

    $text = [string] $Value
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "System.Management.Automation.Internal.AutomationNull") {
        return $DefaultValue
    }

    switch -Regex ($text.Trim()) {
        '^(true|\$true|1|yes|y)$' { return $true }
        '^(false|\$false|0|no|n)$' { return $false }
        default { throw "$Name must be true or false." }
    }
}

function Resolve-OptionalInteger {
    param(
        [object] $Value,
        [int] $DefaultValue,
        [string] $Name
    )

    if ($null -eq $Value) {
        return $DefaultValue
    }

    $text = [string] $Value
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "System.Management.Automation.Internal.AutomationNull") {
        return $DefaultValue
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref] $parsed)) {
        return $parsed
    }

    throw "$Name must be an integer."
}

function Release-ComObject {
    param(
        [object] $ComObject,
        [string] $Name
    )

    if ($null -eq $ComObject) {
        return
    }

    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        Write-Log "Released COM object: $Name"
    }
    catch {
        Write-Log "Failed to release COM object $Name`: $($_.Exception.Message)" "WARN"
    }
}

function Test-ExcelTempLockExists {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $folder = Split-Path -Path $Path -Parent
    $fileName = Split-Path -Path $Path -Leaf
    $lockPath = Join-Path $folder ("~$" + $fileName)
    return (Test-Path -LiteralPath $lockPath -PathType Leaf)
}

function Test-WorkbookWritable {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
        $stream.Close()
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Close()
            $stream.Dispose()
        }
    }
}

function Get-StopRequestMode {
    if ([string]::IsNullOrWhiteSpace($StopRequestPath) -or -not (Test-Path -LiteralPath $StopRequestPath -PathType Leaf)) {
        return $null
    }

    foreach ($line in @(Get-Content -LiteralPath $StopRequestPath -ErrorAction SilentlyContinue)) {
        if ($line -match '^mode=(.+)$') {
            $mode = $matches[1].Trim()
            if ($mode -eq "Now" -or $mode -eq "AfterCurrent") {
                return $mode
            }
        }
    }

    return $null
}

function Test-StopNowRequested {
    return ((Get-StopRequestMode) -eq "Now")
}

function Assert-NotStopNowRequested {
    if (Test-StopNowRequested) {
        $script:StopRequested = $true
        throw "Operator stop-now request observed."
    }
}

function Start-ResponsiveSleep {
    param(
        [int] $Seconds
    )

    for ($i = 0; $i -lt $Seconds; $i++) {
        Assert-NotStopNowRequested
        Start-Sleep -Seconds 1
    }
}

function Test-ConnectionRefreshing {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Connection
    )

    foreach ($propertyPath in @("OLEDBConnection", "ODBCConnection")) {
        try {
            $connectionObject = $Connection.$propertyPath
            if ($null -ne $connectionObject -and $connectionObject.Refreshing) {
                return $true
            }
        }
        catch {
            # Some connection types do not expose these properties.
        }
    }

    return $false
}

function Wait-WorkbookRefresh {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Workbook,

        [int] $TimeoutMinutes
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $lastLog = Get-Date
    $inactivePollCount = 0
    $noPollableConnectionCount = 0
    $requiredInactivePollCount = 3
    $requiredNoPollableConnectionCount = 6
    $pollIntervalSeconds = 5
    $finalSettleSeconds = 15

    do {
        Assert-NotStopNowRequested

        $refreshingCount = 0
        $checkedCount = 0
        $hadPollingWarning = $false

        try {
            foreach ($connection in @($Workbook.Connections)) {
                if ($null -eq $connection) {
                    $hadPollingWarning = $true
                    Write-Log "Connection polling warning: workbook exposed a null connection entry." "WARN"
                    continue
                }

                $checkedCount++
                if (Test-ConnectionRefreshing -Connection $connection) {
                    $refreshingCount++
                }
            }
        }
        catch {
            $hadPollingWarning = $true
            Write-Log "Connection polling warning: $($_.Exception.Message)" "WARN"
        }

        if (((Get-Date) - $lastLog).TotalSeconds -ge 15) {
            Write-Log "Refresh poll: $refreshingCount active connection(s) out of $checkedCount checked. Inactive confirmation $inactivePollCount of $requiredInactivePollCount."
            $lastLog = Get-Date
        }

        if ($hadPollingWarning) {
            $inactivePollCount = 0
            $noPollableConnectionCount = 0
            Start-ResponsiveSleep -Seconds $pollIntervalSeconds
            continue
        }

        if ($checkedCount -gt 0 -and $refreshingCount -eq 0) {
            $inactivePollCount++
            $noPollableConnectionCount = 0

            if ($inactivePollCount -ge $requiredInactivePollCount) {
                Write-Log "No active workbook connections reported for $inactivePollCount consecutive poll(s); waiting final settle period."
                Start-ResponsiveSleep -Seconds $finalSettleSeconds
                return
            }

            Write-Log "No active workbook connections reported; waiting for confirmation poll $inactivePollCount of $requiredInactivePollCount."
            Start-ResponsiveSleep -Seconds $pollIntervalSeconds
            continue
        }

        if ($refreshingCount -gt 0) {
            $inactivePollCount = 0
            $noPollableConnectionCount = 0
            Start-ResponsiveSleep -Seconds $pollIntervalSeconds
            continue
        }

        if ($checkedCount -eq 0) {
            $inactivePollCount = 0
            $noPollableConnectionCount++

            if ($noPollableConnectionCount -ge $requiredNoPollableConnectionCount) {
                Write-Log "Workbook exposes no pollable connections after $noPollableConnectionCount consecutive poll(s); waiting final settle period." "WARN"
                Start-ResponsiveSleep -Seconds $finalSettleSeconds
                return
            }

            Write-Log "Workbook exposes no pollable connections; waiting for confirmation poll $noPollableConnectionCount of $requiredNoPollableConnectionCount." "WARN"
            Start-ResponsiveSleep -Seconds $pollIntervalSeconds
            continue
        }

        Start-ResponsiveSleep -Seconds $pollIntervalSeconds
    } while ((Get-Date) -lt $deadline)

    throw "Refresh exceeded timeout of $TimeoutMinutes minute(s) while polling workbook connections."
}

$excel = $null
$workbook = $null

try {
    $visible = if ($PSBoundParameters.ContainsKey("VisibleOverride")) {
        Resolve-OptionalBoolean -Value $VisibleOverride -DefaultValue $DefaultVisible -Name "VisibleOverride"
    }
    else {
        $DefaultVisible
    }

    $timeoutMinutes = if ($PSBoundParameters.ContainsKey("TimeoutMinutesOverride")) {
        Resolve-OptionalInteger -Value $TimeoutMinutesOverride -DefaultValue $DefaultTimeoutMinutes -Name "TimeoutMinutesOverride"
    }
    else {
        $DefaultTimeoutMinutes
    }

    $useSkipAsyncWait = if ($PSBoundParameters.ContainsKey("SkipAsyncWait")) {
        Resolve-OptionalBoolean -Value $SkipAsyncWait -DefaultValue $DefaultSkipAsyncWait -Name "SkipAsyncWait"
    }
    else {
        $DefaultSkipAsyncWait
    }

    $useCleanupGhostExcelProcesses = if ($PSBoundParameters.ContainsKey("CleanupGhostExcelProcessesOverride")) {
        Resolve-OptionalBoolean -Value $CleanupGhostExcelProcessesOverride -DefaultValue $DefaultCleanupGhostExcelProcesses -Name "CleanupGhostExcelProcessesOverride"
    }
    else {
        $DefaultCleanupGhostExcelProcesses
    }

    $useForceCloseExcelProcesses = if ($PSBoundParameters.ContainsKey("ForceCloseExcelProcessesOverride")) {
        Resolve-OptionalBoolean -Value $ForceCloseExcelProcessesOverride -DefaultValue $DefaultForceCloseExcelProcesses -Name "ForceCloseExcelProcessesOverride"
    }
    else {
        $DefaultForceCloseExcelProcesses
    }

    if ($timeoutMinutes -lt 1) {
        throw "TimeoutMinutes must be at least 1."
    }

    if ($useCleanupGhostExcelProcesses -or $useForceCloseExcelProcesses) {
        throw "ResidentialCare Date Refresh safe mode does not close or kill pre-existing Excel processes."
    }

    $script:LogFile = $LogPath
    $script:StatusFile = $StatusPath

    Write-CurrentStatus -State "RUNNING" -Message "Starting workbook refresh."
    Write-Log "Starting Excel refresh automation."
    Write-Log "Workbook path input: $WorkbookPath"
    Write-Log "Mode: $(if ($visible) { 'displayed' } else { 'hidden' }); timeout: $timeoutMinutes minute(s); async wait skipped: $useSkipAsyncWait."

    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
        throw "Workbook path is missing."
    }

    $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath -ErrorAction Stop).Path
    if (-not (Test-Path -LiteralPath $resolvedWorkbookPath -PathType Leaf)) {
        throw "Workbook path does not exist: $resolvedWorkbookPath"
    }

    Assert-NotStopNowRequested

    if (Test-ExcelTempLockExists -Path $resolvedWorkbookPath) {
        throw "Workbook appears open in Excel because a temporary lock file exists: $resolvedWorkbookPath. ResidentialCare Date Refresh safe mode will not close Excel processes."
    }

    if (-not (Test-WorkbookWritable -Path $resolvedWorkbookPath)) {
        throw "Workbook is not writable for refresh: $resolvedWorkbookPath. ResidentialCare Date Refresh safe mode will not close Excel processes."
    }

    Assert-NotStopNowRequested

    Write-Log "Starting a clean Excel instance."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $visible
    $excel.DisplayAlerts = $false

    Write-Log "Opening workbook."
    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath)
    Write-Log "Workbook opened. Waiting $DefaultOpenSettleSeconds second(s) for Excel to settle."
    Start-ResponsiveSleep -Seconds $DefaultOpenSettleSeconds

    Assert-NotStopNowRequested

    $refreshStart = Get-Date
    Write-Log "Starting workbook refresh."
    $workbook.RefreshAll()
    Write-Log "Refresh command accepted by Excel."

    if ($useSkipAsyncWait) {
        Write-Log "Monitoring workbook connections for refresh activity."
        Wait-WorkbookRefresh -Workbook $workbook -TimeoutMinutes $timeoutMinutes
    }
    else {
        Write-Log "Waiting for asynchronous Excel queries to complete."
        $excel.CalculateUntilAsyncQueriesDone()
    }

    Assert-NotStopNowRequested

    $elapsedMinutes = ((Get-Date) - $refreshStart).TotalMinutes
    if ($elapsedMinutes -gt $timeoutMinutes) {
        throw "Refresh exceeded timeout of $timeoutMinutes minute(s)."
    }

    Write-Log ("Refresh monitoring complete after {0:n2} minute(s)." -f $elapsedMinutes)
    Write-Log "Saving workbook."
    $workbook.Save()
    Write-Log "Workbook saved."

    Write-Log "Closing workbook."
    $workbook.Close($false)
    $script:WorkbookClosed = $true
    Write-Log "Workbook closed."

    Write-Log "Quitting Excel."
    $excel.Quit()
    $script:ExcelQuit = $true
    Write-Log "Excel quit command completed."

    $script:FinalStatus = "Success"
    $script:ExitCode = 0
}
catch {
    $script:PrimaryError = $_.Exception.Message
    if ($script:StopRequested) {
        $script:FinalStatus = "Stopped"
        $script:ExitCode = 3
        Write-Log "Operator stop-now request observed. The active workbook will close without saving." "WARN"
        Write-CurrentStatus -State "STOPPING" -Message "Operator stop-now request observed."
    }
    else {
        $script:FinalStatus = "Error"
        $script:ExitCode = 1
        Write-Log "Refresh run failed: $script:PrimaryError" "ERROR"
        Write-CurrentStatus -State "FAILED" -Message $script:PrimaryError
    }
}
finally {
    if ($null -ne $workbook -and -not $script:WorkbookClosed) {
        try {
            Write-Log "Closing workbook without saving during cleanup."
            $workbook.Close($false)
            $script:WorkbookClosed = $true
        }
        catch {
            Write-Log "Workbook close failed during cleanup: $($_.Exception.Message)" "WARN"
        }
    }

    if ($null -ne $excel -and -not $script:ExcelQuit) {
        try {
            Write-Log "Quitting Excel during cleanup."
            $excel.Quit()
            $script:ExcelQuit = $true
        }
        catch {
            Write-Log "Excel quit failed during cleanup: $($_.Exception.Message)" "WARN"
        }
    }

    Release-ComObject -ComObject $workbook -Name "Workbook"
    Release-ComObject -ComObject $excel -Name "Excel.Application"

    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    catch {
        Write-Log "Garbage collection cleanup failed: $($_.Exception.Message)" "WARN"
    }

    if ($script:FinalStatus -eq "Stopped") {
        Write-CurrentStatus -State "STOPPED" -Message "Operator stop-now request observed."
    }
    elseif ($script:FinalStatus -eq "Success") {
        Write-Log "Refresh run succeeded."
    }

    exit $script:ExitCode
}
