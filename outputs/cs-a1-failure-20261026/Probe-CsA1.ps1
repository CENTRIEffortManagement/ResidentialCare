[CmdletBinding()]
param(
    [switch]$Worker,
    [switch]$RefreshAll,
    [switch]$Writable,
    [switch]$VisibleTest,
    [switch]$SkipPrecalculation,
    [switch]$Ribbon,
    [switch]$SuppressAlerts,
    [switch]$ForceEvents,
    [switch]$ForegroundAll,
    [ValidateSet('CS', 'AIN')]
    [string]$RoleTarget = 'CS',
    [ValidateSet('Original', 'LoadedTables')]
    [string]$ConnectionScope = 'Original',
    [string]$TableSelection = '',
    [ValidateRange(0, 90)]
    [int]$NoComObservationSeconds = 0,
    [ValidateSet('Saved', 'ImportInside', 'PathOutside')]
    [string]$StartView = 'Saved',
    [switch]$EmptyPreviouslyBlankOutputs,
    [int]$QueryTimeoutSeconds = 120,
    [ValidatePattern('^[0-9]{8}-[0-9]{6}$')]
    [string]$RunTag = (Get-Date -Format 'yyyyMMdd-HHmmss')
)
$ErrorActionPreference = 'Stop'
$target = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot ("../../CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/" + $RoleTarget + "/CapacityDistrib(A.1)-shifts.xlsx"))).Path
$statePath = Join-Path $PSScriptRoot ("native-probe-$RunTag.json")
$logPath = Join-Path $PSScriptRoot ("native-probe-$RunTag.log")
function Read-SavedHash {
    $stream = [IO.FileStream]::new($target, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $memory = [IO.MemoryStream]::new()
    try { $stream.CopyTo($memory); return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($memory.ToArray())) }
    finally { $stream.Dispose(); $memory.Dispose() }
}

if (-not $Worker) {
    $workerArgs = @('-NoProfile', '-NonInteractive', '-STA', '-File', ('"' + $PSCommandPath + '"'), '-Worker', '-RunTag', $RunTag)
    if ($RefreshAll) { $workerArgs += '-RefreshAll' }
    if ($Writable) { $workerArgs += '-Writable' }
    if ($VisibleTest) { $workerArgs += '-VisibleTest' }
    if ($SkipPrecalculation) { $workerArgs += '-SkipPrecalculation' }
    if ($Ribbon) { $workerArgs += '-Ribbon' }
    if ($SuppressAlerts) { $workerArgs += '-SuppressAlerts' }
    if ($ForceEvents) { $workerArgs += '-ForceEvents' }
    if ($ForegroundAll) { $workerArgs += '-ForegroundAll' }
    $workerArgs += @('-RoleTarget', $RoleTarget, '-ConnectionScope', $ConnectionScope, '-NoComObservationSeconds', $NoComObservationSeconds)
    $workerArgs += @('-StartView', $StartView)
    if ($EmptyPreviouslyBlankOutputs) { $workerArgs += '-EmptyPreviouslyBlankOutputs' }
    if ($TableSelection) { $workerArgs += @('-TableSelection', ('"' + $TableSelection + '"')) }
    $workerProcess = Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList $workerArgs -WindowStyle Hidden -PassThru
    $lastPhase = ''
    $lastChange = [datetime]::UtcNow
    $timedOut = $false
    $probeState = $null
    $capturedOwnedWindow = $false
    $probeStartedUtc = [datetime]::UtcNow
    $lastOwnedSnapshotUtc = $probeStartedUtc
    while (-not $workerProcess.HasExited) {
        Start-Sleep -Seconds 2
        if (Test-Path -LiteralPath $statePath) {
            try { $probeState = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json } catch {}
            if ($probeState.Sequence -ne $lastPhase) {
                $lastPhase = $probeState.Sequence
                $lastChange = [datetime]::UtcNow
                Write-Output ($probeState.Phase + ': ' + $probeState.Table)
            }
        }
        if (([datetime]::UtcNow - $lastChange).TotalSeconds -gt $QueryTimeoutSeconds) {
            $timedOut = $true
            Write-Output ('Timed out at phase ' + $lastPhase)
            # Only this probe's separately created Excel process may be terminated.
            if ($probeState -and $probeState.OwnedExcelId) {
                $owned = Get-Process -Id $probeState.OwnedExcelId -ErrorAction SilentlyContinue
                if ($owned -and $owned.ProcessName -eq 'EXCEL' -and
                    $owned.StartTime.ToUniversalTime().Ticks -eq [long]$probeState.OwnedExcelStartTicks) {
                    Stop-Process -Id $owned.Id -Force
                }
            }
            if (-not $workerProcess.WaitForExit(10000)) {
                Stop-Process -Id $workerProcess.Id -Force
            }
            break
        }
        if ($probeState -and $probeState.OwnedExcelId -and
            ([datetime]::UtcNow - $lastOwnedSnapshotUtc).TotalSeconds -ge 15) {
            $ownedSnapshot = Get-Process -Id $probeState.OwnedExcelId -ErrorAction SilentlyContinue
            if ($ownedSnapshot -and $ownedSnapshot.ProcessName -eq 'EXCEL' -and
                $ownedSnapshot.StartTime.ToUniversalTime().Ticks -eq [long]$probeState.OwnedExcelStartTicks) {
                $ownedSnapshot.Refresh()
                [IO.File]::AppendAllText($logPath, (([ordered]@{TimeUtc=[datetime]::UtcNow.ToString('o'); OwnedWindowSnapshot=$true; ProcessId=$ownedSnapshot.Id; Responding=$ownedSnapshot.Responding; WindowHandle=$ownedSnapshot.MainWindowHandle.ToInt64(); WindowTitle=$ownedSnapshot.MainWindowTitle; CpuSeconds=$ownedSnapshot.CPU}) | ConvertTo-Json -Compress) + [Environment]::NewLine)
                # Window/thread inspection only, restricted to the separately owned Excel PID.
                if (-not $capturedOwnedWindow) {
                    & (Join-Path $PSScriptRoot 'Inspect-LiveWait.ps1') -SkipWaitChain -TargetProcessIds $ownedSnapshot.Id -KnownUiWindow 0 -KnownUiProcessId $ownedSnapshot.Id
                }
            }
            $capturedOwnedWindow = $true
            $lastOwnedSnapshotUtc = [datetime]::UtcNow
        }
        $workerProcess.Refresh()
    }
    if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath }
    Write-Output ('Workbook SHA256 after probe: ' + (Read-SavedHash))
    if ($timedOut) { exit 2 }
    exit $workerProcess.ExitCode
}

$sequence = 0
$ownedId = $null
$ownedTicks = $null
function Write-ProbeState([string]$Phase, [string]$Table = '', $Data = $null) {
    $script:sequence++
    $entry = [ordered]@{
        Sequence = $script:sequence; Phase = $Phase; Table = $Table
        TimeUtc = [datetime]::UtcNow.ToString('o')
        OwnedExcelId = $script:ownedId; OwnedExcelStartTicks = $script:ownedTicks
        Data = $Data
    }
    $json = $entry | ConvertTo-Json -Depth 5 -Compress
    [IO.File]::WriteAllText($statePath + '.tmp', $json)
    [IO.File]::Move($statePath + '.tmp', $statePath, $true)
    [IO.File]::AppendAllText($logPath, $json + [Environment]::NewLine)
}
$beforeHash = Read-SavedHash
$existingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = $null
$book = $null
$owned = $false
$exitCode = 0
try {
    Write-ProbeState 'Creating isolated Excel' '' @{Workbook = $target; SHA256 = $beforeHash}
    $excel = New-Object -ComObject Excel.Application
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CsA1ProbeWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    [uint32]$excelId = 0
    [void][CsA1ProbeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelId)
    if ($existingIds -contains [int]$excelId) { throw 'COM returned an existing Excel process; probe aborted without modifying it.' }
    $ownedProcess = Get-Process -Id $excelId
    $script:ownedId = [int]$excelId
    $script:ownedTicks = $ownedProcess.StartTime.ToUniversalTime().Ticks
    $owned = $true
    $excel.Visible = [bool]$VisibleTest
    $excel.DisplayAlerts = [bool]($VisibleTest -and -not $SuppressAlerts)
    $excel.EnableEvents = [bool]($Writable -or $VisibleTest -or $ForceEvents)
    $excel.AutomationSecurity = 3
    $openReadOnly = -not [bool]$Writable
    if ($Writable) {
        $ownerFile = Join-Path (Split-Path -Parent $target) '~$CapacityDistrib(A.1)-shifts.xlsx'
        if (Test-Path -LiteralPath $ownerFile) { throw 'The exact CS workbook has an Excel owner file; editable probe aborted.' }
    }
    Write-ProbeState 'Opening exact workbook without saving' '' @{ReadOnly = $openReadOnly; EventsEnabled = $excel.EnableEvents}
    $book = $excel.Workbooks.Open($target, 0, $openReadOnly)
    if ($book.ReadOnly -ne $openReadOnly -or $book.FullName -ne $target) { throw 'Exact workbook identity and requested open mode were not confirmed.' }
    if ($Writable -and $book.AutoSaveOn) { $book.AutoSaveOn = $false }
    if (-not $RefreshAll) { $excel.Calculation = -4135 }
    if ($VisibleTest) { $excel.Caption = $RoleTarget + ' A.1 DIAGNOSTIC - closes without saving' }
    if (-not $SkipPrecalculation) {
        Write-ProbeState 'Calculating worksheets before refresh'
        foreach ($sheet in $book.Worksheets) { $sheet.Calculate() }
    }
    Write-ProbeState 'Refresh environment' '' @{Visible=$excel.Visible; DisplayAlerts=$excel.DisplayAlerts; Calculation=$excel.Calculation; ExplicitPrecalculation=(-not $SkipPrecalculation); Ribbon=[bool]$Ribbon}
    $tables = @{}
    foreach ($sheet in $book.Worksheets) {
        foreach ($table in $sheet.ListObjects) { $tables[$table.Name] = $table }
    }
    $originalConnectionSettings = @()
    foreach ($connection in $book.Connections) {
        $settings = @{Name=$connection.Name; Type=$connection.Type; RefreshWithRefreshAll=$connection.RefreshWithRefreshAll}
        if ($connection.Type -eq 1) { $settings.BackgroundQuery = $connection.OLEDBConnection.BackgroundQuery }
        $originalConnectionSettings += $settings
    }
    $loadedTableSettings = @()
    foreach ($table in $tables.Values) {
        $loadedTableSettings += @{Name=$table.Name; Query=$table.QueryTable.CommandText; Connection=$table.QueryTable.WorkbookConnection.Name; BackgroundQuery=$table.QueryTable.BackgroundQuery}
    }
    Write-ProbeState 'Original connection settings' '' @{Role=$RoleTarget; Connections=$originalConnectionSettings; LoadedTables=$loadedTableSettings; ConnectionScope=$ConnectionScope}
    if ($EmptyPreviouslyBlankOutputs) {
        if ($RoleTarget -ne 'CS') { throw 'Empty-output simulation is restricted to exact BD CS A.1.' }
        $formerlyEmpty = @('A1SpareDayChoices_Prepare','ResPeriodShiftNWDTABLE','ResPeriodAvailabilityTABLE','ResPeriodAllocationTABLE','ResPeriodCapPrioritised','ResPeriodWDTABLE','ResourcePeriodTABLE_empty','MultiDayPeriod_Remove','C__ResSingleShiftDayTABLE','Demand_Prepare')
        $emptyOutputChanges = @()
        foreach ($name in $formerlyEmpty) {
            if (-not $tables.ContainsKey($name)) { throw ('Expected CS output missing: ' + $name) }
            $table = $tables[$name]
            $previousRows = $table.ListRows.Count
            [void]$table.DataBodyRange.ClearContents()
            $placeholderRange = $table.HeaderRowRange.Resize(2, $table.ListColumns.Count)
            [void]$table.Resize($placeholderRange)
            $emptyOutputChanges += @{Name=$name; PreviousRows=$previousRows; Rows=$table.ListRows.Count}
        }
        Write-ProbeState 'Simulated former empty output ranges in memory only' '' @{Tables=$emptyOutputChanges; Saved=$false}
    }
    if ($StartView -eq 'ImportInside') {
        $book.Worksheets.Item('IMPORT ResDayShift !!').Activate()
        $excel.ActiveSheet.Range('A1').Select()
    } elseif ($StartView -eq 'PathOutside') {
        $book.Worksheets.Item('UnitL1Path').Activate()
        $excel.ActiveSheet.Range('D1').Select()
    }
    Write-ProbeState 'Refresh start view' '' @{Mode=$StartView; Sheet=$excel.ActiveSheet.Name; Selection=$excel.Selection.Address(); FilePathFormula=$book.Worksheets.Item('UnitL1Path').Range('B1').Formula}
    if ($TableSelection -and $ConnectionScope -ne 'LoadedTables') { throw 'TableSelection requires LoadedTables scope.' }
    if ($ConnectionScope -eq 'LoadedTables') {
        $selectedRootTables = @($tables.Keys)
        if ($TableSelection) { $selectedRootTables = @($TableSelection.Split(',') | ForEach-Object { $_.Trim() }) }
        $selectedRootConnections = @()
        foreach ($name in $selectedRootTables) {
            if (-not $tables.ContainsKey($name)) { throw ('Selected table is missing: ' + $name) }
            $selectedRootConnections += $tables[$name].QueryTable.WorkbookConnection.Name
        }
        foreach ($connection in $book.Connections) {
            $connection.RefreshWithRefreshAll = [bool]($selectedRootConnections -contains $connection.Name)
        }
        Write-ProbeState 'Limited root refresh in diagnostic memory only' '' @{Tables=$selectedRootTables; Connections=$selectedRootConnections; BackgroundSettingsUnchanged=$true}
    }
    if ($ForegroundAll) {
        Write-ProbeState 'Disabling background refresh in diagnostic memory only'
        foreach ($connection in $book.Connections) {
            if ($connection.Type -eq 1) { $connection.OLEDBConnection.BackgroundQuery = $false }
        }
        foreach ($table in $tables.Values) { $table.QueryTable.BackgroundQuery = $false }
        Write-ProbeState 'Foreground refresh configured' '' @{Connections=$book.Connections.Count; Tables=$tables.Count; Calculation=$excel.Calculation}
    }
    $selected = @(
        'RolePathTABLE', 'ResPeriodAllocationTABLE', 'A1SpareDayChoices_Prepare',
        'ResPeriodAvailabilityTABLE', 'ResourcePeriodTABLE_empty',
        'ResPeriodCapPrioritised', 'ResPeriodWDTABLE', 'ResPeriodShiftNWDTABLE',
        'MultiDayPeriod_Remove', 'C__ResSingleShiftDayTABLE',
        'CapacityDiagnostics_SUMMARY', 'CapacityDiagnostics_DETAILS', 'Demand_Prepare', 'IMPORT_ResDayShift'
    )
    if ($RefreshAll) {
        $allTimer = [Diagnostics.Stopwatch]::StartNew()
        Write-ProbeState 'Starting Refresh All'
        if ($Ribbon) {
            $book.Activate()
            Write-ProbeState 'Executing ribbon command' '' @{Label=$excel.CommandBars.GetLabelMso('RefreshAll'); Enabled=$excel.CommandBars.GetEnabledMso('RefreshAll')}
            $excel.CommandBars.ExecuteMso('RefreshAll')
        } else { $book.RefreshAll() }
        Write-ProbeState 'Waiting for Refresh All'
        if ($NoComObservationSeconds -gt 0) {
            Write-ProbeState 'No COM observation interval' '' @{Seconds=$NoComObservationSeconds}
            $noComTimer = [Diagnostics.Stopwatch]::StartNew()
            while ($noComTimer.Elapsed.TotalSeconds -lt $NoComObservationSeconds) { Start-Sleep -Seconds 1 }
            Write-ProbeState 'No COM observation interval ended'
        }
        $idleSamples = 0
        do {
            $busy = @()
            Write-ProbeState 'Reading connection refresh states'
            foreach ($connection in $book.Connections) {
                if ($connection.Type -eq 1 -and $connection.OLEDBConnection.Refreshing) { $busy += $connection.Name }
            }
            Write-ProbeState 'Reading table refresh states'
            foreach ($table in $tables.Values) {
                if ($table.QueryTable.Refreshing) { $busy += $table.Name }
            }
            Write-ProbeState 'Reading calculation state'
            $calculation = [int]$excel.CalculationState
            $snapshot = @{TimeUtc = [datetime]::UtcNow.ToString('o'); Active = $busy; CalculationState = $calculation; Seconds = [math]::Round($allTimer.Elapsed.TotalSeconds, 2)}
            [IO.File]::AppendAllText($logPath, ($snapshot | ConvertTo-Json -Depth 4 -Compress) + [Environment]::NewLine)
            if ($busy.Count -eq 0 -and $calculation -eq 0) { $idleSamples++ } else { $idleSamples = 0 }
            if ($idleSamples -lt 2) { Start-Sleep -Seconds 3 }
        } while ($idleSamples -lt 2)
        $allTimer.Stop()
        $counts = @{}
        foreach ($name in $selected) { $counts[$name] = $tables[$name].ListRows.Count }
        Write-ProbeState 'Refresh All completed' '' @{Seconds = [math]::Round($allTimer.Elapsed.TotalSeconds, 3); Rows = $counts}
    }
    else { foreach ($name in $selected) {
        if (-not $tables.ContainsKey($name)) { throw "Expected loaded table missing: $name" }
        $queryTable = $tables[$name].QueryTable
        $queryTable.BackgroundQuery = $false
        Write-ProbeState 'Refreshing selected table' $name
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $ok = $queryTable.Refresh($false)
        $timer.Stop()
        if ($ok -isnot [bool] -or -not $ok) { throw "Refresh did not return Boolean true: $name" }
        Write-ProbeState 'Reading refreshed row count' $name
        Write-ProbeState 'Refreshed selected table' $name @{
            Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 3)
            Rows = $tables[$name].ListRows.Count
        }
    } }
    $summaryFailures = @()
    $summaryValues = $tables['CapacityDiagnostics_SUMMARY'].DataBodyRange.Value2
    for ($row = 1; $row -le $summaryValues.GetLength(0); $row++) {
        if ($summaryValues[$row, 3] -ne 'Pass') {
            $summaryFailures += @{Check = $summaryValues[$row, 2]; Status = $summaryValues[$row, 3]; Failures = $summaryValues[$row, 4]; Details = $summaryValues[$row, 5]}
        }
    }
    Write-ProbeState 'Native probe completed' '' @{SummaryNonPass = $summaryFailures}
}
catch {
    $exitCode = 1
    Write-ProbeState 'Error' '' @{Message = $_.Exception.Message; Details = ($_ | Out-String)}
}
finally {
    Write-ProbeState 'Closing workbook without saving'
    if ($book) { try { $book.Close($false) } catch {} }
    if ($owned -and $excel) { try { $excel.Quit() } catch {} }
    if ($book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $afterHash = Read-SavedHash
    if ($beforeHash -ne $afterHash) { $exitCode = 1 }
    Write-ProbeState 'Closed without saving' '' @{BeforeSHA256 = $beforeHash; AfterSHA256 = $afterHash; Unchanged = $beforeHash -eq $afterHash}
}
exit $exitCode
