# Shared by the standalone and integrated all-units refresh engines.
# The supervisor never calls Excel COM: deadlines remain enforceable during Save/Open/Quit.
. (Join-Path $PSScriptRoot 'RefreshJson.ps1')
. (Join-Path $PSScriptRoot 'RefreshGitAccess.ps1')
function Get-RefreshFileUsers {
    param([string] $Path)
    if (-not ('ResidentialCare.FileUsers' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace ResidentialCare {
 public static class FileUsers {
  [StructLayout(LayoutKind.Sequential)] public struct ProcessKey { public int Id; public System.Runtime.InteropServices.ComTypes.FILETIME Started; }
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] public struct ProcessInfo {
   public ProcessKey Process;
   [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string AppName;
   [MarshalAs(UnmanagedType.ByValTStr,SizeConst=64)] public string Service;
   public int AppType; public uint Status; public uint Session; [MarshalAs(UnmanagedType.Bool)] public bool Restartable;
  }
  [DllImport("rstrtmgr.dll",CharSet=CharSet.Unicode)] static extern int RmStartSession(out uint handle,int flags,StringBuilder key);
  [DllImport("rstrtmgr.dll",CharSet=CharSet.Unicode)] static extern int RmRegisterResources(uint handle,uint count,string[] files,uint appCount,ProcessKey[] apps,uint serviceCount,string[] services);
  [DllImport("rstrtmgr.dll")] static extern int RmGetList(uint handle,out uint needed,ref uint count,[In,Out] ProcessInfo[] processes,ref uint reason);
  [DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint handle);
  public static ProcessInfo[] Query(string file) {
   uint handle; int result=RmStartSession(out handle,0,new StringBuilder(33));
   if(result!=0) throw new Exception("Cannot inspect file users: RmStartSession error "+result);
   try {
    result=RmRegisterResources(handle,1,new[]{file},0,null,0,null);
    if(result!=0) throw new Exception("Cannot inspect file users: RmRegisterResources error "+result);
    uint needed=0,count=0,reason=0; result=RmGetList(handle,out needed,ref count,null,ref reason);
    if(result==0) return new ProcessInfo[0];
    for(int attempt=0;result==234 && attempt<3;attempt++) {
     count=needed; var owners=new ProcessInfo[count];
     result=RmGetList(handle,out needed,ref count,owners,ref reason);
     if(result==0) { Array.Resize(ref owners,(int)count); return owners; }
    }
    throw new Exception("Cannot inspect file users: RmGetList error "+result);
   } finally { RmEndSession(handle); }
  }
 }
}
'@
    }
    # Restart Manager is used only to list resource users; never shut down their applications.
    return [ResidentialCare.FileUsers]::Query([IO.Path]::GetFullPath($Path))
}

function Assert-NoExternalFileUsers {
    param([string] $Path, [int] $AllowedProcessId = 0,
        [ValidateRange(0, 600)] [double] $WaitSeconds = 0,
        [ValidateRange(1, 1000)] [int] $PollMilliseconds = 500,
        [scriptblock] $CheckStop = {}, [scriptblock] $OnWait = {},
        [scriptblock] $OnDecision = { param($decision) Write-RefreshAccessDecision $decision })
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $lastNotice = -5.0
    $notices = @{}
    while ($true) {
        & $CheckStop
        # Resource use is evidence, not proof of a blocking OS lock. Only
        # verified read-only Git is exempted; unidentified processes still wait.
        $blockers = @(foreach ($user in @(Get-RefreshFileUsers $Path)) {
            if ($user.Process.Id -eq $AllowedProcessId) { continue }
            $decision = Get-RefreshFileUserDecision $user
            $decision | Add-Member NoteProperty Target $Path
            $key = $decision | ConvertTo-Json -Compress
            if (-not $notices.ContainsKey($key)) { & $OnDecision $decision; $notices[$key] = $true }
            if (-not $decision.Allowed) { $user }
        })
        if ($blockers.Count -eq 0) { return }
        $description = ($blockers | ForEach-Object { "$($_.AppName) (PID $($_.Process.Id))" }) -join ', '
        if ($timer.Elapsed.TotalSeconds -ge $WaitSeconds) {
            throw "External-user policy blocked access: $description after waiting $WaitSeconds second(s). No save attempted by this check; this is not an Excel save error. Close/disconnect those applications before retrying: $Path"
        }
        if (($timer.Elapsed.TotalSeconds - $lastNotice) -ge 5) {
            & $OnWait $description
            $lastNotice = $timer.Elapsed.TotalSeconds
        }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
}

function Get-RefreshTargetStamp {
    param([string] $Path)
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    return "$($item.Length):$($item.LastWriteTimeUtc.Ticks)"
}

function Wait-RefreshFileAccess {
    param([string] $Path, [int] $AllowedProcessId = 0,
        [ValidateRange(0, 600)] [double] $WaitSeconds = 60)
    $originalStamp = Get-RefreshTargetStamp $Path
    Assert-NoExternalFileUsers -Path $Path -AllowedProcessId $AllowedProcessId -WaitSeconds $WaitSeconds `
        -CheckStop { Assert-NotStopNowRequested } -OnWait {
            param($description)
            $script:RefreshState.detail = "Waiting for file access: $description"
            Write-RefreshWorkerState
            Write-CurrentStatus -State 'RUNNING' -Message $script:RefreshState.detail
            Write-Log "$($script:RefreshState.detail). Will recheck; no save has been attempted." 'WARN'
        }
    Assert-NotStopNowRequested
    if ((Get-RefreshTargetStamp $Path) -ne $originalStamp) {
        throw "Workbook changed while waiting for external file users; refusing to overwrite it: $Path"
    }
    $script:RefreshState.detail = $null
    Write-RefreshWorkerState
    Write-Log 'File access policy passed; any permitted Git readers were recorded. OS access and save checks remain enforced.'
}

function New-RefreshFileSnapshot {
    param([string] $Path, [string] $InputSnapshotPath = '')
    $snapshot = @{ Target = [IO.Path]::GetFullPath($Path); TargetStamp = (Get-RefreshTargetStamp $Path); Inputs = @{} }
    if ($InputSnapshotPath) {
        $source = Read-RefreshWorkerState $InputSnapshotPath
        if ($source.SchemaVersion -ne 1 -or -not $source.Target.Equals($snapshot.Target, [StringComparison]::OrdinalIgnoreCase)) { throw 'Input snapshot does not belong to this workbook.' }
        foreach ($entry in $source.Inputs.PSObject.Properties) {
            if (-not [IO.Path]::IsPathRooted($entry.Name)) { throw 'Input snapshot paths must be fully resolved.' }
            $snapshot.Inputs[$entry.Name] = [string] $entry.Value
        }
    }
    Assert-RefreshFileSnapshot $snapshot
    return $snapshot
}

function Assert-RefreshFileSnapshot {
    param([hashtable] $Snapshot, [switch] $AfterSave)
    if (-not $AfterSave -and (Get-RefreshTargetStamp $Snapshot.Target) -ne $Snapshot.TargetStamp) {
        throw "Target workbook changed since pre-open capture; refusing to overwrite it: $($Snapshot.Target)"
    }
    foreach ($path in $Snapshot.Inputs.Keys) {
        if ((Get-RefreshTargetStamp $path) -ne $Snapshot.Inputs[$path]) {
            throw "Declared input changed since dispatch: $path. The refresh cannot be accepted."
        }
    }
}

function Wait-RefreshWorkbookIdentity {
    param($Workbook, [string] $ExpectedPath,
        [ValidateRange(0, 120)] [double] $WaitSeconds = 15,
        [ValidateRange(1, 1000)] [int] $PollMilliseconds = 250)
    $expected = [IO.Path]::GetFullPath($ExpectedPath)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $lastNotice = -5.0
    while ($true) {
        Assert-NotStopNowRequested
        $name = $null; $readOnly = $null; $detail = 'Excel has not returned a usable workbook object/path.'
        try {
            if ($null -ne $Workbook) {
                $name = [string] $Workbook.FullName
                $readOnly = $Workbook.ReadOnly
            }
        } catch { $detail = "Workbook identity is temporarily unavailable: $($_.Exception.Message)" }
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            # A nonblank wrong identity is a hard failure, not a reason to retry,
            # substitute another workbook, rename it or use Save As.
            if (-not [IO.Path]::IsPathRooted($name) -or
                -not ([IO.Path]::GetFullPath($name).Equals($expected, [StringComparison]::OrdinalIgnoreCase))) {
                throw "Excel workbook identity mismatch. Expected '$expected'; returned '$name'. No refresh or save is permitted."
            }
            if ($null -ne $readOnly) {
                if ([bool] $readOnly) { throw 'Excel opened the selected workbook read-only. Refusing refresh or Save As.' }
                return
            }
        }
        if ($timer.Elapsed.TotalSeconds -ge $WaitSeconds) {
            throw "Excel did not confirm the requested workbook path within $WaitSeconds second(s): $expected. $detail No refresh or save was permitted."
        }
        if (($timer.Elapsed.TotalSeconds - $lastNotice) -ge 5) {
            Write-Log "Waiting for Excel to confirm the opened workbook path: $expected. $detail" 'WARN'
            $lastNotice = $timer.Elapsed.TotalSeconds
        }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
}

function Release-ExcelComReference {
    param([object] $ComObject)
    if ($null -eq $ComObject -or -not [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) { return }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
}

function Disable-WorkbookBackgroundRefresh {
    param($Workbook)
    $settings = [Collections.Generic.List[object]]::new()
    $connectionSettingCount = 0
    $queryTableSettingCount = 0
    $connections = $Workbook.Connections
    try {
        for ($connectionIndex = 1; $connectionIndex -le [int] $connections.Count; $connectionIndex++) {
            $connection = if ([Runtime.InteropServices.Marshal]::IsComObject($connections)) { $connections.Item($connectionIndex) } else { $connections[$connectionIndex - 1] }
            $target = $null
            try {
                if ($null -eq $connection) { throw 'Excel returned a null connection while preparing synchronous refresh.' }
                $connectionName = [string] $connection.Name
                $target = switch ([int] $connection.Type) {
                    1 { $connection.OLEDBConnection; break }
                    2 { $connection.ODBCConnection; break }
                    default { $null }
                }
                if ($null -eq $target) { continue }
                $settings.Add([pscustomobject] @{
                    Target = $target
                    Label = "connection '$connectionName'"
                    Original = [bool] $target.BackgroundQuery
                    Changed = $false
                })
                $connectionSettingCount++
                $target = $null # The settings list owns this COM reference until restore.
            }
            catch { Write-Log "Background refresh setting unavailable for connection '$connectionName': $($_.Exception.Message)" 'WARN' }
            finally {
                Release-ExcelComReference $target
                Release-ExcelComReference $connection
            }
        }
    }
    finally { Release-ExcelComReference $connections }

    $worksheets = $Workbook.Worksheets
    try {
        for ($sheetIndex = 1; $sheetIndex -le [int] $worksheets.Count; $sheetIndex++) {
            $sheet = if ([Runtime.InteropServices.Marshal]::IsComObject($worksheets)) { $worksheets.Item($sheetIndex) } else { $worksheets[$sheetIndex - 1] }
            $queryTables = $null
            try {
                if ($null -eq $sheet) { throw 'Excel returned a null worksheet while preparing synchronous refresh.' }
                $sheetName = [string] $sheet.Name
                $queryTables = $sheet.QueryTables
                for ($queryIndex = 1; $queryIndex -le [int] $queryTables.Count; $queryIndex++) {
                    $queryTable = if ([Runtime.InteropServices.Marshal]::IsComObject($queryTables)) { $queryTables.Item($queryIndex) } else { $queryTables[$queryIndex - 1] }
                    try {
                        if ($null -eq $queryTable) { throw 'Excel returned a null query table while preparing synchronous refresh.' }
                        $queryTableName = [string] $queryTable.Name
                        $settings.Add([pscustomobject] @{
                            Target = $queryTable
                            Label = "query table '$sheetName/$queryTableName'"
                            Original = [bool] $queryTable.BackgroundQuery
                            Changed = $false
                        })
                        $queryTableSettingCount++
                        $queryTable = $null # The settings list owns this COM reference until restore.
                    }
                    catch { Write-Log "Background refresh setting unavailable for query table '$sheetName/$queryTableName': $($_.Exception.Message)" 'WARN' }
                    finally { Release-ExcelComReference $queryTable }
                }
            }
            finally {
                Release-ExcelComReference $queryTables
                Release-ExcelComReference $sheet
            }
        }
    }
    finally { Release-ExcelComReference $worksheets }

    $changedCount = 0
    foreach ($setting in $settings) {
        if (-not $setting.Original) { continue }
        try {
            $setting.Target.BackgroundQuery = $false
            $setting.Changed = $true
            $changedCount++
        }
        catch { Write-Log "Could not force synchronous refresh for $($setting.Label): $($_.Exception.Message)" 'WARN' }
    }
    Write-Log "Temporarily disabled background refresh for $changedCount of $($settings.Count) supported refresh object(s): $connectionSettingCount connection(s), $queryTableSettingCount worksheet query table(s)."
    return @($settings)
}

function Restore-WorkbookBackgroundRefresh {
    param([object[]] $Settings)
    $restoreError = $null
    foreach ($setting in @($Settings)) {
        try {
            if ($setting.Changed) {
                $setting.Target.BackgroundQuery = [bool] $setting.Original
                $setting.Changed = $false
            }
        }
        catch {
            if ($null -eq $restoreError) { $restoreError = $_ }
        }
        finally {
            Release-ExcelComReference $setting.Target
            $setting.Target = $null
        }
    }
    if ($null -ne $restoreError) { throw $restoreError }
    Write-Log 'Restored workbook background-refresh settings.'
}

function Write-RefreshWorkerState {
    $script:RefreshState.updatedAt = (Get-Date).ToString('o')
    Write-RefreshJson $WorkerStatePath $script:RefreshState -Depth 5
}

function Set-RefreshPhase {
    param([string] $Phase)
    $script:RefreshState.phase = $Phase
    $script:RefreshState.detail = $null
    Write-RefreshWorkerState
    Write-CurrentStatus -State 'RUNNING' -Message $Phase
    Write-Log "Phase: $Phase"
}

function Get-RefreshProcessIdentity {
    param([Diagnostics.Process] $Process)
    return @{ id = $Process.Id; startTicks = $Process.StartTime.ToUniversalTime().Ticks.ToString(); name = $Process.ProcessName }
}

function Read-RefreshWorkerState {
    param([string] $Path)
    # Permit the worker's atomic rename while the supervisor reads the previous state.
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
        ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    $reader = [IO.StreamReader]::new($stream)
    try { return ($reader.ReadToEnd() | ConvertFrom-Json) }
    finally { $reader.Dispose() }
}

function Get-MatchingRefreshProcess {
    param($Identity)
    if ($null -eq $Identity) { return $null }
    $candidate = Get-Process -Id ([int] $Identity.id) -ErrorAction SilentlyContinue
    if ($null -eq $candidate) { return $null }
    if ($candidate.ProcessName -ne $Identity.name -or
        $candidate.StartTime.ToUniversalTime().Ticks.ToString() -ne $Identity.startTicks) { return $null }
    return $candidate
}

function Register-RefreshExcel {
    param($Excel, [int[]] $PreExistingExcelIds)
    if (-not ('ResidentialCare.ExcelWindow' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ResidentialCare {
    public static class ExcelWindow {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    }
}
'@
    }
    [uint32] $excelProcessId = 0
    [void] [ResidentialCare.ExcelWindow]::GetWindowThreadProcessId([IntPtr] $Excel.Hwnd, [ref] $excelProcessId)
    if ($excelProcessId -eq 0 -or $excelProcessId -in $PreExistingExcelIds) {
        throw 'Could not establish ownership of a new Excel process. Refusing to open the workbook.'
    }
    $ownedExcel = Get-Process -Id $excelProcessId -ErrorAction Stop
    if ($ownedExcel.ProcessName -ne 'EXCEL') { throw 'Excel window belongs to an unexpected process.' }
    $script:RefreshState.excel = Get-RefreshProcessIdentity $ownedExcel
    Set-RefreshPhase 'ExcelCreated'
    Write-Log "Owned Excel PID: $excelProcessId"
}

function Update-RefreshChildren {
    param($ExcelIdentity, [hashtable] $Children)
    if ($null -eq (Get-MatchingRefreshProcess $ExcelIdentity)) { return }
    foreach ($child in @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $($ExcelIdentity.id)" -ErrorAction Stop)) {
        if ($child.Name -notlike 'Microsoft.Mashup.Container*.exe') { continue }
        $process = Get-Process -Id $child.ProcessId -ErrorAction SilentlyContinue
        if ($null -ne $process) {
            $identity = Get-RefreshProcessIdentity $process
            $Children["$($identity.id):$($identity.startTicks)"] = $identity
        }
    }
}

function Stop-OwnedRefreshProcesses {
    param($State, [hashtable] $Children, [Diagnostics.Process] $Worker)
    # Never select processes by name alone or infer ownership from creation time proximity.
    if ($null -ne $State -and $null -ne $State.excel) {
        Update-RefreshChildren $State.excel $Children
        $ownedExcel = Get-MatchingRefreshProcess $State.excel
        if ($null -ne $ownedExcel) {
            $saveStatus = if ($State.saved) { 'workbook save was confirmed' } else { 'workbook save is not confirmed' }
            Write-Log "Terminating owned Excel PID $($ownedExcel.Id); $saveStatus." 'WARN'
            $ownedExcel.Kill()
            if (-not $ownedExcel.WaitForExit(10000)) { throw 'Owned Excel did not exit.' }
        }
    }
    foreach ($identity in @($Children.Values)) {
        $child = Get-MatchingRefreshProcess $identity
        if ($null -ne $child) {
            $child.Kill()
            if (-not $child.WaitForExit(10000)) { throw "Power Query worker $($child.Id) did not exit." }
        }
    }
    if (-not $Worker.HasExited -and -not $Worker.WaitForExit(3000)) {
        $Worker.Kill()
        if (-not $Worker.WaitForExit(10000)) { throw 'Refresh worker did not exit.' }
    }
}

function Get-RemainingRefreshProcesses {
    param($ExcelIdentity, [hashtable] $Children)
    $excelProcess = Get-MatchingRefreshProcess $ExcelIdentity
    if ($null -ne $excelProcess) { $excelProcess }
    foreach ($identity in @($Children.Values)) {
        $process = Get-MatchingRefreshProcess $identity
        if ($null -ne $process) { $process }
    }
}

function Get-RefreshStatePhase {
    param($State, [string] $MissingStatePhase = 'StartingWorker')
    if ($null -eq $State) { return $MissingStatePhase }
    $detailProperty = $State.PSObject.Properties['detail']
    if ($null -ne $detailProperty -and -not [string]::IsNullOrWhiteSpace([string] $detailProperty.Value)) {
        return "$($State.phase): $($detailProperty.Value)"
    }
    return [string] $State.phase
}

function Invoke-SupervisedExcelRefresh {
    param([string] $WorkerScript, [hashtable] $Parameters, [double] $TimeoutSeconds,
        [int] $StopGraceSeconds = 5, [ValidateRange(0, 300)] [int] $ExitGraceSeconds = 5)
    if ($TimeoutSeconds -le 0) { throw 'Timeout must be positive.' }
    $stateFolder = if ($Parameters.LogPath) { Split-Path $Parameters.LogPath -Parent } else { [IO.Path]::GetTempPath() }
    $statePath = Join-Path $stateFolder ("refresh-worker-{0}.json" -f [guid]::NewGuid().ToString('N'))
    $parametersForWorker = @{} + $Parameters
    $parametersForWorker.WorkerStatePath = $statePath
    $startInfo = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh -ErrorAction Stop).Source)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    foreach ($argument in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $WorkerScript)) {
        $startInfo.ArgumentList.Add($argument)
    }
    foreach ($key in $parametersForWorker.Keys) {
        if ($null -eq $parametersForWorker[$key]) { continue }
        $startInfo.ArgumentList.Add("-$key")
        $startInfo.ArgumentList.Add([string] $parametersForWorker[$key])
    }
    $worker = [Diagnostics.Process]::Start($startInfo)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $state = $null
    $children = @{}
    $stopStarted = $null
    $lastHeartbeat = -15.0
    $exitCode = 1
    try {
        do {
            if (Test-Path -LiteralPath $statePath) {
                $state = Read-RefreshWorkerState $statePath
            }
            if ($null -ne $state) { Update-RefreshChildren $state.excel $children }
            $stopNow = Test-StopNowRequested
            if ($stopNow -and $null -eq $stopStarted) { $stopStarted = $timer.Elapsed.TotalSeconds }
            if ($stopNow -and ($timer.Elapsed.TotalSeconds - $stopStarted) -ge $StopGraceSeconds) {
                $exitCode = 3
                Write-Log 'Supervisor enforcing stop-now request.' 'WARN'
                break
            }
            if ($timer.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
                $exitCode = 124
                Write-Log 'Supervisor deadline exceeded (includes opening, refresh, save and cleanup).' 'ERROR'
                break
            }
            if ($worker.HasExited) {
                $exitCode = if ($stopNow) { 3 } else { $worker.ExitCode }
                break
            }
            if (($timer.Elapsed.TotalSeconds - $lastHeartbeat) -ge 15) {
                $phase = Get-RefreshStatePhase $state
                Write-Log ("Supervisor: {0}; elapsed {1:n1} minute(s)." -f $phase, ($timer.Elapsed.TotalMinutes))
                $lastHeartbeat = $timer.Elapsed.TotalSeconds
            }
            Start-Sleep -Milliseconds 500
        } while ($true)

        # Re-read the final phase after process exit; success requires saved/closed/quit, not just exit 0.
        if (Test-Path -LiteralPath $statePath) { $state = Read-RefreshWorkerState $statePath }
        if ($exitCode -eq 0) {
            if ($null -eq $state -or $state.phase -ne 'Complete' -or -not $state.saved -or $null -eq $state.excel) {
                throw 'Worker exited without confirmed save/close/quit completion.'
            }
            $exitWait = [Diagnostics.Stopwatch]::StartNew()
            do {
                Update-RefreshChildren $state.excel $children
                $remaining = @(Get-RemainingRefreshProcesses $state.excel $children)
                if ($remaining.Count -eq 0) { break }
                if (Test-StopNowRequested) { $exitCode = 3; break }
                if ($timer.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
                    $exitCode = 124
                    Write-Log 'Supervisor deadline exceeded during process exit; workbook save was confirmed.' 'ERROR'
                    break
                }
                if ($exitWait.Elapsed.TotalSeconds -ge $ExitGraceSeconds) {
                    # Only a successful worker with confirmed save/close/Quit may recover
                    # lingering processes. Never reinterpret an interrupted save as success.
                    Write-Log 'Save, close and Quit confirmed; cleaning up lingering owned Excel/Power Query processes.' 'WARN'
                    Stop-OwnedRefreshProcesses $state $children $worker
                    if (@(Get-RemainingRefreshProcesses $state.excel $children).Count -gt 0) {
                        throw 'Owned processes remain after cleanup; refusing the next workbook.'
                    }
                    break
                }
                Start-Sleep -Milliseconds 500
            } while ($true)
            if ($exitCode -eq 0) {
                Write-Log 'Supervisor confirmed save, workbook closure, and owned Excel/Power Query process exit.'
            }
        }
        if ($exitCode -ne 0) { Stop-OwnedRefreshProcesses $state $children $worker }
    }
    catch {
        Write-Log "Supervisor failure: $($_.Exception.Message)" 'ERROR'
        $exitCode = 1
        Stop-OwnedRefreshProcesses $state $children $worker
    }
    finally {
        $worker.Dispose()
    }
    if ($exitCode -ne 0) {
        $phase = Get-RefreshStatePhase $state 'StartingWorker (Excel ownership not registered)'
        $finalState = if ($exitCode -eq 3) { 'STOPPED' } else { 'FAILED' }
        $saveStatus = if ($null -ne $state -and $state.saved) { 'workbook save was confirmed' } else { 'workbook save is not confirmed' }
        Write-CurrentStatus -State $finalState -Message "Supervisor exit $exitCode during $phase; $saveStatus. State: $statePath"
    }
    return $exitCode
}

function Get-ExcelActivity {
    param($Excel, $Workbook)
    $active = @()
    # Unsupported connection types are ignored explicitly; failed reads on pollable types fail closed.
    $connections = $Workbook.Connections
    try {
        for ($connectionIndex = 1; $connectionIndex -le [int] $connections.Count; $connectionIndex++) {
            $connection = if ([Runtime.InteropServices.Marshal]::IsComObject($connections)) { $connections.Item($connectionIndex) } else { $connections[$connectionIndex - 1] }
            $target = $null
            try {
                if ($null -eq $connection) { throw 'Excel returned a null connection while checking readiness.' }
                $connectionName = [string] $connection.Name
                $connectionType = [int] $connection.Type
                switch ($connectionType) {
                    1 { $target = $connection.OLEDBConnection }
                    2 { $target = $connection.ODBCConnection }
                }
                if ($null -ne $target -and $target.Refreshing) { $active += $connectionName }
            }
            finally {
                Release-ExcelComReference $target
                Release-ExcelComReference $connection
            }
        }
    }
    finally { Release-ExcelComReference $connections }

    $worksheets = $Workbook.Worksheets
    try {
        for ($sheetIndex = 1; $sheetIndex -le [int] $worksheets.Count; $sheetIndex++) {
            $sheet = if ([Runtime.InteropServices.Marshal]::IsComObject($worksheets)) { $worksheets.Item($sheetIndex) } else { $worksheets[$sheetIndex - 1] }
            $queryTables = $null
            try {
                if ($null -eq $sheet) { throw 'Excel returned a null worksheet while checking readiness.' }
                $sheetName = [string] $sheet.Name
                $queryTables = $sheet.QueryTables
                for ($queryIndex = 1; $queryIndex -le [int] $queryTables.Count; $queryIndex++) {
                    $queryTable = if ([Runtime.InteropServices.Marshal]::IsComObject($queryTables)) { $queryTables.Item($queryIndex) } else { $queryTables[$queryIndex - 1] }
                    try {
                        if ($null -eq $queryTable) { throw 'Excel returned a null query table while checking readiness.' }
                        $queryTableName = [string] $queryTable.Name
                        if ($queryTable.Refreshing) { $active += "$sheetName/$queryTableName" }
                    }
                    finally { Release-ExcelComReference $queryTable }
                }
            }
            finally {
                Release-ExcelComReference $queryTables
                Release-ExcelComReference $sheet
            }
        }
    }
    finally { Release-ExcelComReference $worksheets }
    $calculationState = [int] $Excel.CalculationState
    return [pscustomobject] @{ Active = $active; CalculationState = $calculationState }
}

function Wait-ExcelReadyToSave {
    param($Excel, $Workbook)
    # Give Power Query the same minimum 15-second launch/settle interval that the
    # former three-poll gate provided, without touching connection COM objects
    # while they can be invalidated as refresh work starts or completes.
    Set-RefreshPhase 'SettlingAfterRefresh'
    Start-ResponsiveSleep -Seconds 15
    Assert-NotStopNowRequested
    # CalculateUntilAsyncQueriesDone can spin indefinitely for some Power Query
    # workbooks under /automation. Poll the supported connection/query-table
    # refresh flags after the untouched settling interval instead.
    Set-RefreshPhase 'WaitingForConnections'
    # Excel can report all connections idle while its calculation state is still
    # pending. Calling Calculate in that state intermittently fails through COM.
    Wait-ExcelQuietPolls -Excel $Excel -Workbook $Workbook -RequireCalculationDone -RequiredQuietPolls 1
    Set-RefreshPhase 'Calculating'
    $Excel.Calculate()
    Set-RefreshPhase 'CheckingReadiness'
    # The external supervisor bounds every COM readiness read.
    Wait-ExcelQuietPolls -Excel $Excel -Workbook $Workbook -RequireCalculationDone
}

function Wait-ExcelQuietPolls {
    param($Excel, $Workbook, [switch] $RequireCalculationDone, [ValidateRange(1, 10)] [int] $RequiredQuietPolls = 3)
    $quietPolls = 0
    do {
        Assert-NotStopNowRequested
        try { $activity = Get-ExcelActivity $Excel $Workbook }
        catch {
            # Excel can briefly invalidate connection entries as background loads finish.
            # Unknown state resets confirmation; only subsequent successful reads count.
            $quietPolls = 0
            Write-Log "Readiness unavailable; waiting and rechecking: $($_.Exception.Message)" 'WARN'
            Start-ResponsiveSleep -Seconds 5
            continue
        }
        if ($activity.Active.Count -eq 0 -and (-not $RequireCalculationDone -or $activity.CalculationState -eq 0)) { $quietPolls++ } else { $quietPolls = 0 }
        Write-Log "Readiness: $($activity.Active.Count) active query/connection(s); calculation state $($activity.CalculationState); quiet polls $quietPolls/$RequiredQuietPolls. $($activity.Active -join ', ')"
        if ($quietPolls -lt $RequiredQuietPolls) { Start-ResponsiveSleep -Seconds 5 }
    } while ($quietPolls -lt $RequiredQuietPolls)
}
