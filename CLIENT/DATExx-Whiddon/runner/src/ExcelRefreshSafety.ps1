# Shared by the standalone and integrated all-units refresh engines.
# The supervisor never calls Excel COM: deadlines remain enforceable during Save/Open/Quit.
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
    param([string] $Path, [int] $AllowedProcessId = 0)
    $blockers = @(Get-RefreshFileUsers $Path | Where-Object { $_.Process.Id -ne $AllowedProcessId })
    if ($blockers.Count -gt 0) {
        $description = ($blockers | ForEach-Object { "$($_.AppName) (PID $($_.Process.Id))" }) -join ', '
        throw "Workbook is in use by $description. Close/disconnect those applications before refresh: $Path"
    }
}

function Set-RefreshPhase {
    param([string] $Phase)
    $script:RefreshState.phase = $Phase
    $script:RefreshState.updatedAt = (Get-Date).ToString('o')
    $temporaryState = $WorkerStatePath + '.tmp'
    [IO.File]::WriteAllText($temporaryState, ($script:RefreshState | ConvertTo-Json -Depth 5))
    [IO.File]::Move($temporaryState, $WorkerStatePath, $true)
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

function Invoke-SupervisedExcelRefresh {
    param([string] $WorkerScript, [hashtable] $Parameters, [double] $TimeoutSeconds,
        [int] $StopGraceSeconds = 5, [ValidateRange(0, 300)] [int] $ExitGraceSeconds = 15)
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
                $phase = if ($null -eq $state) { 'StartingWorker' } else { $state.phase }
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
        $phase = if ($null -eq $state) { 'StartingWorker (Excel ownership not registered)' } else { $state.phase }
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
    foreach ($connection in @($Workbook.Connections)) {
        if ($null -eq $connection) { throw 'Excel returned a null connection while checking readiness.' }
        switch ([int] $connection.Type) {
            1 { if ($connection.OLEDBConnection.Refreshing) { $active += [string] $connection.Name } }
            2 { if ($connection.ODBCConnection.Refreshing) { $active += [string] $connection.Name } }
        }
    }
    foreach ($sheet in @($Workbook.Worksheets)) {
        foreach ($queryTable in @($sheet.QueryTables)) {
            if ($null -eq $queryTable) { throw 'Excel returned a null query table while checking readiness.' }
            if ($queryTable.Refreshing) { $active += "$($sheet.Name)/$($queryTable.Name)" }
        }
    }
    return [pscustomobject] @{ Active = $active; CalculationState = [int] $Excel.CalculationState }
}

function Wait-ExcelReadyToSave {
    param($Excel, $Workbook)
    # Let RefreshAll finish its background connection work before entering the
    # synchronous async-drain call; keep that call out of the active connection phase.
    Set-RefreshPhase 'WaitingForConnections'
    Wait-ExcelQuietPolls -Excel $Excel -Workbook $Workbook
    Set-RefreshPhase 'WaitingForAsyncQueries'
    # Unlike connection flags alone, this drains pending OLE DB and OLAP query work.
    # The external supervisor enforces the deadline if this COM method never returns.
    $Excel.CalculateUntilAsyncQueriesDone()
    Assert-NotStopNowRequested
    Set-RefreshPhase 'Calculating'
    $Excel.Calculate()
    Set-RefreshPhase 'CheckingReadiness'
    Wait-ExcelQuietPolls -Excel $Excel -Workbook $Workbook -RequireCalculationDone
}

function Wait-ExcelQuietPolls {
    param($Excel, $Workbook, [switch] $RequireCalculationDone)
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
        Write-Log "Readiness: $($activity.Active.Count) active query/connection(s); calculation state $($activity.CalculationState); quiet polls $quietPolls/3. $($activity.Active -join ', ')"
        if ($quietPolls -lt 3) { Start-ResponsiveSleep -Seconds 5 }
    } while ($quietPolls -lt 3)
}
