[CmdletBinding()]
param(
    [switch] $SkipWaitChain,
    [int[]] $TargetProcessIds = @(14132, 71952),
    [long] $KnownUiWindow = 7082060,
    [int] $KnownUiProcessId = 14132,
    [ValidateRange(250, 5000)] [int] $SampleMilliseconds = 1000,
    [ValidateRange(1000, 15000)] [int] $WaitChainDeadlineMilliseconds = 8000,
    [ValidateRange(1, 6)] [int] $MaxWaitChainThreadsPerProcess = 4,
    [ValidatePattern('^[0-9]{8}-[0-9]{6}-[0-9]{3}$')]
    [string] $RunTag = (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
)

# Read-only process/window/thread inspection. No Excel COM attachment, workbook
# access, refresh, save, close, process termination, suspension or privilege change.
# Only this script's output JSON is written, beside this script.
# SDK contracts:
# https://learn.microsoft.com/en-us/windows/win32/api/wct/nf-wct-getthreadwaitchain
# https://learn.microsoft.com/en-us/windows/win32/api/wct/ns-wct-waitchain_node_info
# https://learn.microsoft.com/en-us/windows/win32/api/wct/nf-wct-registerwaitchaincomcallback
$ErrorActionPreference = 'Stop'
$outputPath = Join-Path $PSScriptRoot "live-wait-$RunTag.json"
if (Test-Path -LiteralPath $outputPath) { throw 'Evidence already exists; choose a new RunTag.' }
if (-not $SkipWaitChain) {
    throw 'WCT is disabled after a diagnostic-host native crash. Use -SkipWaitChain for window/thread-only inspection.'
}

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

public sealed class CsLiveWindow {
    public long Handle; public uint ProcessId; public uint ThreadId;
    public string Title; public string ClassName; public bool Visible;
    public bool Enabled; public long Owner;
}
public sealed class CsLiveWaitNode {
    public int Index; public int ObjectTypeValue; public string ObjectType;
    public int StatusValue; public string Status;
    public uint ProcessId; public uint ThreadId;
    public uint WaitTimeRaw; public uint ContextSwitches;
    public string ObjectName;
}
public sealed class CsLiveWaitResult {
    public uint RequestedThreadId; public bool Success; public int Win32Error;
    public string Error; public uint ReportedNodeCount; public bool IsCycle;
    public CsLiveWaitNode[] Nodes = new CsLiveWaitNode[0];
}
public static class CsLiveWaitNative {
    private delegate bool EnumWindowCallback(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll", SetLastError=true)]
    private static extern bool EnumWindows(EnumWindowCallback callback, IntPtr data);
    [DllImport("user32.dll", SetLastError=true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern IntPtr LoadLibrary(string name);
    [DllImport("kernel32.dll", CharSet=CharSet.Ansi, ExactSpelling=true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("advapi32.dll")]
    private static extern void RegisterWaitChainCOMCallback(IntPtr callState, IntPtr activationState);
    [DllImport("advapi32.dll", SetLastError=true)]
    private static extern IntPtr OpenThreadWaitChainSession(uint flags, IntPtr callback);
    [DllImport("advapi32.dll", SetLastError=true)]
    private static extern bool GetThreadWaitChain(IntPtr session, UIntPtr context,
        uint flags, uint threadId, ref uint count, IntPtr nodes, out int cycle);
    [DllImport("advapi32.dll")]
    private static extern void CloseThreadWaitChainSession(IntPtr session);

    // Sequential SDK layout: two 32-bit enums followed by the 8-byte-aligned
    // union. Its largest member is WCHAR[128], LARGE_INTEGER, BOOL (280 total).
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    private struct NodeLayout {
        public int Type; public int Status;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)] public string Name;
        public long ReservedTimeout; public int ReservedAlertable;
    }
    private static IntPtr ole32Module;
    private static readonly string[] TypeNames = {"Invalid", "CriticalSection",
        "SendMessage", "Mutex", "Alpc", "Com", "ThreadWait", "ProcessWait",
        "Thread", "ComActivation", "Unknown", "SocketIo", "SmbIo", "Max"};
    private static readonly string[] StatusNames = {"Invalid", "NoAccess", "Running",
        "Blocked", "PidOnly", "PidOnlyRpcss", "Owned", "NotOwned", "Abandoned",
        "Unknown", "Error", "Max"};
    private static string Label(string[] labels, int value) {
        return value >= 0 && value < labels.Length ? labels[value] : "Value:" + value;
    }
    public static bool RegisterComCallbacks() {
        // Load only into this diagnostic process; no target-process injection.
        ole32Module = LoadLibrary("ole32.dll");
        if (ole32Module == IntPtr.Zero) return false;
        IntPtr calls = GetProcAddress(ole32Module, "CoGetCallState");
        IntPtr activation = GetProcAddress(ole32Module, "CoGetActivationState");
        if (calls == IntPtr.Zero || activation == IntPtr.Zero) return false;
        RegisterWaitChainCOMCallback(calls, activation);
        return true;
    }
    public static CsLiveWindow DescribeWindow(long handle, int expectedProcessId) {
        IntPtr hwnd = new IntPtr(handle);
        uint processId;
        uint threadId = GetWindowThreadProcessId(hwnd, out processId);
        if (threadId == 0 || processId != (uint)expectedProcessId) return null;
        var title = new StringBuilder(2048);
        var cls = new StringBuilder(256);
        GetWindowText(hwnd, title, title.Capacity);
        GetClassName(hwnd, cls, cls.Capacity);
        return new CsLiveWindow { Handle=handle, ProcessId=processId, ThreadId=threadId,
            Title=title.ToString(), ClassName=cls.ToString(), Visible=IsWindowVisible(hwnd),
            Enabled=IsWindowEnabled(hwnd), Owner=GetWindow(hwnd, 4).ToInt64() };
    }
    public static CsLiveWindow[] WindowsForProcesses(int[] processIds) {
        var allowed = new HashSet<int>(processIds);
        var result = new List<CsLiveWindow>();
        EnumWindowCallback callback = delegate(IntPtr hwnd, IntPtr data) {
            uint owner; GetWindowThreadProcessId(hwnd, out owner);
            if (allowed.Contains((int)owner)) {
                CsLiveWindow window = DescribeWindow(hwnd.ToInt64(), (int)owner);
                if (window != null) result.Add(window);
            }
            return true;
        };
        if (!EnumWindows(callback, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        GC.KeepAlive(callback);
        return result.ToArray();
    }
    private static CsLiveWaitResult ReadChain(uint threadId) {
        var result = new CsLiveWaitResult {RequestedThreadId=threadId};
        IntPtr session = IntPtr.Zero, buffer = IntPtr.Zero;
        try {
            int stride = Marshal.SizeOf(typeof(NodeLayout));
            if (stride != 280) throw new InvalidOperationException("Unexpected WCT layout size: " + stride);
            session = OpenThreadWaitChainSession(0, IntPtr.Zero);
            if (session == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            buffer = Marshal.AllocHGlobal(stride * 16);
            Marshal.Copy(new byte[stride * 16], 0, buffer, stride * 16);
            uint count = 16; int cycle;
            // WCTP_GETINFO_ALL_FLAGS = OUT_OF_PROC | OUT_OF_PROC_COM | OUT_OF_PROC_CS.
            bool ok = GetThreadWaitChain(session, UIntPtr.Zero, 0x7, threadId,
                ref count, buffer, out cycle);
            int error = ok ? 0 : Marshal.GetLastWin32Error();
            result.Success=ok; result.Win32Error=error; result.ReportedNodeCount=count;
            result.IsCycle=cycle != 0;
            if (!ok) result.Error = new Win32Exception(error).Message;
            // ERROR_MORE_DATA still supplies a valid partial chain.
            if (!ok && error != 234) return result;
            var nodes = new List<CsLiveWaitNode>();
            for (int i=0; i < Math.Min(count, 16u); i++) {
                IntPtr node = IntPtr.Add(buffer, i * stride);
                int type = Marshal.ReadInt32(node, 0), status = Marshal.ReadInt32(node, 4);
                var entry = new CsLiveWaitNode {Index=i, ObjectTypeValue=type,
                    ObjectType=Label(TypeNames,type), StatusValue=status, Status=Label(StatusNames,status)};
                if (type == 8) {
                    entry.ProcessId=unchecked((uint)Marshal.ReadInt32(node, 8));
                    entry.ThreadId=unchecked((uint)Marshal.ReadInt32(node, 12));
                    entry.WaitTimeRaw=unchecked((uint)Marshal.ReadInt32(node, 16));
                    entry.ContextSwitches=unchecked((uint)Marshal.ReadInt32(node, 20));
                } else {
                    string name = Marshal.PtrToStringUni(IntPtr.Add(node, 8), 128);
                    int end = name.IndexOf('\0');
                    entry.ObjectName = end >= 0 ? name.Substring(0,end) : name;
                }
                nodes.Add(entry);
            }
            result.Nodes=nodes.ToArray();
        } catch (Exception ex) {result.Error=ex.ToString();}
        finally {
            if (buffer != IntPtr.Zero) Marshal.FreeHGlobal(buffer);
            if (session != IntPtr.Zero) CloseThreadWaitChainSession(session);
        }
        return result;
    }
    public static Task<CsLiveWaitResult> BeginChain(uint threadId) {
        // A slow provider cannot block the PowerShell diagnostic indefinitely.
        // Each task owns/frees its own WCT handle and buffer; it never owns Excel.
        return Task.Run(() => ReadChain(threadId));
    }
}
'@

function Get-LiveThreadSnapshot([int] $TargetId) {
    $errors = [Collections.Generic.List[string]]::new()
    try {
        $targetProcess = Get-Process -Id $TargetId -ErrorAction Stop
        if ($targetProcess.ProcessName -ne 'EXCEL') { throw 'PID no longer identifies EXCEL; no substitute process inspected.' }
        $start = $targetProcess.StartTime.ToUniversalTime()
        $threads = @(
            foreach ($thread in $targetProcess.Threads) {
                $state = $null; $reason = $null; $cpu = $null; $threadStart = $null
                try { $state = $thread.ThreadState.ToString() } catch { $errors.Add($_.Exception.Message) }
                if ($state -eq 'Wait') {
                    try { $reason = $thread.WaitReason.ToString() } catch { $errors.Add($_.Exception.Message) }
                }
                try { $cpu = $thread.TotalProcessorTime.TotalMilliseconds } catch { $errors.Add($_.Exception.Message) }
                try { $threadStart = $thread.StartTime.ToUniversalTime().ToString('o') } catch { $errors.Add($_.Exception.Message) }
                [pscustomobject]@{ThreadId=$thread.Id; State=$state; WaitReason=$reason; CpuMilliseconds=$cpu; StartUtc=$threadStart}
            }
        )
        return [pscustomobject]@{ProcessId=$TargetId; ProcessName=$targetProcess.ProcessName;
            StartUtc=$start.ToString('o'); StartTicks=$start.Ticks; WorkingSetBytes=$targetProcess.WorkingSet64;
            Threads=$threads; Errors=@($errors)}
    }
    catch { return [pscustomobject]@{ProcessId=$TargetId; Error=$_.Exception.Message; Threads=@()} }
}

$startedUtc = [datetime]::UtcNow.ToString('o')
$before = @($TargetProcessIds | ForEach-Object { Get-LiveThreadSnapshot $_ })
$eligibleIds = @($before | Where-Object { $_.PSObject.Properties['StartTicks'] } | ForEach-Object ProcessId)
$windows = @([CsLiveWaitNative]::WindowsForProcesses([int[]]$eligibleIds))
$knownWindow = if ($eligibleIds -contains $KnownUiProcessId) {
    [CsLiveWaitNative]::DescribeWindow($KnownUiWindow, $KnownUiProcessId)
} else { $null }
$knownWindowConfirmed = $null -ne $knownWindow -and $knownWindow.ProcessId -eq $KnownUiProcessId
$sampleTimer = [Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Milliseconds $SampleMilliseconds
$after = @($TargetProcessIds | ForEach-Object { Get-LiveThreadSnapshot $_ })
$sampleTimer.Stop()
$deltas = @(
    foreach ($current in $after) {
        $old = @($before | Where-Object ProcessId -eq $current.ProcessId)
        if ($old.Count -ne 1 -or -not $current.PSObject.Properties['StartTicks'] -or
            $old[0].StartTicks -ne $current.StartTicks) { continue }
        $oldThreads = @{}; foreach ($thread in $old[0].Threads) { $oldThreads[$thread.ThreadId] = $thread }
        foreach ($thread in $current.Threads) {
            $delta = $null
            if ($oldThreads.ContainsKey($thread.ThreadId) -and $null -ne $thread.CpuMilliseconds -and
                $null -ne $oldThreads[$thread.ThreadId].CpuMilliseconds) {
                $delta = $thread.CpuMilliseconds - $oldThreads[$thread.ThreadId].CpuMilliseconds
            }
            [pscustomobject]@{ProcessId=$current.ProcessId; ThreadId=$thread.ThreadId;
                CpuDeltaMilliseconds=$delta; State=$thread.State; WaitReason=$thread.WaitReason; StartUtc=$thread.StartUtc}
        }
    }
)
$comCallbacks = $false
$requests = [Collections.Generic.List[object]]::new()
$tasks = [Collections.Generic.List[System.Threading.Tasks.Task]]::new()
if (-not $SkipWaitChain) {
    $comCallbacks = [CsLiveWaitNative]::RegisterComCallbacks()
foreach ($processSnapshot in $after) {
    if (-not $processSnapshot.PSObject.Properties['StartTicks']) { continue }
    $priorIdentity = @($before | Where-Object ProcessId -eq $processSnapshot.ProcessId)
    if ($priorIdentity.Count -ne 1 -or -not $priorIdentity[0].PSObject.Properties['StartTicks'] -or
        $priorIdentity[0].StartTicks -ne $processSnapshot.StartTicks) { continue }
    $candidates = [Collections.Generic.List[object]]::new()
    if ($knownWindowConfirmed -and $processSnapshot.ProcessId -eq $KnownUiProcessId) {
        $candidates.Add([pscustomobject]@{ThreadId=$knownWindow.ThreadId; Basis='User-confirmed exact Excel window'})
    }
    foreach ($window in @($windows | Where-Object { $_.ProcessId -eq $processSnapshot.ProcessId -and $_.ClassName -eq 'XLMAIN' })) {
        $candidates.Add([pscustomobject]@{ThreadId=$window.ThreadId; Basis='XLMAIN top-level window owner'})
    }
    $earliest = @($processSnapshot.Threads | Where-Object StartUtc | Sort-Object StartUtc,ThreadId | Select-Object -First 1)
    if ($earliest.Count) { $candidates.Add([pscustomobject]@{ThreadId=$earliest[0].ThreadId; Basis='Earliest observed thread start; startup-thread candidate, not proof'}) }
    foreach ($window in @($windows | Where-Object ProcessId -eq $processSnapshot.ProcessId)) {
        $candidates.Add([pscustomobject]@{ThreadId=$window.ThreadId; Basis='Other target-process top-level window owner'})
    }
    $seen = @{}; $accepted = 0
    foreach ($candidate in $candidates) {
        if ($seen.ContainsKey($candidate.ThreadId) -or $accepted -ge $MaxWaitChainThreadsPerProcess) { continue }
        $seen[$candidate.ThreadId] = $true; $accepted++
        $task = [CsLiveWaitNative]::BeginChain([uint32]$candidate.ThreadId)
        $tasks.Add($task)
        $requests.Add([pscustomobject]@{ProcessId=$processSnapshot.ProcessId; ThreadId=$candidate.ThreadId; Basis=$candidate.Basis; Task=$task})
    }
}
if ($tasks.Count) { try { [void][System.Threading.Tasks.Task]::WaitAll($tasks.ToArray(), $WaitChainDeadlineMilliseconds) } catch {} }
}
$waitChains = @(
    foreach ($request in $requests) {
        $result = $null; $errorText = $null
        if ($request.Task.IsCompletedSuccessfully) { $result = $request.Task.Result }
        elseif ($request.Task.IsFaulted) { $errorText = $request.Task.Exception.ToString() }
        else { $errorText = 'WCT deadline exceeded; no target process was interrupted.' }
        [pscustomobject]@{ProcessId=$request.ProcessId; ThreadId=$request.ThreadId;
            IdentificationBasis=$request.Basis; Result=$result; Error=$errorText}
    }
)
$evidence = [ordered]@{
    StartedUtc=$startedUtc; FinishedUtc=[datetime]::UtcNow.ToString('o')
    TargetProcessIds=$TargetProcessIds; ReadOnlyInspection=$true
    KnownUiWindow=$KnownUiWindow; KnownUiWindowConfirmed=$knownWindowConfirmed
    KnownUiWindowDetails=$(if ($knownWindowConfirmed) { $knownWindow } else { $null })
    WaitChainSkipped=[bool]$SkipWaitChain; WctDisabledReason='Previous diagnostic host terminated with native access violation; no WCT retry.'
    ComWaitCallbacksRegistered=$comCallbacks; SampleElapsedMilliseconds=$sampleTimer.Elapsed.TotalMilliseconds
    Before=$before; After=$after; ThreadCpuAndWaitSamples=$deltas; TargetWindows=$windows; WaitChains=$waitChains
    Limitations=@('A single WCT node can mean running or an unsupported wait; it does not prove absence of a hang.',
        'Access-denied or NoAccess results are incomplete evidence; this script does not change token privileges.',
        'Thread state and wait chains are point-in-time observations, not an M query execution trace.',
        'Earliest-start selection is explicitly a candidate when no main-window identity is available.')
}
$evidence | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $outputPath -Encoding utf8
Write-Output $outputPath
$waitChains | Select-Object ProcessId,ThreadId,IdentificationBasis,Error,@{Name='Cycle';Expression={$_.Result.IsCycle}},@{Name='Success';Expression={$_.Result.Success}} | Format-Table -AutoSize
