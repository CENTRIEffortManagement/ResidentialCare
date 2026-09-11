param([int]$ExcelProcessId = 23440)
$ErrorActionPreference = 'Stop'
# Read-only process diagnostics. Does not attach Excel COM, refresh, save, or terminate processes.
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class ExcelWaitDiagnostic {
    [DllImport("advapi32.dll", SetLastError=true)] static extern IntPtr OpenThreadWaitChainSession(uint flags, IntPtr callback);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool GetThreadWaitChain(IntPtr session, IntPtr context, uint flags, uint tid, ref uint count, IntPtr nodes, out int cycle);
    [DllImport("advapi32.dll")] static extern void CloseThreadWaitChainSession(IntPtr session);
    public class Node { public int Type; public int Status; public int PID; public int TID; public string Name; }
    public class Result { public uint TID; public int Error; public bool Cycle; public List<Node> Nodes = new List<Node>(); }
    public static Result Read(uint tid) {
        var result = new Result { TID = tid };
        var session = OpenThreadWaitChainSession(0, IntPtr.Zero);
        if (session == IntPtr.Zero) { result.Error = Marshal.GetLastWin32Error(); return result; }
        var buffer = Marshal.AllocHGlobal(280 * 16);
        try {
            uint count = 16; int cycle;
            if (!GetThreadWaitChain(session, IntPtr.Zero, 7, tid, ref count, buffer, out cycle)) {
                result.Error = Marshal.GetLastWin32Error(); return result;
            }
            result.Cycle = cycle != 0;
            for (int i=0; i<Math.Min(count,16); i++) {
                var p = IntPtr.Add(buffer, i*280);
                var node = new Node { Type = Marshal.ReadInt32(p,0), Status = Marshal.ReadInt32(p,4) };
                if (node.Type == 8) { node.PID = Marshal.ReadInt32(p,8); node.TID = Marshal.ReadInt32(p,12); }
                else { node.Name = Marshal.PtrToStringUni(IntPtr.Add(p,8),128).Split('\0')[0]; }
                result.Nodes.Add(node);
            }
            return result;
        } finally { Marshal.FreeHGlobal(buffer); CloseThreadWaitChainSession(session); }
    }
}
'@
$excel = Get-Process -Id $ExcelProcessId
$threadRows = @($excel.Threads | ForEach-Object {
    [pscustomobject]@{TID=$_.Id; State=$_.ThreadState.ToString(); CPUSeconds=$_.TotalProcessorTime.TotalSeconds; WaitChain=[ExcelWaitDiagnostic]::Read([uint32]$_.Id)}
})
$processes = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(EXCEL|Microsoft.Mashup.Container.Loader)\.exe$' } |
    Select-Object ProcessId,ParentProcessId,Name,CreationDate,KernelModeTime,UserModeTime,ReadTransferCount,WriteTransferCount,WorkingSetSize)
$events = @(Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-2); Id=1000,1001,1002,1025,1026} -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match 'EXCEL.EXE|Microsoft.Mashup' } | Select-Object -First 60 TimeCreated,Id,ProviderName,Message)
$memory = Get-CimInstance Win32_OperatingSystem | Select-Object FreePhysicalMemory,TotalVisibleMemorySize
$office = Get-ItemProperty 'HKLM:/SOFTWARE/Microsoft/Office/ClickToRun/Configuration' | Select-Object VersionToReport,Platform,UpdateChannel
$report = [pscustomobject]@{CapturedAt=(Get-Date).ToString('o'); ExcelPID=$ExcelProcessId; Responding=$excel.Responding; Title=$excel.MainWindowTitle; Memory=$memory; Office=$office; Processes=$processes; Threads=$threadRows; Events=$events}
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'capture.json') -Encoding utf8
$threadRows | Where-Object { $_.WaitChain.Cycle -or $_.WaitChain.Error -ne 0 -or $_.WaitChain.Nodes.Count -gt 1 } | ConvertTo-Json -Depth 6 -Compress
$events | Group-Object Id | Select-Object Name,Count | ConvertTo-Json -Compress
