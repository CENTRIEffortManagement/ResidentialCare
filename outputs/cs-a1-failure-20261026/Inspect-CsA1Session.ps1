[CmdletBinding()]
param([switch]$Worker, [string]$RunTag = (Get-Date -Format 'yyyyMMdd-HHmmss'))
$ErrorActionPreference = 'Stop'
$log = Join-Path $PSScriptRoot ("live-session-$RunTag.log")
if (-not $Worker) {
    $observerArgs = @('-NoProfile','-NonInteractive','-STA','-File',('"' + $PSCommandPath + '"'),'-Worker','-RunTag',$RunTag)
    $observer = Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList $observerArgs -WindowStyle Hidden -PassThru
    if (-not $observer.WaitForExit(30000)) {
        # Only the observer created above is stopped. No Excel/Mashup process is controlled.
        $observer.Kill()
        [IO.File]::AppendAllText($log, '{"phase":"Observer timed out; Excel left untouched"}' + [Environment]::NewLine)
    }
    if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log }
    exit
}
function Record([string]$Phase, $Data = $null) {
    $item = @{phase=$Phase; timeUtc=[datetime]::UtcNow.ToString('o'); data=$Data}
    [IO.File]::AppendAllText($log, ($item | ConvertTo-Json -Depth 5 -Compress) + [Environment]::NewLine)
}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CsLiveRead {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objid, ref Guid iid,
        [MarshalAs(UnmanagedType.IDispatch)] out object obj);
}
'@
Record 'Checking exact frozen window'
[IntPtr]$window = 7082060
[uint32]$windowProcess = 0
[void][CsLiveRead]::GetWindowThreadProcessId($window,[ref]$windowProcess)
if ($windowProcess -ne 14132) { Record 'Window identity changed; observer stopped'; exit }
$desktop = [CsLiveRead]::FindWindowEx($window,[IntPtr]::Zero,'XLDESK',$null)
$excelWindow = [CsLiveRead]::FindWindowEx($desktop,[IntPtr]::Zero,'EXCEL7',$null)
if ($excelWindow -eq [IntPtr]::Zero) { Record 'Native Excel window not found'; exit }
try {
Record 'Reading native Excel object model'
$iid = [guid]'00020400-0000-0000-C000-000000000046'
$native = $null
$result = [CsLiveRead]::AccessibleObjectFromWindow($excelWindow,[uint32]4294967280,[ref]$iid,[ref]$native)
if ($result -ne 0) { Record 'Object model unavailable' @{HRESULT=$result}; exit }
$app = $native.Application
Record 'Reading application state'
Record 'Application' @{Ready=$app.Ready; StatusBar=$app.StatusBar; Calculation=$app.Calculation;
    CalculationState=$app.CalculationState; DisplayAlerts=$app.DisplayAlerts; Interactive=$app.Interactive;
    ScreenUpdating=$app.ScreenUpdating; EnableEvents=$app.EnableEvents}
$expected = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx'))
Record 'Locating exact CS workbook'
$book = $app.Workbooks.Item('CapacityDistrib(A.1)-shifts.xlsx')
if ($book.FullName -ne $expected) { Record 'Different workbook path' @{Path=$book.FullName}; exit }
Record 'Exact workbook confirmed' @{Path=$book.FullName; ReadOnly=$book.ReadOnly; Saved=$book.Saved}
foreach ($connection in $book.Connections) {
    if ($connection.Type -eq 1) {
        $name = $connection.Name
        Record 'Reading connection' @{Name=$name}
        Record 'Connection' @{Name=$name; Refreshing=$connection.OLEDBConnection.Refreshing;
            BackgroundQuery=$connection.OLEDBConnection.BackgroundQuery; RefreshWithRefreshAll=$connection.RefreshWithRefreshAll}
    }
}
Record 'Completed read-only inspection'
} catch { Record 'Read-only inspection unavailable' @{Message=$_.Exception.Message; HRESULT=$_.Exception.HResult} }
