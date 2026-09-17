param(
    [Parameter(Mandatory)] [string] $RequestPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'RefreshRunGate.ps1')
$request = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json
$lease = $null; $process = $null
function Test-BatchParent {
    $owner = Get-Process -Id $request.ParentId -ErrorAction SilentlyContinue
    return ($null -ne $owner -and $owner.StartTime.ToUniversalTime().Ticks -eq $request.ParentStarted)
}
try {
    if (-not (Test-BatchParent)) { throw 'Batch coordinator no longer exists; Excel was not started.' }
    $lease = Enter-RefreshRunGate -DateRoot $request.DateRoot -BatchWorker
    if (-not (Test-BatchParent)) { throw 'Batch coordinator exited before worker startup.' }
    $info = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh).Source)
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    foreach ($arg in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'Invoke-AllUnitsExcelWorkbookRefresh.ps1'),
        '-WorkbookPath', $request.Path, '-WorkbookName', $request.Id, '-LogPath', $request.LogPath,
        '-StatusPath', $request.StatusPath, '-StopRequestPath', $request.StopPath,
        '-VisibleOverride', $request.Visible, '-TimeoutMinutesOverride', $request.TimeoutMinutes)) {
        $info.ArgumentList.Add([string] $arg)
    }
    if ($null -ne $request.PSObject.Properties['InputSnapshotPath']) {
        $info.ArgumentList.Add('-InputSnapshotPath')
        $info.ArgumentList.Add([string] $request.InputSnapshotPath)
    }
    $process = [Diagnostics.Process]::Start($info)
    while (-not $process.WaitForExit(250)) {
        if (-not (Test-BatchParent)) {
            [IO.File]::WriteAllText($request.StopPath, "mode=Now`n")
        }
    }
    exit $process.ExitCode
}
catch { Write-Error $_ -ErrorAction Continue; exit 1 }
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        # Let the independent supervisor enforce cancellation and owned-process cleanup.
        [IO.File]::WriteAllText($request.StopPath, "mode=Now`n")
        $process.WaitForExit()
    }
    if ($null -ne $process) { $process.Dispose() }
    if ($null -ne $lease) { $lease.Dispose() }
}
