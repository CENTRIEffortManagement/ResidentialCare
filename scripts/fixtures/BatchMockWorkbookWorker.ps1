# Test-only replacement supervisor. Its targets must be synthetic .data files.
param($WorkbookPath, $WorkbookName, $LogPath, $StatusPath, $StopRequestPath, $VisibleOverride, $TimeoutMinutesOverride, $InputSnapshotPath)
$ErrorActionPreference = 'Stop'
if ([IO.Path]::GetExtension($WorkbookPath) -ne '.data') { throw 'Mock worker refuses workbook targets.' }
$verifiedInputs = 0
if ($InputSnapshotPath) {
    $snapshot = Get-Content -LiteralPath $InputSnapshotPath -Raw | ConvertFrom-Json
    if ($snapshot.Target -ne $WorkbookPath) { throw 'Mock input snapshot target mismatch.' }
    foreach ($entry in $snapshot.Inputs.PSObject.Properties) {
        if ([IO.Path]::GetExtension($entry.Name) -ne '.data') { throw 'Mock worker refuses workbook inputs.' }
        $item = Get-Item -LiteralPath $entry.Name
        if ("$($item.Length):$($item.LastWriteTimeUtc.Ticks)" -ne $entry.Value) { throw 'Mock input changed before refresh.' }
        $verifiedInputs++
    }
}
Start-Sleep -Milliseconds 200
$stopped = Test-Path -LiteralPath $StopRequestPath
if (-not $stopped) { [IO.File]::WriteAllText($WorkbookPath, 'mock saved output') }
$state = @{ phase = $(if ($stopped) { 'Refreshing' } else { 'Complete' }); saved = (-not $stopped); verifiedInputs = $verifiedInputs }
$state | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path $LogPath -Parent) 'refresh-worker-mock.json')
if ($stopped) { exit 3 }; exit 0
