# Test-only helper. Call only with disposable synthetic JSON paths.
param([string] $Path, [string] $ReadyPath, [int] $HoldMilliseconds = 700, [switch] $ReadOnly)
$ErrorActionPreference = 'Stop'
$stream = $null
$originalAttributes = [IO.File]::GetAttributes($Path)
try {
    if ($ReadOnly) { [IO.File]::SetAttributes($Path, ($originalAttributes -bor [IO.FileAttributes]::ReadOnly)) }
    else { $stream = [IO.File]::Open($Path, 'Open', 'Read', [IO.FileShare]::ReadWrite) }
    [IO.File]::WriteAllText($ReadyPath, 'ready')
    Start-Sleep -Milliseconds $HoldMilliseconds
} finally {
    if ($null -ne $stream) { $stream.Dispose() }
    if ($ReadOnly) { [IO.File]::SetAttributes($Path, $originalAttributes) }
}
