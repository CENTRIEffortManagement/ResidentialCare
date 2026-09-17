# Atomic publication for coordinator state and supervised-worker receipts.
# Never truncate the published file or change its permissions to bypass a reader.
function Write-RefreshJson {
    param([string] $Path, $Value, [int] $Depth = 40,
        [ValidateRange(0, 60000)] [int] $RetryMilliseconds = 30000)
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth $Depth))
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        try { [IO.File]::Move($temporary, $Path, $true); return }
        catch {
            # PowerShell wraps .NET exceptions; Windows replacement can report
            # access denied (5), sharing violation (32), or lock violation (33).
            $cause = $_.Exception.GetBaseException()
            $nativeCode = $cause.HResult -band 0xffff
            $retryable = $cause -is [UnauthorizedAccessException] -or
                ($cause -is [IO.IOException] -and $nativeCode -in @(5, 32, 33))
            if ($retryable -and $timer.ElapsedMilliseconds -lt $RetryMilliseconds) {
                Start-Sleep -Milliseconds ([Math]::Min(100, [Math]::Max(1, $RetryMilliseconds - $timer.ElapsedMilliseconds)))
                continue
            }
            # Retain both the old valid JSON and the prepared replacement. The
            # latter is evidence for inspection, never an automatic resume input.
            $errorMessage = "Cannot publish JSON '$Path' after $($timer.ElapsedMilliseconds) ms: $($cause.Message) (Windows code $nativeCode). Prepared recovery snapshot: '$temporary'. Close any application holding this status file before retrying."
            $failure = [IO.IOException]::new($errorMessage, $cause)
            $failure.Data['TargetPath'] = $Path
            $failure.Data['RecoveryPath'] = $temporary
            throw $failure
        }
    }
}
