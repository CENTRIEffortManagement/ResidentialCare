[CmdletBinding()]
param(
    [string] $ClientRoot = 'CLIENT/DATExx-Whiddon',
    [int] $ExpectedWorkbookCount = 131,
    [int] $CalculationTimeoutMinutes = 15,
    [ValidateRange(1, 2147483647)]
    [int] $StartAtSequence = 1,
    [string] $PriorResultsCsv,
    [int[]] $OnlySequences
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($CalculationTimeoutMinutes -lt 1) {
    throw 'CalculationTimeoutMinutes must be at least 1.'
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedClientRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $ClientRoot))
$resolvedRepoRoot = [IO.Path]::GetFullPath($repoRoot)

if (-not $resolvedClientRoot.StartsWith($resolvedRepoRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "ClientRoot must resolve inside the repository: $ClientRoot"
}
if (-not (Test-Path -LiteralPath $resolvedClientRoot -PathType Container)) {
    throw "Client root does not exist: $resolvedClientRoot"
}

$extensions = @('.xlsx', '.xlsm', '.xlsb', '.xls')
$workbooks = @(
    Get-ChildItem -LiteralPath $resolvedClientRoot -Recurse -File |
        Where-Object {
            $extension = $_.Extension.ToLowerInvariant()
            $relativePath = [IO.Path]::GetRelativePath($resolvedClientRoot, $_.FullName)
            $extensions -contains $extension -and
            $_.Name -notlike '~$*' -and
            $relativePath -notmatch '(?i)(^|[\\/])3\. Report([\\/]|$)' -and
            $relativePath -notmatch '(?i)(^|[\\/])ExcelPQRoundTrip_Backups([\\/]|$)' -and
            $relativePath -notmatch '(?i)^UNITS[\\/](Unit1|Unit2)([\\/]|$)' -and
            $_.Name -notmatch '(?i)^TEMPLATE\.xlsx$' -and
            $_.Name -notmatch '(?i)Old\.xlsx$' -and
            $_.Name -notmatch '(?i) - Copy\.xlsx$'
        } |
        Sort-Object FullName
)

if ($workbooks.Count -ne $ExpectedWorkbookCount) {
    throw "Safety check failed: expected $ExpectedWorkbookCount workbooks but resolved $($workbooks.Count). No workbook was opened."
}
if ($StartAtSequence -gt $workbooks.Count) {
    throw "StartAtSequence $StartAtSequence exceeds the selected workbook count of $($workbooks.Count)."
}

$sequencePlan = if ($null -ne $OnlySequences -and $OnlySequences.Count -gt 0) {
    $requestedSequences = @($OnlySequences | Sort-Object -Unique)
    $invalidSequences = @($requestedSequences | Where-Object { $_ -lt 1 -or $_ -gt $workbooks.Count })
    if ($invalidSequences.Count -gt 0) {
        throw "OnlySequences contains values outside 1-$($workbooks.Count): $($invalidSequences -join ', ')"
    }
    $requestedSequences
} else {
    @($StartAtSequence..$workbooks.Count)
}

$excelProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($excelProcesses.Count -gt 0) {
    throw "Excel is already running (PID(s): $($excelProcesses.Id -join ', ')). Close Excel before running this bulk update."
}

$lockFiles = @(Get-ChildItem -LiteralPath $resolvedClientRoot -Recurse -File -Filter '~$*' -ErrorAction SilentlyContinue)
if ($lockFiles.Count -gt 0) {
    throw "Excel lock files exist under the client root. Close the associated workbooks before running this bulk update: $($lockFiles.FullName -join '; ')"
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logDirectory = Join-Path $repoRoot "outputs/automatic-calculation-$timestamp"
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$csvPath = Join-Path $logDirectory 'workbook-results.csv'
$summaryPath = Join-Path $logDirectory 'summary.json'
$selectionPath = Join-Path $logDirectory 'selected-workbooks.txt'
$workbooks.FullName | ForEach-Object { [IO.Path]::GetRelativePath($repoRoot, $_) } | Set-Content -LiteralPath $selectionPath -Encoding utf8

function Release-ComObject {
    param([object] $ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Wait-ForCalculation {
    param(
        [Parameter(Mandatory)] [object] $Excel,
        [Parameter(Mandatory)] [int] $TimeoutMinutes
    )

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    while ([int] (Invoke-ComRetry { $Excel.CalculationState }) -ne 0) {
        if ($stopwatch.Elapsed.TotalMinutes -ge $TimeoutMinutes) {
            throw "Excel calculation did not finish within $TimeoutMinutes minute(s)."
        }
        Start-Sleep -Milliseconds 500
    }
}

function Get-StableFileHash {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [ValidateRange(1, 100)] [int] $Attempts = 40,
        [ValidateRange(10, 5000)] [int] $DelayMilliseconds = 250
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        }
        catch {
            if ($attempt -eq $Attempts) { throw }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Invoke-ComRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $Action,
        [ValidateRange(1, 300)] [int] $Attempts = 120,
        [ValidateRange(10, 5000)] [int] $DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return & $Action
        }
        catch {
            $isCallRejected = $_.Exception.HResult -eq -2147418111 -or $_.Exception.Message -match 'RPC_E_CALL_REJECTED|Call was rejected by callee'
            if (-not $isCallRejected -or $attempt -eq $Attempts) { throw }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

$xlCalculationAutomatic = -4105
$results = [Collections.Generic.List[object]]::new()
$retryMode = $null -ne $OnlySequences -and $OnlySequences.Count -gt 0
$expectedPriorCount = if ($retryMode) { $workbooks.Count } else { $StartAtSequence - 1 }
if ($expectedPriorCount -gt 0) {
    if ([string]::IsNullOrWhiteSpace($PriorResultsCsv)) {
        throw 'PriorResultsCsv is required when StartAtSequence is greater than 1.'
    }
    $resolvedPriorResultsCsv = [IO.Path]::GetFullPath((Join-Path $repoRoot $PriorResultsCsv))
    if (-not $resolvedPriorResultsCsv.StartsWith($resolvedRepoRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "PriorResultsCsv must resolve inside the repository: $PriorResultsCsv"
    }
    $priorResults = @(Import-Csv -LiteralPath $resolvedPriorResultsCsv)
    if ($priorResults.Count -ne $expectedPriorCount) {
        throw "Prior results contain $($priorResults.Count) row(s); expected $expectedPriorCount."
    }
    $expectedSequences = if ($retryMode) { 1..$workbooks.Count } else { 1..$expectedPriorCount }
    $actualSequences = @($priorResults | ForEach-Object { [int] $_.Sequence })
    if (@(Compare-Object $expectedSequences $actualSequences).Count -ne 0) {
        throw "Prior results do not contain the complete sequence 1 through $expectedPriorCount."
    }
    foreach ($priorResult in $priorResults) {
        if (-not $retryMode -or [int] $priorResult.Sequence -notin $sequencePlan) {
            $results.Add($priorResult)
        }
    }
    $results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
}
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    foreach ($sequence in $sequencePlan) {
        $index = $sequence - 1
        $file = $workbooks[$index]
        $relativePath = [IO.Path]::GetRelativePath($repoRoot, $file.FullName)
        $started = Get-Date
        $beforeHash = Get-StableFileHash -Path $file.FullName
        $status = 'Succeeded'
        $message = ''
        $saved = $false

        Write-Host ("[{0}/{1}] {2}" -f ($index + 1), $workbooks.Count, $relativePath)

        try {
            # UpdateLinks=0 prevents Excel from updating external workbook links on open.
            $workbook = Invoke-ComRetry { $excel.Workbooks.Open($file.FullName, 0, $false) }
            if (Invoke-ComRetry { $workbook.ReadOnly }) {
                throw 'Workbook opened read-only.'
            }

            # Calculation mode is application-wide in Excel. Setting it after every open
            # also handles workbooks that were last saved in Manual mode.
            Invoke-ComRetry { $excel.Calculation = $xlCalculationAutomatic }
            Invoke-ComRetry { $excel.CalculateBeforeSave = $true }

            # Recalculate native worksheet formulas, including volatile CELL("filename")
            # formulas used by FilePathUrl. This does not call RefreshAll.
            Invoke-ComRetry { $excel.CalculateFull() }
            Wait-ForCalculation -Excel $excel -TimeoutMinutes $CalculationTimeoutMinutes

            Invoke-ComRetry { $workbook.Save() }
            if (-not (Invoke-ComRetry { $workbook.Saved })) {
                throw 'Excel returned from Save without marking the workbook as saved.'
            }
            $saved = $true
        }
        catch {
            $status = 'Failed'
            $message = $_.Exception.Message
        }
        finally {
            if ($null -ne $workbook) {
                try { Invoke-ComRetry { $workbook.Close($saved) } } catch {
                    if ($status -eq 'Succeeded') {
                        $status = 'Failed'
                        $message = "Close failed: $($_.Exception.Message)"
                    }
                }
                Release-ComObject $workbook
                $workbook = $null
            }
        }

        $afterHash = $null
        $hashWarning = ''
        if (Test-Path -LiteralPath $file.FullName) {
            try {
                $afterHash = Get-StableFileHash -Path $file.FullName
            }
            catch {
                $hashWarning = "Post-save hash unavailable after retry: $($_.Exception.Message)"
                $message = if ([string]::IsNullOrWhiteSpace($message)) { $hashWarning } else { "$message | $hashWarning" }
            }
        }

        $result = [pscustomobject]@{
            Sequence = $index + 1
            Workbook = $relativePath
            Status = $status
            Saved = $saved
            Changed = if ($null -eq $afterHash) { $null } else { $beforeHash -ne $afterHash }
            StartedAt = $started.ToString('o')
            DurationSeconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 2)
            BeforeSha256 = $beforeHash
            AfterSha256 = $afterHash
            Message = $message
        }
        $results.Add($result)
        $results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
    }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$results = @($results | Sort-Object { [int] $_.Sequence })
$results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
$failed = @($results | Where-Object Status -eq 'Failed')
$summary = [ordered]@{
    clientRoot = [IO.Path]::GetRelativePath($repoRoot, $resolvedClientRoot)
    selectedWorkbookCount = $workbooks.Count
    resumedAtSequence = $StartAtSequence
    processedSequences = @($sequencePlan)
    succeeded = $workbooks.Count - $failed.Count
    failed = $failed.Count
    calculationMode = 'Automatic'
    recalculation = 'Excel CalculateFull'
    powerQueryRefreshRequested = $false
    completedAt = (Get-Date).ToString('o')
    resultsCsv = [IO.Path]::GetRelativePath($repoRoot, $csvPath)
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $summaryPath -Encoding utf8

Write-Host "Results: $csvPath"
Write-Host "Succeeded: $($summary.succeeded); failed: $($summary.failed)"

if ($failed.Count -gt 0) {
    exit 1
}
