[CmdletBinding()]
param([string] $RunFolder = '')

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $RunFolder) {
    $RunFolder = (Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim()
}
if ([IO.Path]::IsPathRooted($RunFolder)) { throw 'RunFolder must be repository-relative.' }
$runRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $RunFolder))
$repoPrefix = $repoRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $runRoot.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase) -or
    -not (Test-Path -LiteralPath $runRoot -PathType Container)) { throw 'An existing deployment evidence directory inside the repository is required.' }
$comparisonId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$summaryPath = Join-Path $runRoot 'post-refresh-source-comparison.json'
$selection = @(
    @{ Target = 'Core'; Workbook = 'CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx';
        Source = 'Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq' },
    @{ Target = 'Shifts'; Workbook = 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx';
        Source = 'Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq' },
    @{ Target = 'Allocation'; Workbook = 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx';
        Source = 'outputs/overtime-roster-checks/source-gate/BD-AllocationByShiftAverage.before.m' }
)

# Load only the approved text parser/normalizer. The synchronization entry point
# itself is never executed: this comparison does not create an Excel COM object.
$tokens = $null; $errors = $null
$syncAst = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'), [ref] $tokens, [ref] $errors)
if ($errors.Count) { throw 'The approved source-comparison helper has syntax errors.' }
foreach ($name in @('Get-Sha256', 'Normalize-Formula', 'Read-MDefinitions')) {
    $definition = @($syncAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false) | Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Missing approved read-only comparison helper: $name" }
    . ([scriptblock]::Create($definition[0].Extent.Text))
}
function Get-RelativeEvidencePath([string] $Path) { [IO.Path]::GetRelativePath($repoRoot, $Path).Replace('\', '/') }
function Assert-NoExcelLock([string] $Path) {
    $lockPath = Join-Path (Split-Path $Path -Parent) ('~$' + [IO.Path]::GetFileName($Path))
    if (Test-Path -LiteralPath $lockPath) { throw "Target is open or its Excel lock remains: $Path" }
}

$rows = [Collections.Generic.List[object]]::new()
$summary = [ordered]@{
    Status = 'Started'; ComparisonId = $comparisonId; StartedAtUTC = [DateTime]::UtcNow.ToString('o')
    Scope = 'Approved read-only extraction and normalized per-query comparison after all three authorized refreshes'
    ExcelStarted = $false; WorkbookWrites = $false; AllWorkbooksUnchanged = $false; AllSourcesUnchanged = $false
    Results = @()
}
try {
    # Check all three successful save/close receipts before hashing or extracting
    # any workbook. Old successful evidence alone cannot authorize a changed file.
    $refreshReceipts = @{}
    foreach ($item in $selection) {
        $receiptPath = Join-Path $runRoot "$($item.Target)-refresh-receipt.json"
        $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
        if ($receipt.Target -cne $item.Target -or $receipt.Workbook -cne $item.Workbook -or
            $receipt.Status -cne 'Passed' -or $receipt.Saved -isnot [bool] -or -not $receipt.Saved -or
            $receipt.Closed -isnot [bool] -or -not $receipt.Closed -or $receipt.ExitCode -ne 0) {
            throw "Current refresh, save and close are not confirmed for $($item.Target)."
        }
        $refreshReceipts[$item.Target] = @{ Receipt = $receipt; Path = $receiptPath; Hash = Get-Sha256 $receiptPath }
    }

    $snapshots = @{}
    foreach ($item in $selection) {
        $workbookPath = Join-Path $repoRoot $item.Workbook
        $sourcePath = Join-Path $repoRoot $item.Source
        Assert-NoExcelLock $workbookPath
        $workbookHash = Get-Sha256 $workbookPath
        $sourceHash = Get-Sha256 $sourcePath
        if ($workbookHash -cne $refreshReceipts[$item.Target].Receipt.WorkbookSHA256After) {
            throw "Workbook differs from its successful refresh receipt: $($item.Target)."
        }
        if ($item.Target -in @('Core', 'Shifts')) {
            $syncPath = Join-Path $runRoot "$($item.Target)-sync-receipt.json"
            $sync = Get-Content -LiteralPath $syncPath -Raw | ConvertFrom-Json
            if ($sync.Target -cne $item.Target -or $sync.Workbook -cne $item.Workbook -or $sync.Source -cne $item.Source -or
                $sync.SourceSHA256 -cne $sourceHash -or -not $sync.ReExtractionMatches -or -not $sync.Saved -or -not $sync.Closed) {
                throw "Approved synchronized source no longer matches its receipt: $($item.Target)."
            }
        } else {
            $baselineEvidence = Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/source-gate/BD-AllocationByShiftAverage.extraction.json') -Raw | ConvertFrom-Json
            if (-not $baselineEvidence.workbookUnchanged -or $baselineEvidence.extractedMSHA256 -cne $sourceHash) {
                throw 'The approved AllocationByShiftAverage before-extraction has changed.'
            }
        }
        $snapshots[$item.Target] = @{ WorkbookHash = $workbookHash; SourceHash = $sourceHash }
    }

    foreach ($item in $selection) {
        $workbookPath = Join-Path $repoRoot $item.Workbook
        $sourcePath = Join-Path $repoRoot $item.Source
        $extractedPath = Join-Path $runRoot "$($item.Target)-after-refresh-$comparisonId.m"
        $extractionEvidencePath = Join-Path $runRoot "$($item.Target)-after-refresh-$comparisonId-extraction.json"
        Assert-NoExcelLock $workbookPath
        if ((Get-Sha256 $workbookPath) -cne $snapshots[$item.Target].WorkbookHash) { throw "Workbook changed before read-only extraction: $($item.Target)." }
        & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') `
            -WorkbookPath $workbookPath -OutputMPath $extractedPath -EvidencePath $extractionEvidencePath | Out-Null
        $extractionEvidence = Get-Content -LiteralPath $extractionEvidencePath -Raw | ConvertFrom-Json
        $workbookHashAfter = Get-Sha256 $workbookPath
        if (-not $extractionEvidence.workbookUnchanged -or
            $extractionEvidence.workbookSHA256 -cne $snapshots[$item.Target].WorkbookHash -or
            $workbookHashAfter -cne $snapshots[$item.Target].WorkbookHash) {
            throw "Workbook changed during read-only extraction: $($item.Target)."
        }
        $approved = Read-MDefinitions $sourcePath
        $actual = Read-MDefinitions $extractedPath
        $approvedNames = @($approved.Keys)
        $actualNames = @($actual.Keys)
        $missing = @($approvedNames | Where-Object { $actualNames -cnotcontains $_ })
        $extra = @($actualNames | Where-Object { $approvedNames -cnotcontains $_ })
        $different = @($approvedNames | Where-Object {
            $actualNames -ccontains $_ -and (Normalize-Formula $approved[$_]) -cne (Normalize-Formula $actual[$_])
        })
        $same = $missing.Count -eq 0 -and $extra.Count -eq 0 -and $different.Count -eq 0 -and $approved.Count -eq $actual.Count
        $rows.Add([ordered]@{
            Target = $item.Target; Workbook = $item.Workbook; ApprovedSource = $item.Source
            ApprovedSourceSHA256 = $snapshots[$item.Target].SourceHash
            RefreshReceipt = Get-RelativeEvidencePath $refreshReceipts[$item.Target].Path
            WorkbookSHA256Before = $snapshots[$item.Target].WorkbookHash; WorkbookSHA256After = $workbookHashAfter
            ExtractedSource = Get-RelativeEvidencePath $extractedPath; ExtractedSourceSHA256 = Get-Sha256 $extractedPath
            ExtractionEvidence = Get-RelativeEvidencePath $extractionEvidencePath
            ApprovedQueryCount = $approved.Count; ExtractedQueryCount = $actual.Count
            MissingQueries = $missing; ExtraQueries = $extra; DifferentFormulas = $different
            SameFormulas = $same; WorkbookUnchanged = $true
        })
        if (-not $same) { throw "Post-refresh query names or formulas differ from the approved source: $($item.Target)." }
    }
    # Repeat the complete snapshot check at the end, so an earlier target or its
    # receipt cannot change unnoticed while later targets are being extracted.
    foreach ($item in $selection) {
        Assert-NoExcelLock (Join-Path $repoRoot $item.Workbook)
        if ((Get-Sha256 (Join-Path $repoRoot $item.Workbook)) -cne $snapshots[$item.Target].WorkbookHash -or
            (Get-Sha256 (Join-Path $repoRoot $item.Source)) -cne $snapshots[$item.Target].SourceHash -or
            (Get-Sha256 $refreshReceipts[$item.Target].Path) -cne $refreshReceipts[$item.Target].Hash) {
            throw "Workbook, approved source or refresh receipt changed during final comparison: $($item.Target)."
        }
    }
    $summary.Status = 'Passed'
    $summary.AllWorkbooksUnchanged = $true
    $summary.AllSourcesUnchanged = $true
} catch {
    $summary.Status = 'Failed'
    $summary['Error'] = $_.Exception.Message
    throw
} finally {
    $summary.Results = $rows.ToArray()
    $summary['FinishedAtUTC'] = [DateTime]::UtcNow.ToString('o')
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding utf8
}
$summary | ConvertTo-Json -Depth 8
