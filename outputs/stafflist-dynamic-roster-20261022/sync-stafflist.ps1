[CmdletBinding()]
param([switch] $CheckOnly)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$mechanismPath = Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'

# Reuse the approved mechanism's M parser, normalization and COM cleanup.
# Excel's Queries.Count returns null on this installation; named Item lookup works.
$tokens = $null; $errors = $null
$mechanism = [Management.Automation.Language.Parser]::ParseFile($mechanismPath, [ref] $tokens, [ref] $errors)
if ($errors.Count) { throw 'Approved sync mechanism did not parse.' }
foreach ($name in @('Get-Sha256', 'Normalize-Formula', 'Read-MDefinitions', 'Release-ComObject')) {
    $definition = @($mechanism.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false) |
        Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Missing approved helper: $name" }
    . ([scriptblock]::Create($definition[0].Extent.Text))
}

function Set-VerifiedExistingQueries {
    param($Queries, $Before, $Approved)
    if ($Before.Count -ne $Approved.Count -or @($Before.Keys | Where-Object { -not $Approved.Contains($_) }).Count) {
        throw 'The verified workbook and approved source have different query names.'
    }
    # Check every native formula before changing any definition. Never add a query
    # or infer that the native collection is empty from a missing Count value.
    foreach ($name in $Before.Keys) {
        $query = $null
        try {
            $query = $Queries.Item($name)
            if ($null -eq $query -or [string] $query.Name -cne $name -or
                (Normalize-Formula ([string] $query.Formula)) -cne (Normalize-Formula $Before[$name])) {
                throw "Native query '$name' does not match the verified workbook extraction."
            }
        } finally { Release-ComObject $query }
    }
    $updated = [Collections.Generic.List[string]]::new()
    foreach ($name in $Approved.Keys) {
        if ((Normalize-Formula $Before[$name]) -ceq (Normalize-Formula $Approved[$name])) { continue }
        $query = $null
        try {
            $query = $Queries.Item($name)
            $query.Formula = [string] $Approved[$name]
            if ((Normalize-Formula ([string] $query.Formula)) -cne (Normalize-Formula $Approved[$name])) {
                throw "Native query '$name' did not accept the approved formula."
            }
            $updated.Add($name)
        } finally { Release-ComObject $query }
    }
    return ,$updated.ToArray()
}

# Exercise the missing-Count case and a changed-query rejection without Excel.
$beforeFixture = [ordered]@{ Existing = 'let Value = 1 in Value'; Unrelated = '42' }
$approvedFixture = [ordered]@{ Existing = 'let Value = 2 in Value'; Unrelated = '42' }
$mockQueries = [pscustomobject]@{ ByName = @{
    Existing = [pscustomobject]@{Name='Existing'; Formula=$beforeFixture.Existing}
    Unrelated = [pscustomobject]@{Name='Unrelated'; Formula='42'}
}}
$mockQueries | Add-Member ScriptMethod Item { param($name) $this.ByName[$name] }
$updatedFixture = @(Set-VerifiedExistingQueries $mockQueries $beforeFixture $approvedFixture)
if ($updatedFixture.Count -ne 1 -or $updatedFixture[0] -ne 'Existing' -or
    $mockQueries.ByName.Existing.Formula -cne $approvedFixture.Existing -or $mockQueries.ByName.Unrelated.Formula -cne '42') {
    throw 'Named-query synchronization fixture failed.'
}
$mockQueries.ByName.Existing.Formula = $beforeFixture.Existing
$mockQueries.ByName.Unrelated.Formula = 'changed externally'
$rejected = $false
try { $null = Set-VerifiedExistingQueries $mockQueries $beforeFixture $approvedFixture }
catch { $rejected = $_.Exception.Message -like '*does not match the verified workbook extraction*' }
if (-not $rejected -or $mockQueries.ByName.Existing.Formula -cne $beforeFixture.Existing) {
    throw 'Changed-query fixture did not block mutation.'
}
if ($CheckOnly) { 'PASS: named-query synchronization without Count; unrelated definitions preserved; changed native formula blocks all updates. No Excel opened.'; return }

$boundaries = @(Get-Content -LiteralPath (Join-Path $PSScriptRoot 'source-boundary-validation.json') -Raw | ConvertFrom-Json)
$nativeTests = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'native-tests/summary.json') -Raw | ConvertFrom-Json
$targets = @()
foreach ($unitId in @('BD', 'TE', 'JH-RY')) {
    $boundary = @($boundaries | Where-Object Unit -eq $unitId)
    if ($boundary.Count -ne 1) { throw "Missing source boundary evidence: $unitId" }
    $expectedWorkbook = "CLIENT/DATExx-Whiddon/UNITS/$unitId/2. Calculations/StaffListMaster.xlsx"
    $expectedSource = $expectedWorkbook + '_PowerQuery.m'
    if ($boundary[0].Workbook -cne $expectedWorkbook -or $boundary[0].Source -cne $expectedSource) { throw 'Target pair is outside the approved scope.' }
    $workbookPath = Join-Path $repoRoot $expectedWorkbook
    $sourcePath = Join-Path $repoRoot $expectedSource
    if ((Get-Sha256 $workbookPath) -ne $boundary[0].WorkbookSHA256) { throw "Workbook changed since verification: $unitId" }
    $sourceHash = Get-Sha256 $sourcePath
    if ($nativeTests.Status -ne 'Passed' -or $sourceHash -ne $boundary[0].SourceSHA256 -or
        $sourceHash -ne $nativeTests.SourceSHA256.$unitId) { throw "Validated source changed: $unitId" }
    if (Test-Path -LiteralPath (Join-Path (Split-Path $workbookPath -Parent) '~$StaffListMaster.xlsx')) { throw "Workbook is open: $unitId" }
    $before = Read-MDefinitions (Join-Path $PSScriptRoot "$unitId.before.embedded.m")
    $approved = Read-MDefinitions $sourcePath
    $changed = @($approved.Keys | Where-Object {
        -not $before.Contains($_) -or (Normalize-Formula $before[$_]) -cne (Normalize-Formula $approved[$_])
    })
    if ($before.Count -ne 27 -or $approved.Count -ne 27 -or $changed.Count -ne 1 -or
        $changed[0] -cne 'EXTRACT StaffList RosterDates') { throw "Unexpected query changes: $unitId" }
    $backup = Join-Path $PSScriptRoot "backups/$unitId.StaffListMaster.before-sync.xlsx"
    [void][IO.Directory]::CreateDirectory((Split-Path $backup -Parent))
    if (-not (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $workbookPath -Destination $backup }
    if ((Get-Sha256 $backup) -ne $boundary[0].WorkbookSHA256) { throw "Backup hash does not match: $unitId" }
    $targets += [pscustomobject]@{Unit=$unitId;Workbook=$workbookPath;Source=$sourcePath;SourceHash=$sourceHash;
        BeforeHash=$boundary[0].WorkbookSHA256;Before=$before;Approved=$approved;Backup=$backup}
}

$excel = $null
$receipts = [Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    foreach ($target in $targets) {
        $workbook = $null; $queries = $null
        Write-Host "Syncing $($target.Unit) StaffListMaster."
        try {
            if ((Get-Sha256 $target.Workbook) -ne $target.BeforeHash -or
                (Get-Sha256 $target.Source) -ne $target.SourceHash -or
                (Test-Path -LiteralPath (Join-Path (Split-Path $target.Workbook -Parent) '~$StaffListMaster.xlsx'))) {
                throw "Target changed or opened after preflight: $($target.Unit)"
            }
            $workbook = $excel.Workbooks.Open($target.Workbook, 0, $false)
            if ($workbook.ReadOnly -or -not [string]::Equals([string] $workbook.FullName, $target.Workbook, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Excel opened a read-only or incorrect target: $($target.Unit)"
            }
            $queries = $workbook.Queries
            $updated = @(Set-VerifiedExistingQueries $queries $target.Before $target.Approved)
            $workbook.Save()
            if (-not $workbook.Saved) { throw "Excel did not confirm save: $($target.Unit)" }
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            Release-ComObject $queries; Release-ComObject $workbook
        }
        $extractedPath = Join-Path $PSScriptRoot "$($target.Unit).after.embedded.m"
        & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $target.Workbook `
            -OutputMPath $extractedPath -EvidencePath (Join-Path $PSScriptRoot "$($target.Unit).after.extraction.json") | Out-Null
        $actual = Read-MDefinitions $extractedPath
        if ($actual.Count -ne $target.Approved.Count -or @($actual.Keys | Where-Object { -not $target.Approved.Contains($_) }).Count) {
            throw "Saved query names differ: $($target.Unit)"
        }
        foreach ($name in $target.Approved.Keys) {
            if ((Normalize-Formula $actual[$name]) -cne (Normalize-Formula $target.Approved[$name])) {
                throw "Re-extracted query does not match the approved source: $($target.Unit)/$name"
            }
        }
        if (Test-Path -LiteralPath (Join-Path (Split-Path $target.Workbook -Parent) '~$StaffListMaster.xlsx')) { throw "Workbook remains open: $($target.Unit)" }
        $receipt = [ordered]@{Unit=$target.Unit;Workbook=$target.Workbook;Source=$target.Source;SourceSHA256=$target.SourceHash;
            WorkbookSHA256Before=$target.BeforeHash;WorkbookSHA256After=(Get-Sha256 $target.Workbook);Backup=$target.Backup;
            UpdatedQueries=$updated;VerifiedQueries=$actual.Count;ReExtractionMatches=$true;Saved=$true;Closed=$true;
            RefreshInvoked=$false;RecordedAt=(Get-Date).ToString('o')}
        $receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot "$($target.Unit).sync.json") -Encoding utf8
        $receipts.Add($receipt)
        Write-Host "Verified $($target.Unit): $($actual.Count) saved query definitions match; workbook closed."
    }
} finally {
    if ($null -ne $excel) { $excel.Quit() }
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$receipts | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'sync-summary.json') -Encoding utf8
'PASS: all three approved StaffListMaster workbooks synchronized, saved, closed and re-extracted; all query definitions verified.'
