param([switch]$Apply)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$relativePath = 'CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx'
$targetPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $relativePath))
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$workbook = $null
foreach ($candidate in $excel.Workbooks) {
    if ([string]::Equals($candidate.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) { $workbook = $candidate; break }
}
if ($null -eq $workbook) { throw 'The exact Effort-All workbook was not found in the active Excel instance. No workbook opened or modified.' }
if ($workbook.ReadOnly) { throw 'Effort-All is read-only.' }
function Get-Filters($table) {
    $result = @()
    for ($i = 1; $i -le $table.ListColumns.Count; $i++) {
        $filter = $table.AutoFilter.Filters.Item($i)
        if (!$filter.On) { continue }
        $first = $null; $second = $null
        try { $first = $filter.Criteria1 } catch {}
        try { $second = $filter.Criteria2 } catch {}
        $result += [pscustomobject]@{ column = [string]$table.ListColumns.Item($i).Name; field = $i; operator = [int]$filter.Operator; criteria1 = $first; criteria2 = $second }
    }
    return $result
}
$targets = @(
    @{ Sheet = 'RoleAvailabilityDevelopedMATRIX'; Table = 'RoleAvailabilityDevelopedMATRIXDELTA' },
    @{ Sheet = 'EffortAllMatrixAG1_1D'; Table = 'EffortAllMatrixAG1_1D_2' }
)
$prepared = @()
foreach ($target in $targets) {
    $table = $workbook.Worksheets.Item($target.Sheet).ListObjects.Item($target.Table)
    $roleField = $table.ListColumns.Item('Role').Index
    $before = @(Get-Filters $table)
    $roleFilters = @($before | Where-Object column -eq 'Role')
    foreach ($roleFilter in $roleFilters) {
        $values = @($roleFilter.criteria1 | ForEach-Object { ([string]$_).TrimStart('=') })
        if ($values.Count -ne 1 -or $values[0] -cne 'AIN' -or $null -ne $roleFilter.criteria2) {
            throw "Role filter on $($target.Sheet) is not the previously identified AIN-only filter. No changes made."
        }
    }
    $prepared += [pscustomobject]@{ sheet = $target.Sheet; table = $table; roleField = $roleField; before = $before }
}
if (!$Apply) {
    [pscustomobject]@{ workbook = $relativePath; saved = [bool]$workbook.Saved; planned = @($prepared | ForEach-Object { [pscustomobject]@{ sheet = $_.sheet; clearField = 'Role'; filters = $_.before } }) } | ConvertTo-Json -Depth 8
    exit
}
$report = @()
foreach ($item in $prepared) {
    if ($item.table.AutoFilter.Filters.Item($item.roleField).On) { $null = $item.table.Range.AutoFilter($item.roleField) }
    if ($item.table.AutoFilter.Filters.Item($item.roleField).On) { throw "Role filter remains on $($item.sheet)." }
    $after = @(Get-Filters $item.table)
    $otherBefore = @($item.before | Where-Object column -ne 'Role') | ConvertTo-Json -Depth 6 -Compress
    $otherAfter = @($after | Where-Object column -ne 'Role') | ConvertTo-Json -Depth 6 -Compress
    if ($otherBefore -cne $otherAfter) { throw "Another filter changed on $($item.sheet)." }
    $report += [pscustomobject]@{ sheet = $item.sheet; roleFilterCleared = $true; otherFiltersPreserved = $true; before = $item.before; after = $after }
}
$workbook.Save()
if (!$workbook.Saved) { throw 'Excel did not confirm the workbook was saved.' }
$evidence = [pscustomobject]@{ workbook = $relativePath; saved = $true; refreshed = $false; results = $report }
$evidence | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $PSScriptRoot 'role-filter-removal.json') -Encoding UTF8
$evidence | ConvertTo-Json -Depth 8
