[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string] $EvidencePath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-CellText($Range, [int] $Row, [int] $Column) {
    $cell = $null
    try {
        $cell = $Range.Cells.Item($Row, $Column)
        $value = $cell.Value2
        if ($null -eq $value) { return $null }
        return [string] $value
    }
    finally {
        Release-ComObject $cell
    }
}

function Read-ListObject($Worksheet) {
    $listObjects = $null
    $listObject = $null
    $headerRange = $null
    $dataRange = $null
    try {
        $listObjects = $Worksheet.ListObjects
        if ($listObjects.Count -ne 1) {
            throw "Expected one table on worksheet '$($Worksheet.Name)', found $($listObjects.Count)."
        }
        $listObject = $listObjects.Item(1)
        $headerRange = $listObject.HeaderRowRange
        $columnCount = $headerRange.Columns.Count
        $headers = for ($column = 1; $column -le $columnCount; $column++) {
            Get-CellText $headerRange 1 $column
        }

        $dataRange = $listObject.DataBodyRange
        $rowCount = if ($null -eq $dataRange) { 0 } else { $dataRange.Rows.Count }
        $rows = [Collections.Generic.List[object]]::new()
        for ($row = 1; $row -le $rowCount; $row++) {
            $record = [ordered]@{}
            for ($column = 1; $column -le $columnCount; $column++) {
                $record[$headers[$column - 1]] = Get-CellText $dataRange $row $column
            }
            $rows.Add([pscustomobject] $record)
        }

        [pscustomobject]@{
            TableName = [string] $listObject.Name
            Headers = @($headers)
            RowCount = $rowCount
            Rows = @($rows)
        }
    }
    finally {
        Release-ComObject $dataRange
        Release-ComObject $headerRange
        Release-ComObject $listObject
        Release-ComObject $listObjects
    }
}

function Group-Rows($Rows, [string[]] $Columns) {
    if ($Rows.Count -eq 0) { return @() }
    $available = @($Columns | Where-Object { $Rows[0].PSObject.Properties.Name -contains $_ })
    if ($available.Count -eq 0) { return @() }
    @($Rows |
        Group-Object -Property $available |
        Sort-Object Count -Descending |
        ForEach-Object {
            $values = [ordered]@{}
            for ($index = 0; $index -lt $available.Count; $index++) {
                $values[$available[$index]] = $_.Values[$index]
            }
            [pscustomobject]@{
                Count = $_.Count
                Values = [pscustomobject] $values
            }
        })
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedEvidence = [IO.Path]::GetFullPath($EvidencePath)
$workbookHashBefore = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
$excel = $null
$workbook = $null
$worksheets = $null
$evidence = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
    if (-not $workbook.ReadOnly) { throw 'Inspection workbook was not opened read-only.' }
    $worksheets = $workbook.Worksheets

    $targets = [ordered]@{
        WorkerIdentityFailures_DIAGNOST = @('Check', 'Role', 'Details')
        ReconciledWorkers_CHECK = @('Check', 'Status', 'Failures', 'Details')
        ReconciledWorkers_Prepare = @('SourceRole', 'MappedRoleGroup', 'RoleMappingStatus', 'Role', 'RoleMappingIssue', 'Issue')
        WorkerRolePipeline_DIAGNOST = @('SourceRole', 'MappedRoleGroup', 'RoleMappingStatus', 'Role', 'Issue', 'Stage')
        ResDayShift = @('Role')
    }
    $sheets = [Collections.Generic.List[object]]::new()
    foreach ($sheetName in $targets.Keys) {
        $worksheet = $null
        try {
            try { $worksheet = $worksheets.Item($sheetName) }
            catch {
                $sheets.Add([pscustomobject]@{
                    Worksheet = $sheetName
                    Found = $false
                    RowCount = $null
                    Headers = @()
                    Groups = @()
                })
                continue
            }
            $table = Read-ListObject $worksheet
            $groups = Group-Rows $table.Rows $targets[$sheetName]
            $sheets.Add([pscustomobject]@{
                Worksheet = $sheetName
                Found = $true
                Table = $table.TableName
                RowCount = $table.RowCount
                Headers = $table.Headers
                Groups = $groups
            })
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    $evidence = [ordered]@{
        workbook = $resolvedWorkbook
        workbookSHA256 = $workbookHashBefore
        openedReadOnly = [bool] $workbook.ReadOnly
        sheets = @($sheets)
        inspectedAt = (Get-Date).ToString('o')
    }
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$workbookHashAfter = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
if ($workbookHashAfter -ne $workbookHashBefore) {
    throw 'Workbook changed during read-only inspection.'
}
$evidence['workbookUnchanged'] = $true
$evidenceDirectory = Split-Path -Parent $resolvedEvidence
if (-not (Test-Path -LiteralPath $evidenceDirectory)) {
    New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
}
$evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedEvidence -Encoding utf8
$evidence | ConvertTo-Json -Depth 10
