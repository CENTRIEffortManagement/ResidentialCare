[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string] $QueryName,

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

function Get-CellValue($Range, [int] $Row, [int] $Column) {
    $cell = $null
    try {
        $cell = $Range.Cells.Item($Row, $Column)
        return $cell.Value2
    }
    finally {
        Release-ComObject $cell
    }
}

function Get-CodePoints([string] $Value) {
    if ($null -eq $Value) { return $null }
    $points = foreach ($character in $Value.ToCharArray()) {
        'U+{0:X4}' -f [int] $character
    }
    return ($points -join ' ')
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedEvidence = [IO.Path]::GetFullPath($EvidencePath)
$hashBefore = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
$excel = $null
$workbook = $null
$worksheet = $null
$destination = $null
$listObject = $null
$queryTable = $null
$headerRange = $null
$dataRange = $null
$result = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    if ($workbook.ReadOnly) { throw 'The diagnostic workbook opened read-only.' }

    # This temporary worksheet and query table exist only in memory. The workbook is closed without saving.
    $worksheet = $workbook.Worksheets.Add()
    $worksheet.Name = '__CodexQueryInspection'
    $destination = $worksheet.Range('A1')
    $escapedQueryName = $QueryName.Replace('"', '""')
    $connectionString = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="' +
        $escapedQueryName + '";Extended Properties=""'
    $listObject = $worksheet.ListObjects.Add(0, $connectionString, $null, 1, $destination)
    $queryTable = $listObject.QueryTable
    $queryTable.CommandType = 2
    $queryTable.CommandText = @('SELECT * FROM [' + $QueryName.Replace(']', ']]') + ']')
    $queryTable.BackgroundQuery = $false
    $refreshed = $queryTable.Refresh($false)
    if (-not $refreshed) { throw "The temporary query table did not refresh '$QueryName'." }

    $headerRange = $listObject.HeaderRowRange
    $columnCount = $headerRange.Columns.Count
    $headers = for ($column = 1; $column -le $columnCount; $column++) {
        [string] (Get-CellValue $headerRange 1 $column)
    }
    $dataRange = $listObject.DataBodyRange
    $rowCount = if ($null -eq $dataRange) { 0 } else { $dataRange.Rows.Count }
    $rows = [Collections.Generic.List[object]]::new()
    for ($row = 1; $row -le $rowCount; $row++) {
        $record = [ordered]@{}
        for ($column = 1; $column -le $columnCount; $column++) {
            $value = Get-CellValue $dataRange $row $column
            $record[$headers[$column - 1]] = if ($null -eq $value) { $null } else { [string] $value }
        }
        if ($record.Contains('Roster Roles')) {
            $record['Roster Roles code points'] = Get-CodePoints ([string] $record['Roster Roles'])
        }
        $rows.Add([pscustomobject] $record)
    }
    $result = [ordered]@{
        workbook = $resolvedWorkbook
        workbookSHA256 = $hashBefore
        query = $QueryName
        rowCount = $rowCount
        headers = @($headers)
        rows = @($rows)
        inspectionWasInMemoryOnly = $true
        inspectedAt = (Get-Date).ToString('o')
    }
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($value in @($dataRange, $headerRange, $queryTable, $listObject, $destination, $worksheet, $workbook, $excel)) {
        Release-ComObject $value
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$hashAfter = (Get-FileHash -LiteralPath $resolvedWorkbook -Algorithm SHA256).Hash
if ($hashAfter -ne $hashBefore) { throw 'Workbook changed during in-memory query inspection.' }
$result['workbookUnchanged'] = $true
$evidenceDirectory = Split-Path -Parent $resolvedEvidence
if (-not (Test-Path -LiteralPath $evidenceDirectory)) {
    New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedEvidence -Encoding utf8
$result | ConvertTo-Json -Depth 8
