param(
    [ValidateSet('Inventory', 'Profile')]
    [string] $Mode = 'Inventory'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$workbookPaths = @(
    'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx',
    'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx',
    'CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Worker Reconciliation.xlsx',
    'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx'
)

function Read-SavedBytes([string] $path) {
    # Read saved bytes without requesting exclusive access or interacting with Excel.
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $memory = [System.IO.MemoryStream]::new()
    try {
        $stream.CopyTo($memory)
        return ,$memory.ToArray()
    }
    finally { $memory.Dispose(); $stream.Dispose() }
}

function Get-BytesHash([byte[]] $bytes) {
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($hasher.ComputeHash($bytes)).Replace('-', '') }
    finally { $hasher.Dispose() }
}

function Read-EntryXml($archive, [string] $entryName) {
    $entry = $archive.GetEntry($entryName)
    if ($null -eq $entry) { return $null }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return [xml] $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

function Resolve-PackageTarget([string] $part, [string] $target) {
    $baseUri = [uri] ('http://package/' + $part)
    return ([uri]::new($baseUri, $target)).AbsolutePath.TrimStart('/')
}

function Get-WorkbookInventory($archive) {
    $book = Read-EntryXml $archive 'xl/workbook.xml'
    $relations = Read-EntryXml $archive 'xl/_rels/workbook.xml.rels'
    $relMap = @{}
    foreach ($relation in $relations.DocumentElement.ChildNodes) {
        $relMap[$relation.Id] = Resolve-PackageTarget 'xl/workbook.xml' $relation.Target
    }
    $sheets = foreach ($sheet in $book.SelectNodes('//*[local-name()="sheet"]')) {
        $relationId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheetPath = $relMap[$relationId]
        $sheetFolder = $sheetPath.Substring(0, $sheetPath.LastIndexOf('/'))
        $sheetFile = $sheetPath.Substring($sheetPath.LastIndexOf('/') + 1)
        $sheetRelations = Read-EntryXml $archive ($sheetFolder + '/_rels/' + $sheetFile + '.rels')
        $tables = @()
        if ($null -ne $sheetRelations) {
            $tables = @(foreach ($relation in $sheetRelations.DocumentElement.ChildNodes) {
                if ($relation.Type -notlike '*/table') { continue }
                $tablePath = Resolve-PackageTarget $sheetPath $relation.Target
                $table = (Read-EntryXml $archive $tablePath).DocumentElement
                [pscustomobject] @{
                    Name = $table.name
                    DisplayName = $table.displayName
                    Range = $table.ref
                    Columns = @($table.SelectNodes('./*[local-name()="tableColumns"]/*') | ForEach-Object { $_.name })
                    PackagePath = $tablePath
                }
            })
        }
        [pscustomobject] @{ Name = $sheet.name; PackagePath = $sheetPath; Tables = $tables }
    }
    return @($sheets)
}

function Read-SharedStrings($archive) {
    $document = Read-EntryXml $archive 'xl/sharedStrings.xml'
    if ($null -eq $document) { return ,@() }
    $strings = @(foreach ($item in $document.DocumentElement.ChildNodes) {
        $textNodes = $item.SelectNodes('.//*[local-name()="t"]')
        ($textNodes | ForEach-Object { $_.InnerText }) -join ''
    })
    return ,$strings
}

function Get-ColumnNumber([string] $letters) {
    $number = 0
    foreach ($character in $letters.ToUpperInvariant().ToCharArray()) {
        $number = $number * 26 + [int] $character - [int] [char] 'A' + 1
    }
    return $number
}

function Read-TableRows($archive, $sheet, $table, $strings) {
    $bounds = $table.Range -split ':'
    $firstMatch = [regex]::Match($bounds[0], '^([A-Z]+)([0-9]+)$')
    $lastMatch = [regex]::Match($bounds[-1], '^([A-Z]+)([0-9]+)$')
    $firstColumn = Get-ColumnNumber $firstMatch.Groups[1].Value
    $firstRow = [int] $firstMatch.Groups[2].Value
    $lastRow = [int] $lastMatch.Groups[2].Value
    $document = Read-EntryXml $archive $sheet.PackagePath
    $rows = foreach ($row in $document.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')) {
        $rowNumber = [int] $row.r
        if ($rowNumber -le $firstRow -or $rowNumber -gt $lastRow) { continue }
        $values = [ordered] @{ __Row = $rowNumber }
        foreach ($column in $table.Columns) { $values[$column] = $null }
        foreach ($cell in $row.ChildNodes) {
            if ($cell.LocalName -ne 'c') { continue }
            $letters = ([regex]::Match($cell.r, '^([A-Z]+)')).Value
            $columnIndex = (Get-ColumnNumber $letters) - $firstColumn
            if ($columnIndex -lt 0 -or $columnIndex -ge $table.Columns.Count) { continue }
            $valueNode = $cell.SelectSingleNode('./*[local-name()="v"]')
            $value = if ($null -eq $valueNode) { $null } else { $valueNode.InnerText }
            if ($cell.t -eq 's' -and $null -ne $value) { $value = $strings[[int] $value] }
            elseif ($cell.t -eq 'inlineStr') { $value = ($cell.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '' }
            elseif ($cell.t -eq 'b') { $value = $value -eq '1' }
            elseif ($cell.t -eq 'e') { $value = '#ERROR(' + $value + ')' }
            elseif ($null -ne $value -and $cell.t -ne 'str') {
                $number = 0.0
                if ([double]::TryParse($value, [System.Globalization.NumberStyles]::Float,
                        [System.Globalization.CultureInfo]::InvariantCulture, [ref] $number)) { $value = $number }
            }
            $values[$table.Columns[$columnIndex]] = $value
        }
        $populated = @($table.Columns | Where-Object { $null -ne $values[$_] -and $values[$_] -ne '' })
        if ($populated.Count -gt 0) { [pscustomobject] $values }
    }
    return ,@($rows)
}

$selectedTables = @{
    'StaffListMaster.xlsx' = @('Table_Masterlist', 'ResourceContract_CHECK', 'Misaligned_Join', 'UnitL1PathTABLE')
    'Capacity-ShiftAvailability.xlsx' = @('Availability_StaffList', 'ResDayShift', 'ReconciledWorkers_Prepare',
        'ReconciledWorkers_Eligible', 'ReconciledWorkers_CHECK', 'ResDayShift_CHECK', 'WorkerIdentityFailures_DIAGNOSTICS',
        'MultiRoles', 'WorkerShiftRoleCoverage_DIAGNOSTICS', 'IMPORT_Reconciled_Workers', 'UnitL1PathTABLE')
    'Worker Reconciliation.xlsx' = @('Employees_TABLE', 'TerminatedWorkers_DIAGNOSTICS', 'WorkerReconciliation_CHECK',
        'Employees_ALL', 'EmployeesTABLE_Prepare', 'WorkerSourceAttributes_SUMMARY', 'PathTABLE')
    'CapacityDistrib(A.1)-shifts.xlsx' = @('CapacityDiagnostics_SUMMARY', 'CapacityDiagnostics_DETAILS',
        'ResPeriodAvailabilityTABLE', 'ResPeriodAllocationTABLE', 'ResourcePeriodTABLE_empty', 'IMPORT_ResDayShift',
        'A1SpareDayChoices_Prepare', 'Demand_Prepare', 'RolePathTABLE')
}
$namePattern = 'Alexis\s+Griffin|Ainura\s+Borbueva|Deeanne\s+Tracey|Shelby\s+Bon|Majal\s+Dias'
$profiles = [System.Collections.Generic.List[object]]::new()

$inventory = foreach ($relativePath in $workbookPaths) {
    $absolutePath = Join-Path $repoRoot $relativePath
    $savedBytes = Read-SavedBytes $absolutePath
    $before = Get-BytesHash $savedBytes
    $memory = [System.IO.MemoryStream]::new($savedBytes, $false)
    $archive = [System.IO.Compression.ZipArchive]::new($memory, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $sheets = Get-WorkbookInventory $archive
        if ($Mode -eq 'Profile') {
            $strings = Read-SharedStrings $archive
            foreach ($sheet in $sheets) {
                foreach ($table in $sheet.Tables) {
                    if ($table.Name -notin $selectedTables[[System.IO.Path]::GetFileName($relativePath)]) { continue }
                    $rows = Read-TableRows $archive $sheet $table $strings
                    $targetRows = @($rows | Where-Object {
                        $row = $_
                        @($row.PSObject.Properties | Where-Object { $_.Value -is [string] -and $_.Value -match $namePattern }).Count -gt 0
                    })
                    $profile = [ordered] @{
                        Workbook = $relativePath; Sheet = $sheet.Name; Table = $table.Name; Range = $table.Range
                        RowCount = $rows.Count; Columns = $table.Columns; TargetRows = $targetRows
                    }
                    if ($table.Name -in @('UnitL1PathTABLE', 'PathTABLE', 'RolePathTABLE')) { $profile.PathRows = $rows }
                    $dateCoverage = @(foreach ($column in $table.Columns) {
                        if ($column -notin @('Date', 'ShiftDate', 'Roster Start', 'Roster End')) { continue }
                        $dateValues = @($rows | ForEach-Object { $_.$column } | Where-Object { $_ -is [double] })
                        if ($dateValues.Count -eq 0) { continue }
                        $range = $dateValues | Measure-Object -Minimum -Maximum
                        [pscustomobject] @{ Column = $column; Populated = $dateValues.Count
                            Minimum = [datetime]::FromOADate($range.Minimum).ToString('yyyy-MM-dd')
                            Maximum = [datetime]::FromOADate($range.Maximum).ToString('yyyy-MM-dd') }
                    })
                    if ($dateCoverage.Count -gt 0) { $profile.DateCoverage = $dateCoverage }
                    if ($table.Name -like '*CHECK' -or $table.Name -eq 'CapacityDiagnostics_SUMMARY') { $profile.Checks = $rows }
                    if ($table.Name -eq 'CapacityDiagnostics_DETAILS') {
                        $profile.Issues = @($rows | Group-Object Check, Reason, Action | ForEach-Object {
                            [pscustomobject] @{ Count = $_.Count; Check = $_.Group[0].Check; Reason = $_.Group[0].Reason
                                Action = $_.Group[0].Action; Resources = @($_.Group.Resource | Where-Object { $null -ne $_ } | Sort-Object -Unique).Count
                                FirstRow = $_.Group[0].__Row }
                        } | Sort-Object Count -Descending)
                    }
                    if ('SpareCalculationAllowed' -in $table.Columns) {
                        $profile.Permissions = @($rows | Group-Object SpareCalculationAllowed, SpareIssueReason | ForEach-Object {
                            [pscustomobject] @{ Count = $_.Count; Allowed = $_.Group[0].SpareCalculationAllowed; Reason = $_.Group[0].SpareIssueReason
                                Resources = @($_.Group.Resource | Sort-Object -Unique).Count }
                        })
                    }
                    if ($table.Name -eq 'A1SpareDayChoices_Prepare') {
                        $profile.DayDecisions = @($rows | Group-Object SparePlanningAllowed, SpareEligibilityReason | ForEach-Object {
                            [pscustomobject] @{ Count = $_.Count; Example = $_.Group[0] }
                        })
                    }
                    if ($table.Name -eq 'ResDayShift' -or $table.Name -eq 'IMPORT_ResDayShift') {
                        $profile.Roles = @($rows | Group-Object Role | ForEach-Object {
                            [pscustomobject] @{ Role = $_.Name; Rows = $_.Count; Names = @($_.Group.Name | Sort-Object -Unique).Count }
                        })
                    }
                    $profiles.Add([pscustomobject] $profile)
                }
            }
        }
        [pscustomobject] @{
            Workbook = $relativePath
            HashBefore = $before
            Sheets = $sheets
            ConnectionRefresh = @(foreach ($connection in (Read-EntryXml $archive 'xl/connections.xml').DocumentElement.ChildNodes) {
                [pscustomobject] @{ Name = $connection.name; RefreshWithRefreshAll = $connection.refreshWithRefreshAll
                    Background = $connection.background; SaveData = $connection.saveData }
            })
        }
    }
    finally { $archive.Dispose(); $memory.Dispose() }
}

foreach ($item in $inventory) {
    $after = Get-BytesHash (Read-SavedBytes (Join-Path $repoRoot $item.Workbook))
    $item | Add-Member -NotePropertyName HashAfter -NotePropertyValue $after
    $item | Add-Member -NotePropertyName FileUnchanged -NotePropertyValue ($item.HashBefore -eq $after)
}
$inventory | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'inventory.json') -Encoding UTF8
if ($Mode -eq 'Profile') {
    $profiles | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'profiles.json') -Encoding UTF8
    $profiles | Select-Object Table, RowCount, @{Name='TargetRows'; Expression={ $_.TargetRows.Count }} | Format-Table -AutoSize
    exit
}
$inventory | ForEach-Object {
    [pscustomobject] @{
        Workbook = $_.Workbook
        FileUnchanged = $_.FileUnchanged
        Sheets = @($_.Sheets | ForEach-Object {
            [pscustomobject] @{ Name = $_.Name; Tables = @($_.Tables | Select-Object Name, Range) }
        })
    }
} | ConvertTo-Json -Depth 8
