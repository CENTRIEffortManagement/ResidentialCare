$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$workbook = Join-Path $root 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx'
$expectedPath = Join-Path $PSScriptRoot 'day_shift_matrix_expected.csv'
$outputPath = Join-Path $PSScriptRoot 'day_shift_matrix_reconciled.csv'

Add-Type -AssemblyName System.IO.Compression

function Read-XmlEntry($archive, [string] $entryName) {
    $entry = $archive.GetEntry($entryName)
    if ($null -eq $entry) { throw "Missing workbook entry: $entryName" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return [xml] $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

function Read-SheetRows($archive, [int] $sheetNumber, [string[]] $strings) {
    $xml = Read-XmlEntry $archive "xl/worksheets/sheet$sheetNumber.xml"
    $headers = @{}
    $result = [Collections.Generic.List[object]]::new()
    foreach ($row in $xml.worksheet.sheetData.row) {
        $values = @{}
        foreach ($cell in $row.c) {
            $column = [regex]::Match($cell.r, '^[A-Z]+').Value
            $value = switch ($cell.t) {
                's' { $strings[[int] $cell.v] }
                'inlineStr' { $cell.is.InnerText }
                'b' { [bool] ([int] $cell.v) }
                default { if ($null -eq $cell.v) { $null } else { [string] $cell.v } }
            }
            $values[$column] = $value
        }
        if ($row.r -eq '1') { $headers = $values; continue }
        $record = @{}
        foreach ($column in $headers.Keys) { $record[$headers[$column]] = $values[$column] }
        $result.Add([pscustomobject] $record)
    }
    return $result.ToArray()
}

function ShiftKey($workerId, $date, $shift) {
    return "$workerId|$date|$shift"
}

$file = [IO.File]::Open($workbook, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
try {
    $archive = [IO.Compression.ZipArchive]::new($file, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $shared = Read-XmlEntry $archive 'xl/sharedStrings.xml'
        $strings = @($shared.sst.si | ForEach-Object { $_.InnerText })
        $segments = @(Read-SheetRows $archive 8 $strings)
        $checks = @(Read-SheetRows $archive 11 $strings)
        $published = @(Read-SheetRows $archive 18 $strings)
    }
    finally { $archive.Dispose() }
}
finally { $file.Dispose() }

$segmentByKey = @{}
foreach ($row in $segments) {
    $date = [datetime]::FromOADate([double] $row.Date).ToString('yyyy-MM-dd')
    $key = ShiftKey $row.EmployeeID $date $row.Shift
    if ($segmentByKey.ContainsKey($key)) { throw "Duplicate WorkerShiftSegments key: $key" }
    $segmentByKey[$key] = $row
}

$publishedByKey = @{}
foreach ($row in $published) {
    $date = [datetime]::FromOADate([double] $row.Date).ToString('yyyy-MM-dd')
    $key = ShiftKey $row.ID $date $row.Shift
    if ($publishedByKey.ContainsKey($key)) { throw "Duplicate ResDayShift key: $key" }
    $publishedByKey[$key] = $row
}

$expected = @(Import-Csv -LiteralPath $expectedPath)
$seen = @{}
$compared = foreach ($row in $expected) {
    $key = ShiftKey $row.EmployeeID $row.Date $row.Shift
    if ($seen.ContainsKey($key)) { throw "Duplicate expected key: $key" }
    $seen[$key] = $true
    $segment = $segmentByKey[$key]
    if ($null -eq $segment) { throw "Missing WorkerShiftSegments key: $key" }
    $actual = $publishedByKey[$key]
    $actualPresent = $null -ne $actual
    $expectedPresent = $row.ExpectedInResDayShift -eq 'True'
    $availableMatch = [math]::Abs([double] $segment.AvailableHours - [double] $row.ExpectedAvailableHours) -lt 0.000001
    $leaveMatch = [math]::Abs([double] $segment.DistinctLeaveHours - [double] $row.DailyRecognisedLeaveHours) -lt 0.000001
    $unavailMatch = [math]::Abs([double] $segment.UnavailHoursInShift - [double] $row.UNAVAILHoursInShift) -lt 0.000001
    $absenceMatch = [math]::Abs([double] $segment.CombinedAbsenceHoursInShift - [double] $row.CombinedAbsenceHoursInShift) -lt 0.000001
    $dayBlockMatch = ([bool] $segment.FullDayLeaveBlocked) -eq ($row.FullDayLeaveBlocked -eq 'True')
    $publishedHoursMatch = if ($actualPresent) {
        [math]::Abs([double] $actual.EffectiveShiftHrs - [double] $row.ExpectedEffectiveShiftHrs) -lt 0.000001
    } else { -not $expectedPresent }
    $row | Add-Member NoteProperty ActualAvailableHours $segment.AvailableHours
    $row | Add-Member NoteProperty ActualDailyRecognisedLeaveHours $segment.DistinctLeaveHours
    $row | Add-Member NoteProperty ActualUNAVAILHoursInShift $segment.UnavailHoursInShift
    $row | Add-Member NoteProperty ActualCombinedAbsenceHoursInShift $segment.CombinedAbsenceHoursInShift
    $row | Add-Member NoteProperty ActualFullDayLeaveBlocked $segment.FullDayLeaveBlocked
    $row | Add-Member NoteProperty ActualInResDayShift $actualPresent
    $row | Add-Member NoteProperty ActualEffectiveShiftHrs $(if ($actualPresent) { $actual.EffectiveShiftHrs } else { $null })
    $row | Add-Member NoteProperty MatchesRefreshedResult $(
        $availableMatch -and $leaveMatch -and $unavailMatch -and $absenceMatch -and $dayBlockMatch -and
        ($actualPresent -eq $expectedPresent) -and $publishedHoursMatch)
    $row
}

$extraSegments = @($segmentByKey.Keys | Where-Object { -not $seen.ContainsKey($_) })
$extraPublished = @($publishedByKey.Keys | Where-Object { -not $seen.ContainsKey($_) })
if ($extraSegments.Count -or $extraPublished.Count) {
    throw "Unexpected workbook rows: $($extraSegments.Count) segment keys, $($extraPublished.Count) published keys"
}

$compared | Export-Csv -LiteralPath $outputPath -NoTypeInformation
$failures = @($compared | Where-Object MatchesRefreshedResult -eq $false)
$failedChecks = @($checks | Where-Object Status -ne 'Pass')
Write-Output "Expected shifts: $($expected.Count); actual segment rows: $($segments.Count); actual ResDayShift rows: $($published.Count); mismatches: $($failures.Count); failed workbook checks: $($failedChecks.Count)"
$compared | Group-Object EmployeeID | ForEach-Object {
    $positive = @($_.Group | Where-Object ActualInResDayShift -eq $true)
    [pscustomobject]@{ EmployeeID = $_.Name; PositiveRows = $positive.Count;
        EffectiveHours = [math]::Round((($positive | Measure-Object ActualEffectiveShiftHrs -Sum).Sum), 2) }
} | Format-Table -AutoSize
if ($failures.Count) { throw "Refreshed Availability differs from the expected matrix; see $outputPath" }
if ($failedChecks.Count) { throw "ResDayShift_CHECK has $($failedChecks.Count) failed rows" }
