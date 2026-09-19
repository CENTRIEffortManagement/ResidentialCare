$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$target = Join-Path $root 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m'
$reference = Join-Path $PSScriptRoot 'availability.revised-reference.m'
$current = [IO.File]::ReadAllText($target)
$revised = [IO.File]::ReadAllText($reference)
if (-not $current.Contains('HasAvailability') -or -not $current.Contains('AvailabilityDailyExclusions')) {
    throw 'Current source no longer has the expected pre-repair rule.'
}

function Get-QueryBlock([string]$source, [string]$name) {
    $escaped = [regex]::Escape($name)
    $marker = [regex]::Match($source, '(?m)^// Query: ' + $escaped + '\r?$')
    if (-not $marker.Success) { throw "Missing query marker: $name" }
    $declaration = [regex]::Match($source.Substring($marker.Index), '(?m)^shared [^\r\n]+ =')
    if (-not $declaration.Success) { throw "Missing shared declaration: $name" }
    $afterDeclaration = $marker.Index + $declaration.Index + $declaration.Length
    $terminator = [regex]::Match($source.Substring($afterDeclaration), '(?m)^.*;\r?$')
    if (-not $terminator.Success) { throw "Missing query terminator: $name" }
    $end = $afterDeclaration + $terminator.Index + $terminator.Length
    return @{ Start = $marker.Index; End = $end; Text = $source.Substring($marker.Index, $end - $marker.Index) }
}

foreach ($name in @('AvailabilityRecords', 'WorkerAvailabilityRules', 'WorkerShiftSegments', 'ResDayShift_Calculated', 'ResDayShift_CHECK')) {
    $old = Get-QueryBlock $current $name
    $new = Get-QueryBlock $revised $name
    $newText = $new.Text.Replace('#"IMPORT Combined Availabilities"', '#"IMPORT Availability Leave Source"')
    $current = $current.Substring(0, $old.Start) + $newText + $current.Substring($old.End)
}

foreach ($name in @('AvailabilityLeaveReasons', 'AvailabilityIntervalHours', 'AvailabilityClipIntervals', 'AvailabilityShiftMeasurement', 'AvailabilityResidualHours', 'AvailabilityRecordKind', 'AvailabilityDailyWindows', 'AvailabilityLeaveDayTotals', 'AvailabilityEligibleRecords', 'WorkerLeaveDays', 'AvailabilityRules_TEST')) {
    if ($current.Contains('// Query: ' + $name)) { throw "Query already exists: $name" }
    $block = Get-QueryBlock $revised $name
    $text = $block.Text.Replace('#"IMPORT Combined Availabilities"', '#"IMPORT Availability Leave Source"')
    $anchor = $current.IndexOf('// Query: AvailabilityRecords')
    $current = $current.Substring(0, $anchor) + $text + "`n`n" + $current.Substring($anchor)
}

$oldFilter = 'or [Employee Code] = 25376' + "`r`n" + '))'
if (-not $current.Contains($oldFilter)) { $oldFilter = 'or [Employee Code] = 25376' + "`n" + '))' }
if (-not $current.Contains($oldFilter)) { throw 'Contract sample filter has changed.' }
$replacementFilter = 'or [Employee Code] = 19822' + "`n" + '))'
$current = $current.Replace($oldFilter, $replacementFilter)

$obsolete = Get-QueryBlock $current 'AvailabilityDailyExclusions'
$current = $current.Substring(0, $obsolete.Start) + $current.Substring($obsolete.End)
$current = [regex]::Replace($current, '(?m)^// Pathname: .*\r?\n', "// Pathname: CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx`n", 1)
$current = $current.Replace("`r`n", "`n")
[IO.File]::WriteAllText($target, $current, [Text.UTF8Encoding]::new($false))
Write-Output 'Updated current Capacity-ShiftAvailability Power Query source.'
