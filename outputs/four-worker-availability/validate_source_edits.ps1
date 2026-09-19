$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$files = @(
    'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m',
    'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m'
)
foreach ($relative in $files) {
    $path = Join-Path $root $relative
    $source = [IO.File]::ReadAllText($path)
    $names = @([regex]::Matches($source, '(?m)^shared\s+(?<name>#[^\r\n]+?|[^\s=]+)\s*=') | ForEach-Object { $_.Groups['name'].Value })
    $duplicates = @($names | Group-Object | Where-Object Count -gt 1)
    if ($duplicates.Count) { throw "Duplicate shared definitions in ${relative}: $($duplicates.Name -join ', ')" }
    if ($source -match '(?m)^(<<<<<<<|=======|>>>>>>>)') { throw "Conflict markers in $relative" }
    $stack = [Collections.Generic.List[char]]::new()
    $inString = $false
    $inComment = $false
    for ($i = 0; $i -lt $source.Length; $i++) {
        $char = $source[$i]
        $next = if ($i + 1 -lt $source.Length) { $source[$i + 1] } else { [char]0 }
        if ($inComment) {
            if ($char -eq "`n") { $inComment = $false }
            continue
        }
        if ($inString) {
            if ($char -eq '"' -and $next -eq '"') { $i++; continue }
            if ($char -eq '"') { $inString = $false }
            continue
        }
        if ($char -eq '/' -and $next -eq '/') { $inComment = $true; $i++; continue }
        if ($char -eq '"') { $inString = $true; continue }
        if ($char -in @('(', '[', '{')) { $stack.Add($char); continue }
        if ($char -in @(')', ']', '}')) {
            if ($stack.Count -eq 0) { throw "Unmatched delimiter in $relative at character $i" }
            $open = $stack[$stack.Count - 1]
            $expected = switch ($char) { ')' {'('} ']' {'['} '}' {'{'} }
            if ($open -ne $expected) { throw "Mismatched delimiter in $relative at character $i" }
            $stack.RemoveAt($stack.Count - 1)
        }
    }
    if ($inString -or $stack.Count) { throw "Unclosed string or delimiter in $relative" }
    Write-Output "$relative : $($names.Count) unique shared definitions; lexical delimiters balanced"
}
$availability = [IO.File]::ReadAllText((Join-Path $root $files[0]))
$b = [IO.File]::ReadAllText((Join-Path $root $files[1]))
foreach ($required in @('AvailabilityRecordKind', 'AvailabilityDailyWindows', 'WorkerLeaveDays', 'AvailabilityRules_TEST', 'FullDayLeaveBlocked')) {
    if (-not $availability.Contains($required)) { throw "Missing availability rule: $required" }
}
if ($availability.Contains('AvailabilityDailyExclusions') -or $availability.Contains('HasAvailability')) {
    throw 'Withdrawn whole-day or worker-wide rule remains in Availability.'
}
if ($b -notmatch 'and Table.IsEmpty\(AvailabilityLineageFailures\)') { throw 'B output does not gate availability lineage.' }
Write-Output 'Required rule and B publication-gate checks passed.'
