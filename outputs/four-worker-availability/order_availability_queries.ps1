$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$path = Join-Path $root 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m'
$source = [IO.File]::ReadAllText($path)

function Move-Query([string]$name, [string]$before) {
    $script:source = $script:source
    $marker = [regex]::Match($script:source, '(?m)^// Query: ' + [regex]::Escape($name) + '\r?$')
    if (-not $marker.Success) { throw "Missing query: $name" }
    $body = [regex]::Match($script:source.Substring($marker.Index), '(?ms)^// Query: .*?^[^\r\n]*;\r?$')
    if (-not $body.Success) { throw "Missing query terminator: $name" }
    $block = $body.Value.TrimEnd("`r", "`n")
    $script:source = $script:source.Remove($marker.Index, $body.Length).TrimEnd("`r", "`n") + "`n"
    $anchor = $script:source.IndexOf('// Query: ' + $before)
    if ($anchor -lt 0) { throw "Missing insertion anchor: $before" }
    $script:source = $script:source.Insert($anchor, $block + "`n`n")
}

Move-Query 'AvailabilityEligibleRecords' 'WorkerAvailabilityRules'
Move-Query 'WorkerLeaveDays' 'WorkerAvailabilityRules'
Move-Query 'AvailabilityRules_TEST' 'ResDayShift_CHECK'
$source = [regex]::Replace($source, "(`n){3,}", "`n`n")
[IO.File]::WriteAllText($path, $source, [Text.UTF8Encoding]::new($false))
Write-Output 'Ordered eligible records, daily leave and known-answer tests after their inputs.'
