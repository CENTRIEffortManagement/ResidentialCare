$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$actualPath = Join-Path $folder 'day_shift_matrix.csv'
$beforePath = Join-Path $folder 'day_shift_matrix_before.csv'
if (-not (Test-Path $beforePath)) { Copy-Item -LiteralPath $actualPath -Destination $beforePath }
$rows = @(Import-Csv -LiteralPath $beforePath)
if ($rows.Count -ne 336 -or @($rows.EmployeeID | Sort-Object -Unique).Count -ne 4) {
    throw 'Expected the saved four-worker, 336-shift sample.'
}
if (@($rows | Where-Object { $_.AVAILSourceRecords }).Count -gt 0) {
    throw 'This sample calculator expects no AVAIL records; update the model before broadening the sample.'
}

function Parse-IntervalText([string]$source, [string]$kind) {
    foreach ($part in ($source -split '; ')) {
        if (-not $part) { continue }
        if ($part -notmatch '^#(?<id>\d+) (?<start>\d{4}-\d\d-\d\d \d\d:\d\d) to (?<end>\d{4}-\d\d-\d\d \d\d:\d\d)$') {
            throw "Unexpected source interval: $part"
        }
        [pscustomobject]@{
            ID = $Matches.id
            Start = [datetime]::ParseExact($Matches.start, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            End = [datetime]::ParseExact($Matches.end, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
            Kind = $kind
        }
    }
}

function Merge-Intervals($items) {
    $merged = [Collections.Generic.List[object]]::new()
    foreach ($item in @($items | Sort-Object Start,End)) {
        if ($merged.Count -gt 0 -and $item.Start -le $merged[$merged.Count - 1].End) {
            $previous = $merged[$merged.Count - 1]
            if ($item.End -gt $previous.End) { $previous.End = $item.End }
        } else {
            $merged.Add([pscustomobject]@{ Start = $item.Start; End = $item.End })
        }
    }
    return $merged.ToArray()
}

function Clip-Intervals($items, [datetime]$start, [datetime]$end) {
    $parts = foreach ($item in $items) {
        if ($item.Start -lt $end -and $item.End -gt $start) {
            [pscustomobject]@{ Start = $(if ($item.Start -gt $start) { $item.Start } else { $start });
                End = $(if ($item.End -lt $end) { $item.End } else { $end }) }
        }
    }
    return Merge-Intervals $parts
}

function Interval-Hours($items) {
    $total = 0.0
    foreach ($item in $items) { $total += ($item.End - $item.Start).TotalHours }
    return $total
}

$records = @{}
foreach ($row in $rows) {
    if (-not $records.ContainsKey($row.EmployeeID)) { $records[$row.EmployeeID] = @{} }
    foreach ($entry in @(Parse-IntervalText $row.UNAVAILSourceRecords 'UNAVAIL') + @(Parse-IntervalText $row.LeaveOtherSourceRecords 'Leave')) {
        if ($entry) { $records[$row.EmployeeID][$entry.ID] = $entry }
    }
}

$allowance = 7.6
$expected = foreach ($row in $rows) {
    $start = [datetime]::ParseExact($row.ShiftStart, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $end = [datetime]::ParseExact($row.ShiftEnd, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $date = [datetime]::ParseExact($row.Date, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $all = @($records[$row.EmployeeID].Values)
    $leave = @($all | Where-Object Kind -eq 'Leave')
    $unavail = @($all | Where-Object Kind -eq 'UNAVAIL')
    $dailyLeave = @(Clip-Intervals $leave $date $date.AddDays(1))
    $dailyLeaveHours = Interval-Hours $dailyLeave
    $blocked = $dailyLeaveHours -ge $allowance
    $shiftLeave = @(Clip-Intervals $leave $start $end)
    $shiftUnavail = @(Clip-Intervals $unavail $start $end)
    $combined = @(Merge-Intervals ($shiftLeave + $shiftUnavail))
    $remaining = if ($blocked) { 0.0 } else { ($end - $start).TotalHours - (Interval-Hours $combined) }
    $remaining = [math]::Max(0, $remaining)
    $fullShift = -not $blocked -and $combined.Count -eq 0
    $effective = if ($blocked -or $remaining -le 0) { 0.0 }
        elseif ($fullShift) { $allowance }
        else { [math]::Max(0, [math]::Min($remaining, $allowance - (Interval-Hours $combined))) }
    $decision = if ($blocked) { 'All shifts blocked: daily recognised leave reaches ShiftDuration' }
        elseif ($remaining -le 0) { 'Unavailable: recorded absences cover available time in this shift' }
        elseif ((Interval-Hours $combined) -ge $allowance) { 'Unavailable: shift absences exhaust ShiftDuration' }
        elseif ($fullShift) { 'Full shift available' }
        else { 'Residual availability' }
    [pscustomobject]@{
        EmployeeID = $row.EmployeeID; Name = $row.Name; Role = $row.Role; Date = $row.Date;
        Day = $row.Day; Shift = $row.Shift; ShiftStart = $row.ShiftStart; ShiftEnd = $row.ShiftEnd;
        AVAILSourceRecords = $row.AVAILSourceRecords; UNAVAILSourceRecords = $row.UNAVAILSourceRecords;
        LeaveSourceRecords = $row.LeaveOtherSourceRecords;
        DailyRecognisedLeaveHours = [math]::Round($dailyLeaveHours, 8);
        FullDayLeaveBlocked = $blocked;
        LeaveHoursInShift = [math]::Round((Interval-Hours $shiftLeave), 8);
        UNAVAILHoursInShift = [math]::Round((Interval-Hours $shiftUnavail), 8);
        CombinedAbsenceHoursInShift = [math]::Round((Interval-Hours $combined), 8);
        ExpectedAvailableHours = [math]::Round($remaining, 8);
        ExpectedEffectiveShiftHrs = [math]::Round($effective, 8);
        ExpectedInResDayShift = $effective -gt 0;
        SavedWorkerShiftSegmentsHours = $row.SavedWorkerShiftSegmentsHours;
        PresentInSavedResDayShift = $row.PresentInSavedResDayShift;
        AvailabilityDecision = $decision
    }
}

$output = Join-Path $folder 'day_shift_matrix_expected.csv'
$expected | Export-Csv -LiteralPath $output -NoTypeInformation -Encoding utf8
$expected | Export-Csv -LiteralPath $actualPath -NoTypeInformation -Encoding utf8
Write-Output "Rows: $($expected.Count); expected positive shifts: $(@($expected | Where-Object ExpectedInResDayShift).Count)"
Write-Output "Updated $output and $actualPath; preserved $beforePath"
