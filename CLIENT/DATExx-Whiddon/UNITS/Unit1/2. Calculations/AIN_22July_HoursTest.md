# AIN shift availability test: 22 July 2026

Updated 12 September 2026 for separate UNAVAIL and recognised leave rules. Use the revised production M source first. Keep these two diagnostic queries connection-only.

`AIN_20260722_Records` shows original reasons and typed intervals, including records needed by the final NIGHT shift. Workers without records remain visible. `LeaveRecordHoursOnDate` excludes UNAVAIL completely. Unknown reasons and invalid records must be resolved using `AvailabilityRecords_Invalid` and `AvailabilityReasons_DIAGNOSTICS`; unplaceable timestamps cannot be assigned to this date.

`AIN_20260722_HoursTest` builds all three configured shifts for every eligible BD AIN, including zero-hour results. Its expected hours use an independent endpoint calculation, not the production interval helpers. Missing/duplicate calculation rows fail. The day sum is now `SumOfLeaveRecordHours`; only `DistinctLeaveHoursOnDate`, counting overlaps once, decides the full-day block. SourceRows retains the exact evidence.

| Example at ShiftDuration = 7.6 | AM | PM | NIGHT |
| --- | ---: | ---: | ---: |
| Leave 10:00â€“15:00, AM ending 14:00 | 3.6 | 6.6 | 7.6 |
| UNAVAIL 13:00â€“15:00 | 6.6 | 6.6 | 7.6 |
| Leave 08:00â€“16:00 | 0 | 0 | 0 |
| UNAVAIL 06:00â€“10:00 plus leave 10:00â€“14:00 | 0 | 7.6 | 7.6 |

These examples assume otherwise unrestricted availability and AM 06:00â€“14:00, PM 14:00â€“22:00, NIGHT 22:00â€“06:00. The queries read actual Settings boundaries and allowance.

## 1. AIN_20260722_Records

```powerquery
// Query: AIN_20260722_Records
// Purpose: Show typed source evidence for eligible BD AIN workers on the test date and its final overnight shift.
// Notes: UNAVAIL contributes zero leave hours; invalid records remain visible and must be resolved.
let
    TestDate = #date(2026, 7, 22),
    DayStart = DateTime.From(TestDate),
    DayEnd = DayStart + #duration(1, 0, 0, 0),
    Shifts = Table.SelectRows(AvailabilityShiftWindows, each [Role] = "AIN" and [Date] = TestDate),
    LastShiftEnd = List.Max(Shifts[ShiftEnd]),
    // Include the whole final calendar date: a later AVAIL record can restrict early hours on that date.
    EvidenceEnd = if Time.From(LastShiftEnd) = #time(0, 0, 0) then LastShiftEnd
        else DateTime.From(Date.From(LastShiftEnd)) + #duration(1, 0, 0, 0),
    Workers = Table.SelectRows(ReconciledWorkers_Eligible, each [Role] = "AIN" and [#"Facility-Abbrev"] = "BD"),
    Dated = Table.SelectRows(AvailabilityRecords, each [Start] <> null and [End] <> null and [Start] < EvidenceEnd and [End] > DayStart),
    Joined = Table.NestedJoin(Workers, {"EmployeeID", "Facility-Abbrev"}, Dated,
        {"Payroll Code", "Facility-Abbrev"}, "Records", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Records", {"SourceRecord", "OriginalReason", "RecordKind", "Start", "End", "Issue"}),
    Hours = Table.AddColumn(Expanded, "RecordHoursOnTestDate", each if [SourceRecord] = null then 0 else
        List.Max({0, Duration.TotalHours(List.Min({[End], DayEnd}) - List.Max({[Start], DayStart}))}), type number),
    LeaveHours = Table.AddColumn(Hours, "LeaveRecordHoursOnDate", each
        if [RecordKind] = "Leave" and [Issue] = null then [RecordHoursOnTestDate] else 0, type number),
    UnavailHours = Table.AddColumn(LeaveHours, "UNAVAILRecordHoursOnDate", each
        if [RecordKind] = "UNAVAIL" and [Issue] = null then [RecordHoursOnTestDate] else 0, type number)
in Table.Sort(UnavailHours, {{"Name", Order.Ascending}, {"Start", Order.Ascending}})
```

## 2. AIN_20260722_HoursTest

```powerquery
// Query: AIN_20260722_HoursTest
// Purpose: Compare every BD AIN shift on 22 July with an independent endpoint calculation.
// Output: One eligible worker/role/shift row, including zero availability, with source evidence.
// Notes: Expected hours do not call the production interval or residual-hours helpers.
let
    TestDate = #date(2026, 7, 22),
    DayStart = DateTime.From(TestDate),
    DayEnd = DayStart + #duration(1, 0, 0, 0),
    Allowance = #"EXTRACT EffectiveShiftHrs",
    Shifts = Table.SelectRows(AvailabilityShiftWindows, each [Date] = TestDate and [Role] = "AIN"),
    ShiftCheck = if Table.RowCount(Shifts) = 3 and List.Count(List.Distinct(Shifts[Shift])) = 3 then Shifts
        else error "Expected exactly three configured AIN shifts on the test date.",
    Workers = Table.SelectRows(ReconciledWorkers_Eligible, each [Role] = "AIN" and [#"Facility-Abbrev"] = "BD"),
    WithShifts = Table.AddColumn(Workers, "Shifts", each ShiftCheck),
    Grid = Table.ExpandTableColumn(WithShifts, "Shifts", {"Date", "Shift", "ShiftStart", "ShiftEnd"}),
    Sources = Table.NestedJoin(Grid, {"EmployeeID", "Facility-Abbrev"}, AIN_20260722_Records,
        {"EmployeeID", "Facility-Abbrev"}, "SourceRows", JoinKind.LeftOuter),
    Actual = Table.SelectRows(ResDayShift_Calculated, each [Date] = TestDate and [Role] = "AIN"),
    Joined = Table.NestedJoin(Sources, {"EmployeeID", "Facility-Abbrev", "Role", "Date", "Shift"}, Actual,
        {"EmployeeID", "Facility-Abbrev", "Role", "Date", "Shift"}, "Actual", JoinKind.LeftOuter),
    // Split at every source endpoint and measure each covered span once, independently of AvailabilityUnion.
    Measure = (intervals as list, rangeStart as datetime, rangeEnd as datetime) as number =>
        let
            Overlaps = List.Select(intervals, each [Start] < rangeEnd and [End] > rangeStart),
            Points = List.Sort(List.Distinct({rangeStart, rangeEnd} & List.Combine(List.Transform(Overlaps,
                each {List.Max({[Start], rangeStart}), List.Min({[End], rangeEnd})})))),
            Hours = List.Transform({0..(List.Count(Points) - 2)}, (i) =>
                if List.AnyTrue(List.Transform(Overlaps, each [Start] <= Points{i} and [End] >= Points{i + 1}))
                then Duration.TotalHours(Points{i + 1} - Points{i}) else 0)
        in List.Sum(Hours),
    VerifiedMeasure = if Measure({}, DayStart, DayEnd) = 0
        and Measure({[Start = DayStart, End = DayStart + #duration(0, 4, 0, 0)],
            [Start = DayStart, End = DayStart + #duration(0, 4, 0, 0)]}, DayStart, DayEnd) = 4
        then Measure else error "Independent interval measurement failed its known-answer checks.",
    Compared = Table.AddColumn(Joined, "Comparison", (worker) =>
        let
            Raw = Table.SelectRows(worker[SourceRows], each [SourceRecord] <> null),
            Valid = Table.SelectRows(Raw, each [Issue] = null),
            Leave = Table.ToRecords(Table.SelectRows(Valid, each [RecordKind] = "Leave")),
            Unavail = Table.ToRecords(Table.SelectRows(Valid, each [RecordKind] = "UNAVAIL")),
            Avail = Table.ToRecords(Table.SelectRows(Valid, each [RecordKind] = "AVAIL")),
            LeaveDayHours = VerifiedMeasure(Leave, DayStart, DayEnd),
            FullDay = LeaveDayHours >= Allowance,
            Start = worker[ShiftStart], End = worker[ShiftEnd],
            Absences = Leave & Unavail,
            ShiftAbsenceHours = VerifiedMeasure(Absences, Start, End),
            CalendarDays = List.Generate(() => DateTime.From(Date.From(Start)), each _ < End, each _ + #duration(1, 0, 0, 0)),
            // AVAIL is restricted only on dates on which an explicit AVAIL interval exists.
            Baselines = List.Combine(List.Transform(CalendarDays, (d) =>
                let Next = d + #duration(1, 0, 0, 0), OnDay = List.Select(Avail, each [Start] < Next and [End] > d)
                in if List.IsEmpty(OnDay) then {[Start = d, End = Next]} else List.Transform(OnDay,
                    each [Start = List.Max({[Start], d}), End = List.Min({[End], Next})]))),
            Relevant = List.Select(Baselines & Absences, each [Start] < End and [End] > Start),
            Points = List.Sort(List.Distinct({Start, End} & List.Combine(List.Transform(Relevant,
                each {List.Max({[Start], Start}), List.Min({[End], End})})))),
            Remaining = if FullDay then 0 else List.Sum(List.Transform({0..(List.Count(Points) - 2)}, (i) =>
                if List.AnyTrue(List.Transform(Baselines, each [Start] <= Points{i} and [End] >= Points{i + 1}))
                    and not List.AnyTrue(List.Transform(Absences, each [Start] < Points{i + 1} and [End] > Points{i}))
                then Duration.TotalHours(Points{i + 1} - Points{i}) else 0)),
            FullShift = Remaining = Duration.TotalHours(End - Start),
            Expected = if FullDay or Remaining <= 0 then 0 else if FullShift then Allowance
                else List.Max({0, List.Min({Remaining, Allowance - ShiftAbsenceHours})}),
            Count = Table.RowCount(worker[Actual]),
            Production = if Count = 1 then worker[Actual]{0} else null,
            ActualHours = if Production = null then null else Production[EffectiveShiftHrs],
            InvalidCount = Table.RowCount(Raw) - Table.RowCount(Valid),
            Matches = if Production = null then false else
                Number.Abs(Expected - ActualHours) < 0.00000001
                and Number.Abs(LeaveDayHours - Production[DistinctLeaveHours]) < 0.00000001
                and Number.Abs(ShiftAbsenceHours - Production[CombinedAbsenceHoursInShift]) < 0.00000001
                and FullDay = Production[FullDayLeaveBlocked]
        in [SumOfLeaveRecordHours = List.Sum(worker[SourceRows][LeaveRecordHoursOnDate]),
            DistinctLeaveHoursOnDate = LeaveDayHours, ShiftDuration = Allowance, FullDayLeaveBlocked = FullDay,
            LeaveHoursInShift = VerifiedMeasure(Leave, Start, End), UNAVAILHoursInShift = VerifiedMeasure(Unavail, Start, End),
            CombinedDistinctAbsenceHours = ShiftAbsenceHours, ResidualClockHours = Remaining,
            ExpectedEffectiveShiftHrs = Expected, ActualEffectiveShiftHrs = ActualHours,
            AvailabilityDecision = if Production = null then "Missing or duplicate calculated row" else Production[AvailabilityDecision],
            CalculationRows = Count, InvalidEvidenceRows = InvalidCount,
            Check = if InvalidCount > 0 then "FAIL: invalid source evidence" else if Matches then "PASS" else "FAIL: expected/actual mismatch"]),
    Readable = Table.ExpandRecordColumn(Compared, "Comparison", {"SumOfLeaveRecordHours", "DistinctLeaveHoursOnDate", "ShiftDuration",
        "FullDayLeaveBlocked", "LeaveHoursInShift", "UNAVAILHoursInShift", "CombinedDistinctAbsenceHours", "ResidualClockHours",
        "ExpectedEffectiveShiftHrs", "ActualEffectiveShiftHrs", "AvailabilityDecision", "CalculationRows", "InvalidEvidenceRows", "Check"}),
    Output = Table.RemoveColumns(Readable, {"Actual"})
in Table.Sort(Output, {{"Name", Order.Ascending}, {"ShiftStart", Order.Ascending}})
```

The source includes `AvailabilityRules_TEST` with synthetic known-answer cases. No workbook was opened or refreshed during this change; PASS results for the actual workers require evaluation in Power Query.

