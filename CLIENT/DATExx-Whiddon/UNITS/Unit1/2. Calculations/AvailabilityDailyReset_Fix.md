> Historical fix — superseded by the 12 September 2026 source. Do not paste this earlier replacement over the current queries. The daily AVAIL reset is retained; UNAVAIL/leave handling and effective-hour calculations have since changed. See [WorkerAvailabilityRules_Explained.md](WorkerAvailabilityRules_Explained.md).

# Daily availability reset: queries to paste

AVAIL restricts only the calendar dates its recorded interval overlaps. Each next day starts fully available unless it has its own AVAIL or exclusions. An interval ending at midnight does not restrict the next date; an interval recorded across midnight affects both dates for its recorded duration.

Source: `Capacity-ShiftAvailability.xlsx_PowerQuery.m` in this folder. These blocks contain the exact updated production definitions as standalone Advanced Editor expressions. They are not the wider inspection queries.

Apply all three definitions in this order in `Capacity-ShiftAvailability.xlsx`:

1. Add a blank query named **AvailabilityDailyWindows** and paste block 1. Keep this helper connection-only.
2. Replace the Advanced Editor contents of **WorkerAvailabilityRules** with block 2.
3. Replace the Advanced Editor contents of **WorkerShiftSegments** with block 3.

`AvailabilityRecords`, the existing exclusion helper, and the published output columns are unchanged. `HasAvailability` remains an informational flag; `AvailableIntervals` now contains the complete day-by-day baseline, including dates without AVAIL. Apply both replacements together so the consumer reads the new baseline consistently.

The M source and these copyable definitions have been syntax-checked. No workbook synchronization or refresh has been performed, and the real 57-to-5 headcount has not been rechecked. The synthetic regression query is `Workflows/ResidentialCare/Diagnostics/AvailabilityDailyWindows_CHECK.pq` (repo-relative); its cases include AVAIL on adjacent dates, midnight boundaries, exclusion precedence, and a synthetic 57-minus-18 headcount. Its runtime results still need evaluation in Power Query.

## 1. AvailabilityDailyWindows

```powerquery
// Query: AvailabilityDailyWindows
// Purpose: Reset the availability baseline on each calendar day before applying exclusions.
// Inputs: Valid AVAIL intervals and complete calendar boundaries, with an exclusive midnight end.
// Output: Unioned baseline intervals; a day without AVAIL is fully available.
// Notes: An interval spanning midnight affects both dates only for its recorded duration.
(intervals as list, calendarStart as datetime, calendarEnd as datetime) as list =>
    let
        // Reuse the worker's AVAIL list while checking each date in the roster horizon.
        Available = List.Buffer(intervals),
        Days = List.Generate(() => calendarStart, each _ < calendarEnd,
            each _ + #duration(1, 0, 0, 0)),
        DailyWindows = List.Transform(Days, (dayStart) =>
            let
                DayEnd = dayStart + #duration(1, 0, 0, 0),
                Overlapping = List.Select(Available,
                    each [Start] < DayEnd and [End] > dayStart),
                // Only AVAIL on this date restricts this date. Tomorrow is evaluated afresh.
                Baseline = if List.IsEmpty(Overlapping) then {[Start = dayStart, End = DayEnd]}
                    else List.Transform(Overlapping, each [
                        Start = List.Max({[Start], dayStart}), End = List.Min({[End], DayEnd})])
            in Baseline),
        // Merge touching windows across midnight so a fully available night remains one segment.
        Result = AvailabilityUnion(List.Combine(DailyWindows))
    in List.Buffer(Result)
```

## 2. WorkerAvailabilityRules

```powerquery
// Query: WorkerAvailabilityRules
// Purpose: Build eligible workers' daily availability baselines and roster-overlapping exclusions.
// Output: Worker rows with AvailableIntervals covering explicit windows and unrestricted dates.
// Notes: HasAvailability is informational for this calendar horizon; it never gates other dates.
let
    WorkerKeys = Table.Distinct(Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"})),
    ValidRecords = Table.SelectRows(AvailabilityRecords, each [Issue] = null),
    Registered = Table.NestedJoin(ValidRecords, {"Payroll Code", "Facility-Abbrev"},
        WorkerKeys, {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.Inner),
    RequiredRecords = Table.RemoveColumns(Registered, {"Worker"}),
    HorizonStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    FullDayHours = #"EXTRACT EffectiveShiftHrs",
    // Evaluate complete boundary dates before clipping to shift hours, so morning leave is counted.
    CalendarStart = DateTime.From(Date.From(HorizonStart)),
    CalendarEnd = if Time.From(HorizonEnd) = #time(0, 0, 0) then HorizonEnd
        else DateTime.From(Date.From(HorizonEnd)) + #duration(1, 0, 0, 0),
    RosterIntervals = (records as table, available as logical) as list =>
        let
            Overlapping = Table.SelectRows(records, each [Available] = available
                and [Start] < CalendarEnd and [End] > CalendarStart),
            CalendarIntervals = List.Transform(Table.ToRecords(Table.SelectColumns(Overlapping, {"Start", "End"})),
                each [Start = List.Max({[Start], CalendarStart}), End = List.Min({[End], CalendarEnd})]),
            // AVAIL restricts only its own dates; dates without AVAIL start fully available.
            // Full-day exclusion expansion is independent and still applies only to its own dates.
            Adjusted = if available then AvailabilityDailyWindows(CalendarIntervals, CalendarStart, CalendarEnd)
                else AvailabilityDailyExclusions(CalendarIntervals, FullDayHours),
            WithinRoster = List.Select(Adjusted, each [Start] < HorizonEnd and [End] > HorizonStart),
            Clipped = List.Transform(WithinRoster,
                each [Start = List.Max({[Start], HorizonStart}), End = List.Min({[End], HorizonEnd})])
        in List.Buffer(AvailabilityUnion(Clipped)),
    // Build day-specific windows before clipping to exact roster hours.
    // This flag describes explicit AVAIL in the calendar horizon; the windows carry the actual rules.
    WorkerIntervals = Table.Group(RequiredRecords, {"Payroll Code", "Facility-Abbrev"},
        {{"HasAvailability", each not Table.IsEmpty(Table.SelectRows(_, each [Available] = true
            and [Start] < CalendarEnd and [End] > CalendarStart)), type logical},
         {"AvailableIntervals", each RosterIntervals(_, true), type list},
         {"ExcludedIntervals", each RosterIntervals(_, false), type list}}),
    JoinedWorkers = Table.NestedJoin(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"},
        WorkerIntervals, {"Payroll Code", "Facility-Abbrev"}, "Intervals", JoinKind.LeftOuter),
    WithRules = Table.AddColumn(JoinedWorkers, "Rule", each
        if Table.IsEmpty([Intervals]) then [HasAvailability = false,
            AvailableIntervals = {[Start = HorizonStart, End = HorizonEnd]}, ExcludedIntervals = {}]
        else Record.SelectFields([Intervals]{0}, {"HasAvailability", "AvailableIntervals", "ExcludedIntervals"})),
    ExpandedRules = Table.ExpandRecordColumn(Table.RemoveColumns(WithRules, {"Intervals"}), "Rule",
        {"HasAvailability", "AvailableIntervals", "ExcludedIntervals"})
// Grouped interval lists are explicitly buffered before expansion to dated shifts.
in Table.Buffer(ExpandedRules)
```

## 3. WorkerShiftSegments

```powerquery
// Query: WorkerShiftSegments
// Purpose: Intersect available windows with each configured shift and subtract exclusions.
// Output: One worker/shift row with scalar hours and validation flags; no minute-level expansion.
let
    Windows = Table.Group(AvailabilityShiftWindows, {"Role"}, {{"Shifts", each _, type table}}),
    Joined = Table.NestedJoin(WorkerAvailabilityRules, {"Role"}, Windows, {"Role"}, "Windows", JoinKind.LeftOuter),
    ValidRoles = if Table.RowCount(Table.SelectRows(Joined, each Table.IsEmpty([Windows]))) = 0 then Joined
        else error "An eligible worker role has no configured roster shifts.",
    ShiftLists = Table.AddColumn(ValidRoles, "Shifts", each [Windows]{0}[Shifts], type table),
    Expanded = Table.ExpandTableColumn(Table.RemoveColumns(ShiftLists, {"Windows"}), "Shifts",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}),
    Measured = Table.AddColumn(Expanded, "Measurement", each
        let
            WindowStart = [ShiftStart], WindowEnd = [ShiftEnd],
            // WorkerAvailabilityRules already includes whole days without AVAIL, independently per date.
            Base = [AvailableIntervals],
            Overlapping = List.Select(Base, each [Start] < WindowEnd and [End] > WindowStart),
            Clipped = List.Buffer(List.Transform(Overlapping, each
                [Start = List.Max({[Start], WindowStart}), End = List.Min({[End], WindowEnd})])),
            Cuts = List.Buffer(List.Select([ExcludedIntervals], each [Start] < WindowEnd and [End] > WindowStart)),
            Segments = List.Buffer(AvailabilitySubtract(Clipped, Cuts)),
            Hours = List.Accumulate(Segments, 0, (total, segment) => total + Duration.TotalHours(segment[End] - segment[Start])),
            SegmentCount = List.Count(Segments),
            FullShift = SegmentCount = 1 and Segments{0}[Start] = WindowStart and Segments{0}[End] = WindowEnd,
            // Validate while shift-local lists are in scope; retain only scalar check results.
            SegmentsValid = List.AllTrue(List.Transform(Segments, (segment) =>
                segment[Start] < segment[End] and segment[Start] >= WindowStart and segment[End] <= WindowEnd
                and not List.AnyTrue(List.Transform(Cuts, (cut) => segment[Start] < cut[End] and segment[End] > cut[Start]))
                and List.AnyTrue(List.Transform(Clipped, (window) => segment[Start] >= window[Start] and segment[End] <= window[End])))),
            // Subtraction preserves order, so adjacent endpoint checks replace a second interval union.
            Disjoint = if SegmentCount < 2 then true else List.AllTrue(List.Transform({1..(SegmentCount - 1)},
                (i) => Segments{i - 1}[End] <= Segments{i}[Start]))
        in [AvailableHours = Hours, FullShift = FullShift, SegmentsValid = SegmentsValid, Disjoint = Disjoint]),
    ScalarRows = Table.ExpandRecordColumn(Table.SelectColumns(Measured,
        {"EmployeeID", "Facility-Abbrev", "Role", "Name", "Date", "Shift", "Week", "Day", "ShiftStart", "ShiftEnd", "Measurement"}),
        "Measurement", {"AvailableHours", "FullShift", "SegmentsValid", "Disjoint"})
// Retain zero-hour rows for checks; historical lists do not escape this stage.
in Table.Buffer(ScalarRows)
```
