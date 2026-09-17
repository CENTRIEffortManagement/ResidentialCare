# WorkerAvailabilityRules explained

This inspection version breaks `WorkerAvailabilityRules` into commented `let ... in` stages. It keeps intermediate columns, tables and interval lists, while preserving the rules for `HasAvailability`, `AvailableIntervals`, and `ExcludedIntervals`. The output is wider than the production query so you can inspect how those rules were formed.

Source: `Capacity-ShiftAvailability.xlsx_PowerQuery.m` in this folder. The original query in that file remains unchanged.

## What the query decides

For each eligible worker, it decides whether availability is unrestricted by AVAIL records or limited to explicit AVAIL windows, and collects the available and excluded intervals within the roster horizon. It does not subtract exclusions or calculate shift hours; `WorkerShiftSegments` does that next.

- **Any valid AVAIL history sets `HasAvailability` to true**, including history outside the roster dates. If all those AVAIL windows fall outside the roster, the worker still has `HasAvailability = true` but an empty `AvailableIntervals` list.
- **No AVAIL history means `HasAvailability = false`.** Downstream, the baseline is then the whole shift, with any exclusions subtracted.
- **Exclusions take precedence downstream.** A calendar day's distinct excluded hours reaching the configured threshold becomes a whole-day exclusion. Explicit AVAIL windows keep their original times.
- **Boundary dates are evaluated as complete calendar days first.** For example, morning leave can contribute to the full-day threshold even if the roster starts later that day. Only afterwards are intervals trimmed to the exact roster horizon.
- **The horizon is the earliest shift start through the latest shift end.** This query does not intersect intervals with each individual shift; that happens in `WorkerShiftSegments`.

## Follow the stages

1. **EligibleHistory:** Keep valid availability records matching an eligible employee and facility. Retain the nested worker match instead of deleting it.
2. **Roster boundaries:** Read the earliest shift start, latest shift end, and full-day threshold; calculate complete boundary dates.
3. **WorkerHistory:** Group records by payroll code and facility, retain all history in a nested table, and determine AVAIL mode before filtering by dates.
4. **CalendarWindows:** Retain the overlapping available and excluded records, then create lists trimmed to the complete calendar dates.
5. **AdjustedExclusions:** Apply the existing daily exclusion helper, keeping both the original calendar intervals and the adjusted result.
6. **RosterWindows:** Retain overlaps, trim them to exact roster boundaries, and union each list using the existing helper.
7. **WorkerRules:** Attach the rules to every eligible worker. Workers without valid matching history receive `false`, `{}`, and `{}`. Retain the nested calculations and rule record alongside the three output fields.

Each named stage uses `let` for its calculations and `in` to return its final table. The outer `in` returns `WorkerRules`. Top-level Applied Steps show the stages and boundary values; comments explain the smaller steps within each block. In the final table, click a value in `Intervals` to inspect the calculation columns, then `History` to inspect the underlying records.

Grouping changes the row grain, so record-level columns are retained inside nested tables rather than repeated as top-level columns. No `Table.RemoveColumns` or `Table.SelectColumns` steps are used in this inspection code. Filtering still excludes invalid or ineligible records, as the original query does.

## Copy into Power Query

Create a blank query named `WorkerAvailabilityRules_Explained` in the same workbook and paste the following code into its Advanced Editor. Keep it connection-only for inspection. It depends on `ReconciledWorkers_Eligible`, `AvailabilityRecords`, `AvailabilityShiftWindows`, `EXTRACT EffectiveShiftHrs`, `AvailabilityDailyExclusions`, and `AvailabilityUnion`.

This is a standalone `let ... in` expression. The three rule fields follow the original calculation; extra columns support inspection. Syntax validation does not establish runtime equivalence: this version has not been refreshed or compared with live Excel results.

```powerquery
// Query: WorkerAvailabilityRules_Explained
// Purpose: Explain worker availability rules while retaining intermediate columns and nested history.
// Inputs: Eligible workers, availability records, roster windows, effective shift hours and interval helpers.
// Output: Eligible worker rows with the original three rule fields and additional inspection columns.
// Notes: Inspection query; the production WorkerAvailabilityRules remains the downstream dependency.
let
    // Match valid history to eligible workers without discarding source columns.
    EligibleHistory = let
        // One key per employee/facility prevents repeated worker rows multiplying history.
        // Keep the full eligible worker rows inside the grouped table.
        WorkerKeys = Table.Group(
            ReconciledWorkers_Eligible,
            {"EmployeeID", "Facility-Abbrev"},
            {{"EligibleWorkers", each _, type table}}
        ),
        ValidRecords = Table.SelectRows(
            AvailabilityRecords, each [Issue] = null
        ),
        RegisteredRecords = Table.NestedJoin(
            ValidRecords, {"Payroll Code", "Facility-Abbrev"},
            WorkerKeys, {"EmployeeID", "Facility-Abbrev"},
            "Worker", JoinKind.Inner
        )
    in
        RegisteredRecords,

    // The global roster horizon includes all configured shifts and roles.
    HorizonStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    FullDayHours = #"EXTRACT EffectiveShiftHrs",

    // Count exclusions over complete boundary dates before trimming to shift hours.
    CalendarStart = DateTime.From(Date.From(HorizonStart)),
    // Midnight is an exclusive endpoint; do not add another day when already at midnight.
    CalendarEnd = if Time.From(HorizonEnd) = #time(0, 0, 0)
        then HorizonEnd
        else DateTime.From(Date.From(HorizonEnd)) + #duration(1, 0, 0, 0),

    // Change grain to one worker/facility, keeping all valid history in a nested table.
    WorkerHistory = let
        GroupedHistory = Table.Group(
            EligibleHistory, {"Payroll Code", "Facility-Abbrev"},
            {{"History", each _, type table}}
        ),
        // AVAIL mode uses the whole valid extraction, including dates outside the roster.
        AvailabilityMode = Table.AddColumn(
            GroupedHistory, "HasAvailability",
            each List.Contains([History][Available], true), type logical
        )
    in
        AvailabilityMode,

    // Keep record tables and interval lists separately so the date filtering is visible.
    CalendarWindows = let
        Source = WorkerHistory,
        AvailableRecords = Table.AddColumn(
            Source, "AvailableCalendarRecords",
            each Table.SelectRows(
                [History],
                (record) => record[Available] = true
                    and record[Start] < CalendarEnd
                    and record[End] > CalendarStart
            ), type table
        ),
        ExcludedRecords = Table.AddColumn(
            AvailableRecords, "ExcludedCalendarRecords",
            each Table.SelectRows(
                [History],
                (record) => record[Available] = false
                    and record[Start] < CalendarEnd
                    and record[End] > CalendarStart
            ), type table
        ),
        // Build Start/End records for the helpers; the original table columns remain above.
        AvailableCalendarIntervals = Table.AddColumn(
            ExcludedRecords, "AvailableCalendarIntervals",
            each List.Transform(
                Table.ToRecords([AvailableCalendarRecords]),
                (record) => [
                    Start = List.Max({record[Start], CalendarStart}),
                    End = List.Min({record[End], CalendarEnd})
                ]
            ), type list
        ),
        ExcludedCalendarIntervals = Table.AddColumn(
            AvailableCalendarIntervals, "ExcludedCalendarIntervals",
            each List.Transform(
                Table.ToRecords([ExcludedCalendarRecords]),
                (record) => [
                    Start = List.Max({record[Start], CalendarStart}),
                    End = List.Min({record[End], CalendarEnd})
                ]
            ), type list
        )
    in
        ExcludedCalendarIntervals,

    // Only exclusions can expand to full days; available windows retain their times.
    AdjustedExclusions = let
        Source = CalendarWindows,
        // The existing helper counts overlapping excluded hours once per calendar day.
        // At or above FullDayHours, it excludes that whole day; otherwise it keeps the times.
        DailyExclusions = Table.AddColumn(
            Source, "AdjustedExcludedIntervals",
            each AvailabilityDailyExclusions(
                [ExcludedCalendarIntervals], FullDayHours
            ), type list
        )
    in
        DailyExclusions,

    // Trim to the exact roster horizon only after applying the full-day exclusion rule.
    RosterWindows = let
        Source = AdjustedExclusions,
        // Strict inequalities treat an interval ending at roster start as non-overlapping.
        AvailableOverlaps = Table.AddColumn(
            Source, "AvailableRosterOverlaps",
            each List.Select(
                [AvailableCalendarIntervals],
                (window) => window[Start] < HorizonEnd
                    and window[End] > HorizonStart
            ), type list
        ),
        ExcludedOverlaps = Table.AddColumn(
            AvailableOverlaps, "ExcludedRosterOverlaps",
            each List.Select(
                [AdjustedExcludedIntervals],
                (window) => window[Start] < HorizonEnd
                    and window[End] > HorizonStart
            ), type list
        ),
        ClippedAvailability = Table.AddColumn(
            ExcludedOverlaps, "ClippedAvailableIntervals",
            each List.Transform(
                [AvailableRosterOverlaps],
                (window) => [
                    Start = List.Max({window[Start], HorizonStart}),
                    End = List.Min({window[End], HorizonEnd})
                ]
            ), type list
        ),
        ClippedExclusions = Table.AddColumn(
            ClippedAvailability, "ClippedExcludedIntervals",
            each List.Transform(
                [ExcludedRosterOverlaps],
                (window) => [
                    Start = List.Max({window[Start], HorizonStart}),
                    End = List.Min({window[End], HorizonEnd})
                ]
            ), type list
        ),
        // Use the same union helper as production, buffering lists before worker/shift reuse.
        UnitedAvailability = Table.AddColumn(
            ClippedExclusions, "AvailableIntervals",
            each List.Buffer(AvailabilityUnion([ClippedAvailableIntervals])), type list
        ),
        UnitedExclusions = Table.AddColumn(
            UnitedAvailability, "ExcludedIntervals",
            each List.Buffer(AvailabilityUnion([ClippedExcludedIntervals])), type list
        )
    in
        UnitedExclusions,

    // Return to the eligible worker rows, including workers with no matching valid history.
    WorkerRules = let
        JoinedWorkers = Table.NestedJoin(
            ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"},
            RosterWindows, {"Payroll Code", "Facility-Abbrev"},
            "Intervals", JoinKind.LeftOuter
        ),
        // Keep Intervals so all the grouped history and intermediate calculations are inspectable.
        WithRules = Table.AddColumn(
            JoinedWorkers, "Rule",
            each if Table.IsEmpty([Intervals]) then [
                HasAvailability = false,
                AvailableIntervals = {},
                ExcludedIntervals = {}
            ] else [
                HasAvailability = [Intervals]{0}[HasAvailability],
                AvailableIntervals = [Intervals]{0}[AvailableIntervals],
                ExcludedIntervals = [Intervals]{0}[ExcludedIntervals]
            ], type record
        ),
        // Expand a copy to preserve the original rule record as well as the three rule fields.
        RuleCopy = Table.DuplicateColumn(WithRules, "Rule", "RuleToExpand"),
        ExpandedRules = Table.ExpandRecordColumn(
            RuleCopy, "RuleToExpand",
            {"HasAvailability", "AvailableIntervals", "ExcludedIntervals"}
        )
    in
        Table.Buffer(ExpandedRules)
in
    WorkerRules
```
