# Why are only five AINs published?

`ResDayShift` keeps calculated rows with `EffectiveShiftHrs > 0`. This diagnostic counts the workers by the reason they have positive hours, zero hours, or no calculated shift. Each result has a nested `Workers` table containing the identities and the evidence.

## Use this query

1. Create a blank query called **AIN_22July_Diagnosis** in the same workbook.
2. Paste the code below into Advanced Editor and inspect its preview. Keep it connection-only.
3. Read the **Outcome** and **WorkerCount** columns. Those counts identify where the workers were lost.
4. Click a **Workers** table to see employee IDs, names, raw exclusion counts, excluded daily hours, the full-day threshold, and calculated hours. Click **ExcludedRecordsOnDate** for the individual parsed records and their `SourceRecord` identifiers.

The year and shift end come from the configured AIN shift starting at 06:00 on 22 July. If Settings has multiple matching years, the query stops with the candidates instead of choosing one silently.

## What to look for

- **Whole-day exclusion threshold reached:** The current exclusion helper converts that date to a full-day exclusion when distinct excluded hours reach `EXTRACT EffectiveShiftHrs`. This includes UNAVAIL and leave. For example, eight hours of afternoon UNAVAIL can block the morning too when the threshold is 7.6 hours. This is a verified code rule, not yet a verified explanation of your worker counts.
- **No baseline despite no AVAIL today:** The evaluated worker baseline does not match the corrected daily rule. Check that the new helper and both replacement queries are installed together.
- **Today's AVAIL does not cover this shift:** There is an explicit AVAIL record on that date, but none of its windows overlap the shift.
- **Excluded by worker eligibility:** These workers were removed before hours were calculated. The nested detail retains `EligibilityIssue`.
- **Positive hours:** These workers should pass the positive-hours filter in `ResDayShift`, provided its validation checks pass. If this count exceeds the displayed five, compare the same date, year, role, facility and shift in the current query preview.

The summary starts with reconciled BD workers mapped to AIN, before eligibility filtering. Its total is not assumed to be 57. If it differs, first reconcile where the 57 was counted. **FalseRowsStartingAtSix counts records, not distinct workers.** Exclusions that began earlier can overlap the shift, and exclusions at other times can trigger the full-day rule.

This diagnostic does not alter any availability policy or production query. It reads existing query results, retains invalid history as evidence, and does not turn a missing calculation into zero hours. Syntax is checked; the actual summary still needs evaluation in your Power Query session. It does not synchronize or refresh workbooks automatically.

```powerquery
// Query: AIN_22July_Diagnosis
// Purpose: Explain each reconciled BD AIN's result for the configured 22 July 06:00 shift.
// Inputs: Existing preparation, parsed records, worker rules, shift calendar and calculated shift queries.
// Output: Counts by outcome; each Workers table retains identities, eligibility failures and zero-hour evidence.
// Notes: Read-only inspection. Does not change production rules or depend on the filtered ResDayShift output.
let
    // Infer the year from Settings, but stop if more than one matching shift exists.
    CandidateShifts = Table.SelectRows(AvailabilityShiftWindows,
        each [Role] = "AIN" and Date.Month([Date]) = 7 and Date.Day([Date]) = 22
            and Time.From([ShiftStart]) = #time(6, 0, 0)),
    Target = if Table.RowCount(CandidateShifts) = 1 then CandidateShifts{0}
        else error Error.Record("Select diagnostic shift",
            "Expected exactly one AIN shift on 22 July starting at 06:00. Inspect CandidateShifts and specify the required year if necessary.",
            CandidateShifts),
    ShiftStart = Target[ShiftStart],
    ShiftEnd = Target[ShiftEnd],
    DayStart = DateTime.From(Date.From(ShiftStart)),
    DayEnd = DayStart + #duration(1, 0, 0, 0),
    Threshold = #"EXTRACT EffectiveShiftHrs",

    // Start before eligibility filtering so excluded identities do not disappear from the diagnosis.
    Workers = Table.SelectRows(ReconciledWorkers_Prepare,
        each [Role] = "AIN" and [#"Facility-Abbrev"] = "BD"),
    NamedIssues = Table.RenameColumns(Workers, {{"Issue", "EligibilityIssue"}}),
    WithHistory = Table.NestedJoin(NamedIssues, {"EmployeeID", "Facility-Abbrev"},
        AvailabilityRecords, {"Payroll Code", "Facility-Abbrev"}, "History", JoinKind.LeftOuter),
    WithRules = Table.NestedJoin(WithHistory, {"EmployeeID", "Facility-Abbrev", "Role"},
        WorkerAvailabilityRules, {"EmployeeID", "Facility-Abbrev", "Role"}, "Rules", JoinKind.LeftOuter),
    CalculatedShift = Table.SelectRows(ResDayShift_Calculated,
        each [Role] = "AIN" and [Date] = Target[Date] and [Shift] = Target[Shift]
            and [ShiftStart] = ShiftStart and [ShiftEnd] = ShiftEnd),
    WithCalculation = Table.NestedJoin(WithRules, {"EmployeeID", "Facility-Abbrev", "Role"},
        CalculatedShift, {"EmployeeID", "Facility-Abbrev", "Role"}, "Calculated", JoinKind.LeftOuter),

    // Compare raw dated records with the actual baseline, adjusted exclusions and calculated hours.
    WithEvidence = Table.AddColumn(WithCalculation, "Evidence", each
        let
            Valid = Table.SelectRows([History], each [Issue] = null),
            DayRecords = Table.SelectRows(Valid, each [Start] < DayEnd and [End] > DayStart),
            AvailOnDay = Table.SelectRows(DayRecords, each [Available] = true),
            ExcludedOnDay = Table.SelectRows(DayRecords, each [Available] = false),
            ExactStartExclusions = Table.SelectRows(ExcludedOnDay, each [Start] = ShiftStart),
            OverlappingExclusions = Table.SelectRows(ExcludedOnDay,
                each [Start] < ShiftEnd and [End] > ShiftStart),
            RawDayIntervals = AvailabilityUnion(List.Transform(Table.ToRecords(ExcludedOnDay),
                each [Start = List.Max({[Start], DayStart}), End = List.Min({[End], DayEnd})])),
            // Count overlapping records once when measuring the full-day threshold.
            ExcludedDayHours = List.Accumulate(RawDayIntervals, 0,
                (total, window) => total + Duration.TotalHours(window[End] - window[Start])),
            RuleCount = Table.RowCount([Rules]),
            Baseline = if RuleCount = 1 then [Rules]{0}[AvailableIntervals] else {},
            BaselineOverlaps = List.Select(Baseline, each [Start] < ShiftEnd and [End] > ShiftStart),
            BaselineHours = if RuleCount <> 1 then null else List.Accumulate(BaselineOverlaps, 0,
                (total, window) => total + Duration.TotalHours(
                    List.Min({window[End], ShiftEnd}) - List.Max({window[Start], ShiftStart}))),
            AdjustedExclusions = if RuleCount = 1 then [Rules]{0}[ExcludedIntervals] else {},
            AdjustedShiftExclusions = List.Select(AdjustedExclusions,
                each [Start] < ShiftEnd and [End] > ShiftStart),
            CalculationCount = Table.RowCount([Calculated]),
            Hours = if CalculationCount = 1 then [Calculated]{0}[AvailableHours] else null,
            EffectiveHours = if CalculationCount = 1 then [Calculated]{0}[EffectiveShiftHrs] else null,
            HasHistoryAvail = List.Contains(Valid[Available], true),
            Outcome = if [EligibilityIssue] <> null then "Excluded by worker eligibility"
                else if RuleCount <> 1 then "Missing or duplicate worker rules"
                else if CalculationCount <> 1 then "Missing or duplicate calculated shift"
                else if EffectiveHours = null then "Missing calculated hours"
                else if EffectiveHours > 0 then "Positive hours: should appear in ResDayShift"
                else if BaselineHours = 0 and Table.IsEmpty(AvailOnDay) then
                    "No baseline despite no AVAIL today: check installed daily-reset queries"
                else if BaselineHours = 0 then "Today's AVAIL does not cover this shift"
                else if Hours = 0 and ExcludedDayHours >= Threshold then
                    "Whole-day exclusion threshold reached"
                else if Hours = 0 then "Shift availability removed by overlapping exclusions"
                else "Unexpected effective hours: inspect Calculated"
        in [
            Outcome = Outcome,
            AvailableHours = Hours,
            EffectiveShiftHrs = EffectiveHours,
            BaselineHoursBeforeExclusions = BaselineHours,
            AVAILRecordsOnDate = Table.RowCount(AvailOnDay),
            HasAVAILAnywhere = HasHistoryAvail,
            FalseRowsStartingAtSix = Table.RowCount(ExactStartExclusions),
            FalseRowsOverlappingShift = Table.RowCount(OverlappingExclusions),
            FalseRowsOnDate = Table.RowCount(ExcludedOnDay),
            DistinctExcludedHoursOnDate = ExcludedDayHours,
            FullDayThreshold = Threshold,
            FullDayThresholdReached = ExcludedDayHours >= Threshold,
            InvalidHistoryRows = Table.RowCount([History]) - Table.RowCount(Valid),
            ExcludedRecordsOnDate = ExcludedOnDay,
            AdjustedShiftExclusions = AdjustedShiftExclusions,
            CalculationRows = CalculationCount
        ], type record),
    EvidenceColumns = Table.ExpandRecordColumn(WithEvidence, "Evidence",
        {"Outcome", "AvailableHours", "EffectiveShiftHrs", "BaselineHoursBeforeExclusions",
         "AVAILRecordsOnDate", "HasAVAILAnywhere", "FalseRowsStartingAtSix", "FalseRowsOverlappingShift",
         "FalseRowsOnDate", "DistinctExcludedHoursOnDate", "FullDayThreshold", "FullDayThresholdReached",
         "InvalidHistoryRows", "ExcludedRecordsOnDate", "AdjustedShiftExclusions", "CalculationRows"}),
    // Put the explanation beside the worker's identity; retain nested evidence at the right.
    ReadableOrder = Table.ReorderColumns(EvidenceColumns,
        {"EmployeeID", "BaseName", "Role", "Facility-Abbrev", "Outcome", "AvailableHours", "EffectiveShiftHrs",
         "BaselineHoursBeforeExclusions", "AVAILRecordsOnDate", "FalseRowsStartingAtSix",
         "FalseRowsOverlappingShift", "DistinctExcludedHoursOnDate", "FullDayThreshold", "FullDayThresholdReached"}),
    Dated = Table.AddColumn(ReadableOrder, "DiagnosticShiftStart", each ShiftStart, type datetime),
    WorkerDetails = Table.Sort(Dated,
        {{"Outcome", Order.Ascending}, {"BaseName", Order.Ascending}, {"EmployeeID", Order.Ascending}}),
    // Count every prepared BD AIN exactly once by its diagnosed outcome; retain rows for drill-down.
    Summary = Table.Group(WorkerDetails, {"Outcome"},
        {{"WorkerCount", each Table.RowCount(_), Int64.Type}, {"Workers", each _, type table}})
in
    Table.Sort(Summary, {{"WorkerCount", Order.Descending}, {"Outcome", Order.Ascending}})
