// Query: AIN_20260722_HoursTest
// Purpose: Independently check each worker's excluded hours on 22 July 2026 and the 06:00 shift.
// Inputs: AIN_20260722_Records, existing interval helpers, Settings, worker rules and calculated hours.
// Output: One row per prepared BD AIN with arithmetic comparisons, actual threshold and nested record evidence.
// Notes: Endpoint-based measurement does not call AvailabilityUnion or round to minutes.
let
    TestDate = #date(2026, 7, 22),
    DayStart = DateTime.From(TestDate),
    DayEnd = DayStart + #duration(1, 0, 0, 0),
    Threshold = #"EXTRACT EffectiveShiftHrs",
    CandidateShifts = Table.SelectRows(AvailabilityShiftWindows,
        each [Role] = "AIN" and [Date] = TestDate and Time.From([ShiftStart]) = #time(6, 0, 0)),
    Target = if Table.RowCount(CandidateShifts) = 1 then CandidateShifts{0}
        else error Error.Record("Select test shift",
            "Expected exactly one AIN shift on 22 July 2026 starting at 06:00.", CandidateShifts),
    ShiftStart = Target[ShiftStart],
    ShiftEnd = Target[ShiftEnd],
    // This is a date-specific test: compare only the portion of the selected shift on 22 July.
    ComparedShiftEnd = List.Min({ShiftEnd, DayEnd}),

    // Independent calculation: split the range at every endpoint, then count each covered span once.
    // Duplicated or overlapping records cannot add hours to the same span twice.
    MeasureCoveredHours = (windows as list, rangeStart as datetime, rangeEnd as datetime) as number =>
        let
            Overlapping = List.Select(windows, each [Start] < rangeEnd and [End] > rangeStart),
            Clipped = List.Transform(Overlapping,
                each [Start = List.Max({[Start], rangeStart}), End = List.Min({[End], rangeEnd})]),
            Boundaries = List.Sort(List.Distinct({rangeStart, rangeEnd}
                & List.Transform(Clipped, each [Start]) & List.Transform(Clipped, each [End]))),
            Spans = List.Zip({List.RemoveLastN(Boundaries, 1), List.Skip(Boundaries, 1)}),
            Covered = List.Select(Spans, (span) => List.AnyTrue(List.Transform(Clipped,
                (window) => window[Start] < span{1} and window[End] > span{0}))),
            Hours = List.Accumulate(Covered, 0,
                (total, span) => total + Duration.TotalHours(span{1} - span{0}))
        in Hours,

    // Known-answer checks run when the test is evaluated in Power Query.
    At = (hour as number) as datetime => DayStart + #duration(0, hour, 0, 0),
    FixtureChecks = {
        MeasureCoveredHours({}, DayStart, DayEnd) = 0,
        MeasureCoveredHours({[Start = At(6), End = At(10)], [Start = At(6), End = At(10)]}, DayStart, DayEnd) = 4,
        MeasureCoveredHours({[Start = At(6), End = At(10)], [Start = At(8), End = At(12)]}, DayStart, DayEnd) = 6,
        MeasureCoveredHours({[Start = At(14), End = At(22)]}, At(6), At(14)) = 0,
        MeasureCoveredHours({[Start = DayStart - #duration(0, 2, 0, 0), End = At(2)]}, DayStart, DayEnd) = 2,
        MeasureCoveredHours({[Start = At(22), End = DayEnd + #duration(0, 4, 0, 0)]}, DayStart, DayEnd) = 2
    },
    CheckedRecords = if List.AllTrue(FixtureChecks) then AIN_20260722_Records
        else error "Independent hours calculation failed a known-answer fixture.",

    // Preserve worker identity even when the record query has only a no-record placeholder.
    Workers = Table.Group(CheckedRecords, {"EmployeeID", "Facility-Abbrev", "Role", "BaseName", "EligibilityIssue"},
        {{"SourceRows", each _, type table}}),
    WithRules = Table.NestedJoin(Workers, {"EmployeeID", "Facility-Abbrev", "Role"},
        WorkerAvailabilityRules, {"EmployeeID", "Facility-Abbrev", "Role"}, "Rules", JoinKind.LeftOuter),
    CalculatedShift = Table.SelectRows(ResDayShift_Calculated,
        each [Role] = "AIN" and [Date] = TestDate and [Shift] = Target[Shift]
            and [ShiftStart] = ShiftStart and [ShiftEnd] = ShiftEnd),
    WithCalculation = Table.NestedJoin(WithRules, {"EmployeeID", "Facility-Abbrev", "Role"},
        CalculatedShift, {"EmployeeID", "Facility-Abbrev", "Role"}, "Calculated", JoinKind.LeftOuter),
    Measured = Table.AddColumn(WithCalculation, "Test", each
        let
            Rows = [SourceRows],
            FalseRows = Table.SelectRows(Rows, each [CountsAsExclusion]),
            Intervals = List.Transform(Table.ToRecords(FalseRows),
                each [Start = [StartOn22July], End = [EndOn22July]]),
            SimpleSum = List.Accumulate(FalseRows[ExcludedRecordHours], 0, (total, hours) => total + hours),
            IndependentHours = MeasureCoveredHours(Intervals, DayStart, DayEnd),
            UnionIntervals = AvailabilityUnion(Intervals),
            ProductionHours = List.Accumulate(UnionIntervals, 0,
                (total, window) => total + Duration.TotalHours(window[End] - window[Start])),
            // Compare actual worker-rule exclusions with raw recorded exclusions on the morning shift.
            RawMorningHours = MeasureCoveredHours(Intervals, ShiftStart, ComparedShiftEnd),
            RuleCount = Table.RowCount([Rules]),
            Applied = if RuleCount = 1 then [Rules]{0}[ExcludedIntervals] else {},
            AppliedMorningHours = if RuleCount = 1 then MeasureCoveredHours(Applied, ShiftStart, ComparedShiftEnd) else null,
            CalculationCount = Table.RowCount([Calculated]),
            AvailableHours = if CalculationCount = 1 then [Calculated]{0}[AvailableHours] else null,
            EffectiveHours = if CalculationCount = 1 then [Calculated]{0}[EffectiveShiftHrs] else null,
            // Exercise the actual exclusion helper using only this date's recorded intervals.
            HelperOutput = AvailabilityDailyExclusions(Intervals, Threshold),
            HelperDayHours = MeasureCoveredHours(HelperOutput, DayStart, DayEnd),
            ExpectedHelperHours = if IndependentHours >= Threshold then 24 else IndependentHours,
            ArithmeticMatches = Number.Abs(ProductionHours - IndependentHours) < 0.00000001,
            HelperMatches = Number.Abs(HelperDayHours - ExpectedHelperHours) < 0.00000001,
            SourceFailures = Table.RowCount(Table.SelectRows(Rows,
                each [SourceRecord] <> null and ([OriginalMatchCount] <> 1 or [RecordIssue] <> null))),
            Validated = if SourceFailures > 0 then "CHECK SOURCE RECORDS"
                else if not ArithmeticMatches then "FAIL: union hours differ from independent hours"
                else if not HelperMatches then "FAIL: daily exclusion helper disagrees with the threshold rule"
                else "PASS: arithmetic matches (does not approve the full-day rule)"
        in [
            TestDate = TestDate,
            ComparedShiftStart = ShiftStart,
            ComparedShiftEnd = ComparedShiftEnd,
            ExclusionRecordCount = Table.RowCount(FalseRows),
            SumOfExcludedRecordHours = SimpleSum,
            IndependentDistinctHours = IndependentHours,
            ProductionDistinctHours = ProductionHours,
            ArithmeticDifference = ProductionHours - IndependentHours,
            DuplicateOrOverlapHours = SimpleSum - IndependentHours,
            ConfiguredThresholdHours = Threshold,
            ConfiguredThresholdIs7_6 = Number.Abs(Threshold - 7.6) < 0.00000001,
            ThresholdReached = IndependentHours >= Threshold,
            RawExcludedMorningHours = RawMorningHours,
            AppliedExcludedMorningHours = AppliedMorningHours,
            ExtraMorningHoursExcluded = if AppliedMorningHours = null then null else AppliedMorningHours - RawMorningHours,
            HelperExcludedDayHours = HelperDayHours,
            AvailableHours = AvailableHours,
            EffectiveShiftHrs = EffectiveHours,
            ArithmeticCheck = Validated,
            SourceProblemRows = SourceFailures,
            WorkerRuleRows = RuleCount,
            CalculatedShiftRows = CalculationCount,
            ProductionUnionIntervals = UnionIntervals
        ], type record),
    Expanded = Table.ExpandRecordColumn(Measured, "Test",
        {"TestDate", "ComparedShiftStart", "ComparedShiftEnd", "ExclusionRecordCount", "SumOfExcludedRecordHours", "IndependentDistinctHours",
         "ProductionDistinctHours", "ArithmeticDifference", "DuplicateOrOverlapHours", "ConfiguredThresholdHours",
         "ConfiguredThresholdIs7_6", "ThresholdReached", "RawExcludedMorningHours", "AppliedExcludedMorningHours",
         "ExtraMorningHoursExcluded", "HelperExcludedDayHours", "AvailableHours", "EffectiveShiftHrs",
         "ArithmeticCheck", "SourceProblemRows", "WorkerRuleRows", "CalculatedShiftRows", "ProductionUnionIntervals"}),
    Readable = Table.ReorderColumns(Expanded,
        {"EmployeeID", "BaseName", "TestDate", "ExclusionRecordCount", "SumOfExcludedRecordHours",
         "IndependentDistinctHours", "ProductionDistinctHours", "ArithmeticDifference", "ConfiguredThresholdHours",
         "ThresholdReached", "RawExcludedMorningHours", "AppliedExcludedMorningHours", "ExtraMorningHoursExcluded",
         "AvailableHours", "EffectiveShiftHrs", "ArithmeticCheck", "EligibilityIssue"})
in
    Table.Sort(Readable, {{"IndependentDistinctHours", Order.Descending}, {"BaseName", Order.Ascending}})
