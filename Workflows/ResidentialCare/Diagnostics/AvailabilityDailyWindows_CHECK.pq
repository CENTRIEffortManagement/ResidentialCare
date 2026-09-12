// Query: AvailabilityDailyWindows_CHECK
// Purpose: Check that AVAIL and exclusions reset per calendar day using synthetic intervals.
// Inputs: Corrected AvailabilityDailyWindows, plus AvailabilityDailyExclusions and AvailabilitySubtract.
// Output: Named regression checks; any failed comparison raises an error.
// Notes: Standalone inspection query, not a workbook refresh script. Fixture dates are synthetic.
let
    At = (day as number, hour as number) as datetime => #datetime(2026, 7, day, hour, 0, 0),
    Window = (startDay as number, startHour as number, endDay as number, endHour as number) as record =>
        [Start = At(startDay, startHour), End = At(endDay, endHour)],
    Baseline = (available as list) as list => AvailabilityDailyWindows(available, At(22, 0), At(24, 0)),
    Check = (name as text, actual as any, expected as any) as record =>
        [Check = name, Passed = actual = expected, Actual = actual, Expected = expected],
    // Synthetic headcount: 57 eligible workers, 18 distinct exclusions on 22 July, no AVAIL.
    // Check the requested 39 result and next-day reset without reading any real worker data.
    SyntheticCount = (day as number) as number => List.Count(List.Select({1..57}, (worker) =>
        let
            Exclusions = if worker <= 18 then {Window(22, 6, 22, 14)} else {},
            Remaining = AvailabilitySubtract(Baseline({}), AvailabilityDailyExclusions(Exclusions, 7.6))
        in List.AnyTrue(List.Transform(Remaining,
            each [Start] < At(day, 14) and [End] > At(day, 6))))),
    WholeHorizon = {Window(22, 0, 24, 0)},
    Results = Table.FromRecords({
        Check("No AVAIL leaves both dates available", Baseline({}), WholeHorizon),
        Check("AVAIL on 22 July does not restrict 23 July",
            Baseline({Window(22, 6, 22, 14)}),
            {Window(22, 6, 22, 14), Window(23, 0, 24, 0)}),
        Check("AVAIL on 23 July does not restrict 22 July",
            Baseline({Window(23, 6, 23, 14)}),
            {Window(22, 0, 23, 0), Window(23, 6, 23, 14)}),
        Check("AVAIL outside the calendar horizon has no effect",
            Baseline({Window(21, 6, 21, 14), Window(24, 6, 24, 14)}), WholeHorizon),
        Check("Exclusive midnight end permits the next whole day",
            Baseline({Window(22, 6, 23, 0)}), {Window(22, 6, 24, 0)}),
        Check("A recorded overnight AVAIL affects both dates only for its interval",
            Baseline({Window(22, 22, 23, 2)}), {Window(22, 22, 23, 2)}),
        Check("Duplicate and overlapping AVAIL does not double count",
            Baseline({Window(22, 6, 22, 12), Window(22, 6, 22, 12), Window(22, 10, 22, 14)}),
            {Window(22, 6, 22, 14), Window(23, 0, 24, 0)}),
        Check("Full-day exclusion on 22 July does not exclude 23 July",
            AvailabilitySubtract(Baseline({}), AvailabilityDailyExclusions({Window(22, 6, 22, 14)}, 7.6)),
            {Window(23, 0, 24, 0)}),
        Check("Exclusions still override AVAIL on their date",
            AvailabilitySubtract(Baseline({Window(22, 6, 22, 14)}),
                AvailabilityDailyExclusions({Window(22, 10, 22, 12)}, 7.6)),
            {Window(22, 6, 22, 10), Window(22, 12, 22, 14), Window(23, 0, 24, 0)}),
        Check("57 workers minus 18 exclusions leaves 39 on 22 July", SyntheticCount(22), 39),
        Check("All 57 synthetic workers reset to available on 23 July", SyntheticCount(23), 57)
    }),
    Failures = Table.SelectRows(Results, each not [Passed])
in
    if Table.IsEmpty(Failures) then Results
    else error Error.Record("Availability daily reset", "Daily availability regression checks failed.", Failures)
