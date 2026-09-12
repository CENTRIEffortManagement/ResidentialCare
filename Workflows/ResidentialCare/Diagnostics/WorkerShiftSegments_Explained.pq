// Query: WorkerShiftSegments_Explained
// Purpose: Show each worker/shift availability calculation without removing inspection columns.
// Inputs: Existing WorkerAvailabilityRules, AvailabilityShiftWindows and AvailabilitySubtract queries.
// Output: Same worker/shift rows and four measurement results, plus intermediate columns.
// Notes: Create as a new connection-only query. Retained nested data increases preview cost.
let
    // 1. Keep the starting worker rules and group all configured shifts by role.
    WorkerRules = WorkerAvailabilityRules,
    ShiftsByRole = Table.Group(
        AvailabilityShiftWindows,
        {"Role"},
        {{"Shifts", each _, type table}}
    ),

    // 2. Attach the matching shift table without multiplying worker rows yet.
    WorkersWithShiftGroups = Table.NestedJoin(
        WorkerRules, {"Role"},
        ShiftsByRole, {"Role"},
        "Windows", JoinKind.LeftOuter
    ),
    WorkersMissingShifts = Table.SelectRows(
        WorkersWithShiftGroups,
        each Table.IsEmpty([Windows])
    ),
    ValidatedRoleCoverage =
        if Table.IsEmpty(WorkersMissingShifts) then WorkersWithShiftGroups
        else error "An eligible worker role has no configured roster shifts.",

    // 3. Preserve both nested tables. Expand a copy to create one row per worker/shift.
    WorkersWithShiftTables = Table.AddColumn(
        ValidatedRoleCoverage, "Shifts",
        each [Windows]{0}[Shifts], type table
    ),
    ShiftTablesForExpansion = Table.DuplicateColumn(
        WorkersWithShiftTables, "Shifts", "ShiftsToExpand"
    ),
    WorkerShiftRows = Table.ExpandTableColumn(
        ShiftTablesForExpansion, "ShiftsToExpand",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}
    ),

    // 4. Use the corrected baseline, which resets independently on each calendar day.
    // AvailableIntervals includes whole days without AVAIL; exclusions are applied afterwards.
    WithStartingAvailability = Table.AddColumn(
        WorkerShiftRows, "StartingAvailability",
        each [AvailableIntervals],
        type list
    ),

    // 5. Keep baseline intervals intersecting this shift.
    // Intervals include Start and exclude End: touching boundaries do not overlap.
    WithOverlappingAvailability = Table.AddColumn(
        WithStartingAvailability, "OverlappingAvailability",
        each
            let
                WindowStart = [ShiftStart],
                WindowEnd = [ShiftEnd]
            in
                List.Select(
                    [StartingAvailability],
                    (interval) => interval[Start] < WindowEnd
                        and interval[End] > WindowStart
                ),
        type list
    ),

    // 6. Trim available intervals to the shift boundaries.
    // Buffer the resulting lists because subtraction and validation both read them.
    WithClippedAvailability = Table.AddColumn(
        WithOverlappingAvailability, "ClippedAvailability",
        each
            let
                WindowStart = [ShiftStart],
                WindowEnd = [ShiftEnd]
            in
                List.Buffer(List.Transform(
                    [OverlappingAvailability],
                    (interval) => [
                        Start = List.Max({interval[Start], WindowStart}),
                        End = List.Min({interval[End], WindowEnd})
                    ]
                )),
        type list
    ),

    // 7. Select exclusions intersecting this shift.
    // Full-day exclusion rules have already been applied in WorkerAvailabilityRules.
    WithShiftExclusions = Table.AddColumn(
        WithClippedAvailability, "ShiftExclusions",
        each
            let
                WindowStart = [ShiftStart],
                WindowEnd = [ShiftEnd]
            in
                List.Buffer(List.Select(
                    [ExcludedIntervals],
                    (interval) => interval[Start] < WindowEnd
                        and interval[End] > WindowStart
                )),
        type list
    ),

    // 8. Subtract exclusions. An internal exclusion can split one interval into two.
    WithRemainingSegments = Table.AddColumn(
        WithShiftExclusions, "RemainingSegments",
        each List.Buffer(AvailabilitySubtract(
            [ClippedAvailability], [ShiftExclusions]
        )),
        type list
    ),

    // 9. Show the clock hours contributed by every remaining segment.
    WithSegmentHours = Table.AddColumn(
        WithRemainingSegments, "SegmentHours",
        each List.Transform(
            [RemainingSegments],
            (segment) => Duration.TotalHours(segment[End] - segment[Start])
        ),
        type list
    ),
    WithAvailableHours = Table.AddColumn(
        WithSegmentHours, "AvailableHours",
        each List.Accumulate([SegmentHours], 0, (total, hours) => total + hours),
        type number
    ),
    WithSegmentCount = Table.AddColumn(
        WithAvailableHours, "SegmentCount",
        each List.Count([RemainingSegments]), Int64.Type
    ),

    // 10. A full shift requires one uninterrupted segment with exactly the shift endpoints.
    // EffectiveShiftHrs is calculated downstream; AvailableHours remains clock hours here.
    WithFullShiftFlag = Table.AddColumn(
        WithSegmentCount, "FullShift",
        each [SegmentCount] = 1
            and [RemainingSegments]{0}[Start] = [ShiftStart]
            and [RemainingSegments]{0}[End] = [ShiftEnd],
        type logical
    ),

    // 11. Keep a record of each segment's validation results for inspection.
    WithSegmentChecks = Table.AddColumn(
        WithFullShiftFlag, "SegmentChecks",
        each
            let
                WindowStart = [ShiftStart],
                WindowEnd = [ShiftEnd],
                Cuts = [ShiftExclusions],
                AvailableWindows = [ClippedAvailability]
            in
                List.Transform(
                    [RemainingSegments],
                    (segment) => [
                        Start = segment[Start],
                        End = segment[End],
                        PositiveDuration = segment[Start] < segment[End],
                        InsideShift = segment[Start] >= WindowStart
                            and segment[End] <= WindowEnd,
                        AvoidsExclusions = not List.AnyTrue(List.Transform(
                            Cuts,
                            (cut) => segment[Start] < cut[End]
                                and segment[End] > cut[Start]
                        )),
                        InsideAvailability = List.AnyTrue(List.Transform(
                            AvailableWindows,
                            (window) => segment[Start] >= window[Start]
                                and segment[End] <= window[End]
                        ))
                    ]
                ),
        type list
    ),
    WithSegmentsValidFlag = Table.AddColumn(
        WithSegmentChecks, "SegmentsValid",
        each List.AllTrue(List.Transform(
            [SegmentChecks],
            (check) => check[PositiveDuration] and check[InsideShift]
                and check[AvoidsExclusions] and check[InsideAvailability]
        )),
        type logical
    ),

    // 12. Subtraction preserves segment order. Check each pair of neighbours.
    // Zero or one segment is disjoint; touching segments do not overlap.
    WithAdjacentSegmentChecks = Table.AddColumn(
        WithSegmentsValidFlag, "AdjacentSegmentChecks",
        each
            let
                Segments = [RemainingSegments],
                Count = [SegmentCount]
            in
                if Count < 2 then {}
                else List.Transform(
                    {1..(Count - 1)},
                    (index) => [
                        PreviousEnd = Segments{index - 1}[End],
                        NextStart = Segments{index}[Start],
                        NoOverlap = Segments{index - 1}[End] <= Segments{index}[Start]
                    ]
                ),
        type list
    ),
    WithDisjointFlag = Table.AddColumn(
        WithAdjacentSegmentChecks, "Disjoint",
        each List.AllTrue(List.Transform(
            [AdjacentSegmentChecks], (check) => check[NoOverlap]
        )),
        type logical
    )

    // Preserve zero-hour rows for validation and every intermediate column for inspection.
in
    Table.Buffer(WithDisjointFlag)
