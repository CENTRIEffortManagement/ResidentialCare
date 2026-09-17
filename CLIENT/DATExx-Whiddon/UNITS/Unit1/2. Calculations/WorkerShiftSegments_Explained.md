# WorkerShiftSegments explained

This inspection version expands `WorkerShiftSegments` into six commented `let ... in` stages. It retains intermediate columns and uses the original rules for `AvailableHours`, `FullShift`, `SegmentsValid`, and `Disjoint`. Its table is wider than the production output so you can inspect the calculation.

Source: `Capacity-ShiftAvailability.xlsx_PowerQuery.m` in this folder. The original `WorkerShiftSegments` remains unchanged in that file.

## Follow the stages

1. **WorkerShifts:** Match each worker to the shifts configured for their role. Each row then represents one worker and shift.
2. **AvailabilityWindows:** Use explicit availability where present, otherwise the whole shift; trim the available windows to the shift boundaries.
3. **RemainingSegments:** Subtract overlapping exclusions. A break in the middle can split availability into two segments.
4. **MeasuredHours:** Total segment durations in hours and check whether a single segment covers the entire shift.
5. **SegmentChecks:** Check positive durations, shift boundaries, exclusions, and containment within availability.
6. **DisjointChecks:** Compare adjacent segment endpoints to check that segments do not overlap.

Each stage has its own `let` block of named calculations and an `in` that returns its final table. The outer `in` returns the final stage. The top-level Applied Steps show these six stages; the comments explain the smaller steps within each block.

## Copy into Power Query

Create a blank query named `WorkerShiftSegments_Explained` in the same workbook, then paste the following code into its Advanced Editor. It depends on the existing `WorkerAvailabilityRules`, `AvailabilityShiftWindows`, and `AvailabilitySubtract` queries. Keep it connection-only when using it for inspection.

The code below is a standalone `let ... in` expression, ready to paste. M syntax was checked; the inspection query has not been refreshed in Excel.

```powerquery
// Query: WorkerShiftSegments_Explained
// Purpose: Inspect the worker/shift calculation through commented let/in stages without dropping columns.
// Inputs: WorkerAvailabilityRules, AvailabilityShiftWindows and AvailabilitySubtract.
// Output: Same four measurement rules as WorkerShiftSegments, with intermediate tables and lists retained.
// Notes: Inspection query only; keep connection-only. Existing production queries do not depend on it.
let
    // Match each worker to the configured shifts for their role.
    WorkerShifts = let
        // Group the roster shifts into one nested table per role.
        ShiftsByRole = Table.Group(
            AvailabilityShiftWindows, {"Role"},
            {{"Shifts", each _, type table}}
        ),

        // Attach the matching shift group to each worker.
        Joined = Table.NestedJoin(
            WorkerAvailabilityRules, {"Role"},
            ShiftsByRole, {"Role"},
            "Windows", JoinKind.LeftOuter
        ),

        // Preserve the original error when a worker's role has no shifts.
        MissingShifts = Table.SelectRows(Joined, each Table.IsEmpty([Windows])),
        Checked = if Table.IsEmpty(MissingShifts) then Joined
            else error "An eligible worker role has no configured roster shifts.",

        ShiftTables = Table.AddColumn(
            Checked, "Shifts", each [Windows]{0}[Shifts], type table
        ),

        // Expand a copy, retaining Windows and Shifts for inspection.
        ExpansionCopy = Table.DuplicateColumn(ShiftTables, "Shifts", "ShiftsToExpand"),

        // The grain becomes one row per worker and configured shift.
        WorkerShiftRows = Table.ExpandTableColumn(
            ExpansionCopy, "ShiftsToExpand",
            {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}
        )
    in
        WorkerShiftRows,

    // Determine which available intervals fall inside each shift.
    AvailabilityWindows = let
        Source = WorkerShifts,

        // Explicit AVAIL history restricts the baseline to those intervals.
        // Without explicit AVAIL history, start with the whole shift.
        StartingWindows = Table.AddColumn(
            Source, "StartingAvailability",
            each if [HasAvailability] then [AvailableIntervals]
                else {[Start = [ShiftStart], End = [ShiftEnd]]},
            type list
        ),

        // Intervals include Start and exclude End.
        // A window ending exactly at shift start does not overlap the shift.
        OverlappingWindows = Table.AddColumn(
            StartingWindows, "OverlappingAvailability",
            (row) => List.Select(
                row[StartingAvailability],
                (window) => window[Start] < row[ShiftEnd]
                    and window[End] > row[ShiftStart]
            ), type list
        ),

        // Trim each interval to the shift boundaries.
        // Buffer the list because subtraction and validation reuse it.
        ClippedWindows = Table.AddColumn(
            OverlappingWindows, "ClippedAvailability",
            (row) => List.Buffer(List.Transform(
                row[OverlappingAvailability],
                (window) => [
                    Start = List.Max({window[Start], row[ShiftStart]}),
                    End = List.Min({window[End], row[ShiftEnd]})
                ]
            )), type list
        )
    in
        ClippedWindows,

    // Subtract exclusions from the available parts of each shift.
    RemainingSegments = let
        Source = AvailabilityWindows,

        // Full-day exclusion rules were already applied in WorkerAvailabilityRules.
        // Here, select only exclusions that intersect this particular shift.
        MatchingExclusions = Table.AddColumn(
            Source, "ShiftExclusions",
            (row) => List.Buffer(List.Select(
                row[ExcludedIntervals],
                (cut) => cut[Start] < row[ShiftEnd]
                    and cut[End] > row[ShiftStart]
            )), type list
        ),

        // Subtract exclusions. A cut in the middle can leave two separate segments.
        // Buffer the result because hours and validation read the same segments.
        RemainingSegments = Table.AddColumn(
            MatchingExclusions, "RemainingSegments",
            each List.Buffer(AvailabilitySubtract(
                [ClippedAvailability], [ShiftExclusions]
            )), type list
        ),

        CountedSegments = Table.AddColumn(
            RemainingSegments, "SegmentCount",
            each List.Count([RemainingSegments]), Int64.Type
        )
    in
        CountedSegments,

    // Calculate available clock hours and identify fully available shifts.
    MeasuredHours = let
        Source = RemainingSegments,

        // Retain the duration of every remaining segment.
        SegmentDurations = Table.AddColumn(
            Source, "SegmentDurations",
            each List.Transform(
                [RemainingSegments], (segment) => segment[End] - segment[Start]
            ), type list
        ),

        SegmentHours = Table.AddColumn(
            SegmentDurations, "SegmentHours",
            each List.Transform([SegmentDurations], Duration.TotalHours), type list
        ),

        // Preserve the original accumulation order. An empty list totals zero.
        TotalHours = Table.AddColumn(
            SegmentHours, "AvailableHours",
            each List.Accumulate([SegmentHours], 0, (total, hours) => total + hours),
            type number
        ),

        SingleSegment = Table.AddColumn(
            TotalHours, "SingleSegment",
            each [SegmentCount] = 1, type logical
        ),

        StartsAtShiftStart = Table.AddColumn(
            SingleSegment, "StartsAtShiftStart",
            each if [SingleSegment]
                then [RemainingSegments]{0}[Start] = [ShiftStart] else false,
            type logical
        ),

        EndsAtShiftEnd = Table.AddColumn(
            StartsAtShiftStart, "EndsAtShiftEnd",
            each if [SingleSegment]
                then [RemainingSegments]{0}[End] = [ShiftEnd] else false,
            type logical
        ),

        // FullShift requires one uninterrupted segment covering both endpoints.
        // EffectiveShiftHrs is still calculated downstream, not in this stage.
        FullShift = Table.AddColumn(
            EndsAtShiftEnd, "FullShift",
            each [SingleSegment] and [StartsAtShiftStart] and [EndsAtShiftEnd],
            type logical
        )
    in
        FullShift,

    // Verify every segment respects shift boundaries, availability and exclusions.
    SegmentChecks = let
        Source = MeasuredHours,

        PositiveDurations = Table.AddColumn(
            Source, "PositiveDurationChecks",
            each List.Transform(
                [RemainingSegments], (segment) => segment[Start] < segment[End]
            ), type list
        ),

        InsideShift = Table.AddColumn(
            PositiveDurations, "InsideShiftChecks",
            (row) => List.Transform(
                row[RemainingSegments],
                (segment) => segment[Start] >= row[ShiftStart]
                    and segment[End] <= row[ShiftEnd]
            ), type list
        ),

        // No remaining segment may overlap an exclusion.
        AvoidsExclusions = Table.AddColumn(
            InsideShift, "AvoidsExclusionChecks",
            (row) => List.Transform(
                row[RemainingSegments],
                (segment) => not List.AnyTrue(List.Transform(
                    row[ShiftExclusions],
                    (cut) => segment[Start] < cut[End] and segment[End] > cut[Start]
                ))
            ), type list
        ),

        // Each segment must fit completely inside an allowed availability window.
        InsideAvailability = Table.AddColumn(
            AvoidsExclusions, "InsideAvailabilityChecks",
            (row) => List.Transform(
                row[RemainingSegments],
                (segment) => List.AnyTrue(List.Transform(
                    row[ClippedAvailability],
                    (window) => segment[Start] >= window[Start]
                        and segment[End] <= window[End]
                ))
            ), type list
        ),

        // Zero segments passes these checks: no availability is not an invalid result.
        SegmentsValid = Table.AddColumn(
            InsideAvailability, "SegmentsValid",
            each List.AllTrue([PositiveDurationChecks])
                and List.AllTrue([InsideShiftChecks])
                and List.AllTrue([AvoidsExclusionChecks])
                and List.AllTrue([InsideAvailabilityChecks]),
            type logical
        )
    in
        SegmentsValid,

    // Finish the original measurements by checking that remaining segments do not overlap.
    DisjointChecks = let
        Source = SegmentChecks,

        SegmentEnds = Table.AddColumn(
            Source, "SegmentEndTimes",
            each List.Transform([RemainingSegments], (segment) => segment[End]),
            type list
        ),

        SegmentStarts = Table.AddColumn(
            SegmentEnds, "SegmentStartTimes",
            each List.Transform([RemainingSegments], (segment) => segment[Start]),
            type list
        ),

        // Subtraction preserves order. Match each previous end to the next start.
        // These shorten helper lists; all table columns remain available.
        PreviousEnds = Table.AddColumn(
            SegmentStarts, "PreviousSegmentEnds",
            each if [SegmentCount] < 2 then {}
                else List.RemoveLastN([SegmentEndTimes], 1),
            type list
        ),

        NextStarts = Table.AddColumn(
            PreviousEnds, "NextSegmentStarts",
            each if [SegmentCount] < 2 then {}
                else List.Skip([SegmentStartTimes], 1),
            type list
        ),

        EndpointPairs = Table.AddColumn(
            NextStarts, "AdjacentEndpointPairs",
            each List.Zip({[PreviousSegmentEnds], [NextSegmentStarts]}), type list
        ),

        // Touching endpoints are allowed; overlapping segments are not.
        NoOverlapChecks = Table.AddColumn(
            EndpointPairs, "NoOverlapChecks",
            each List.Transform([AdjacentEndpointPairs], (pair) => pair{0} <= pair{1}),
            type list
        ),

        // Zero or one segment is automatically disjoint.
        Disjoint = Table.AddColumn(
            NoOverlapChecks, "Disjoint",
            each List.AllTrue([NoOverlapChecks]), type logical
        )
    in
        Table.Buffer(Disjoint)
in
    DisjointChecks
```
