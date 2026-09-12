// Query: WorkerShiftSegments_StepByStep
// Purpose: Expose the availability calculation as individually inspectable Applied Steps.
// Inputs: Existing WorkerAvailabilityRules, AvailabilityShiftWindows and AvailabilitySubtract.
// Output: Original worker/shift measurements plus retained intermediate columns.
let
    WorkerRules = WorkerAvailabilityRules,

    ShiftsByRole = Table.Group(
        AvailabilityShiftWindows, {"Role"},
        {{"Shifts", each _, type table}}
    ),

    JoinedShiftGroups = Table.NestedJoin(
        WorkerRules, {"Role"}, ShiftsByRole, {"Role"},
        "Windows", JoinKind.LeftOuter
    ),

    MissingShiftGroups = Table.SelectRows(
        JoinedShiftGroups, each Table.IsEmpty([Windows])
    ),

    CheckedRoleCoverage =
        if Table.IsEmpty(MissingShiftGroups) then JoinedShiftGroups
        else error "An eligible worker role has no configured roster shifts.",

    AddedShiftTables = Table.AddColumn(
        CheckedRoleCoverage, "Shifts",
        each [Windows]{0}[Shifts], type table
    ),

    // Expand a copy so both original nested tables remain visible.
    CopiedShiftTables = Table.DuplicateColumn(
        AddedShiftTables, "Shifts", "ShiftsToExpand"
    ),

    // The row grain changes here: one worker becomes one row per configured shift.
    ExpandedWorkerShifts = Table.ExpandTableColumn(
        CopiedShiftTables, "ShiftsToExpand",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}
    ),

    AddedWholeShiftWindow = Table.AddColumn(
        ExpandedWorkerShifts, "WholeShiftWindow",
        each {[Start = [ShiftStart], End = [ShiftEnd]]}, type list
    ),

    // Use daily baseline intervals; HasAvailability does not gate other dates.
    AddedStartingAvailability = Table.AddColumn(
        AddedWholeShiftWindow, "StartingAvailability",
        each [AvailableIntervals],
        type list
    ),

    // Start is inclusive and End is exclusive. Touching endpoints do not overlap.
    AddedOverlappingAvailability = Table.AddColumn(
        AddedStartingAvailability, "OverlappingAvailability",
        (row) => List.Select(
            row[StartingAvailability],
            (interval) => interval[Start] < row[ShiftEnd]
                and interval[End] > row[ShiftStart]
        ), type list
    ),

    AddedClippedAvailability = Table.AddColumn(
        AddedOverlappingAvailability, "ClippedAvailability",
        (row) => List.Buffer(List.Transform(
            row[OverlappingAvailability],
            (interval) => [
                Start = List.Max({interval[Start], row[ShiftStart]}),
                End = List.Min({interval[End], row[ShiftEnd]})
            ]
        )), type list
    ),

    // Exclusion intervals already include the upstream full-day exclusion rules.
    AddedShiftExclusions = Table.AddColumn(
        AddedClippedAvailability, "ShiftExclusions",
        (row) => List.Buffer(List.Select(
            row[ExcludedIntervals],
            (interval) => interval[Start] < row[ShiftEnd]
                and interval[End] > row[ShiftStart]
        )), type list
    ),

    // Buffer the remaining segments because hours and validation reuse them.
    AddedRemainingSegments = Table.AddColumn(
        AddedShiftExclusions, "RemainingSegments",
        each List.Buffer(AvailabilitySubtract(
            [ClippedAvailability], [ShiftExclusions]
        )), type list
    ),

    AddedSegmentCount = Table.AddColumn(
        AddedRemainingSegments, "SegmentCount",
        each List.Count([RemainingSegments]), Int64.Type
    ),

    AddedSegmentDurations = Table.AddColumn(
        AddedSegmentCount, "SegmentDurations",
        each List.Transform(
            [RemainingSegments], (segment) => segment[End] - segment[Start]
        ), type list
    ),

    AddedSegmentHours = Table.AddColumn(
        AddedSegmentDurations, "SegmentHours",
        each List.Transform([SegmentDurations], Duration.TotalHours), type list
    ),

    // Preserve the original sum order and zero result for no remaining segments.
    AddedAvailableHours = Table.AddColumn(
        AddedSegmentHours, "AvailableHours",
        each List.Accumulate([SegmentHours], 0, (total, hours) => total + hours),
        type number
    ),

    AddedSingleSegmentFlag = Table.AddColumn(
        AddedAvailableHours, "SingleSegment",
        each [SegmentCount] = 1, type logical
    ),

    AddedStartsAtShiftStartFlag = Table.AddColumn(
        AddedSingleSegmentFlag, "StartsAtShiftStart",
        each if [SingleSegment]
            then [RemainingSegments]{0}[Start] = [ShiftStart] else false,
        type logical
    ),

    AddedEndsAtShiftEndFlag = Table.AddColumn(
        AddedStartsAtShiftStartFlag, "EndsAtShiftEnd",
        each if [SingleSegment]
            then [RemainingSegments]{0}[End] = [ShiftEnd] else false,
        type logical
    ),

    // AvailableHours is clock time; effective-hours allowances are applied downstream.
    AddedFullShiftFlag = Table.AddColumn(
        AddedEndsAtShiftEndFlag, "FullShift",
        each [SingleSegment] and [StartsAtShiftStart] and [EndsAtShiftEnd],
        type logical
    ),

    AddedPositiveDurationChecks = Table.AddColumn(
        AddedFullShiftFlag, "PositiveDurationChecks",
        each List.Transform(
            [RemainingSegments], (segment) => segment[Start] < segment[End]
        ), type list
    ),

    AddedInsideShiftChecks = Table.AddColumn(
        AddedPositiveDurationChecks, "InsideShiftChecks",
        (row) => List.Transform(
            row[RemainingSegments],
            (segment) => segment[Start] >= row[ShiftStart]
                and segment[End] <= row[ShiftEnd]
        ), type list
    ),

    AddedAvoidsExclusionChecks = Table.AddColumn(
        AddedInsideShiftChecks, "AvoidsExclusionChecks",
        (row) => List.Transform(
            row[RemainingSegments],
            (segment) => not List.AnyTrue(List.Transform(
                row[ShiftExclusions],
                (cut) => segment[Start] < cut[End] and segment[End] > cut[Start]
            ))
        ), type list
    ),

    AddedInsideAvailabilityChecks = Table.AddColumn(
        AddedAvoidsExclusionChecks, "InsideAvailabilityChecks",
        (row) => List.Transform(
            row[RemainingSegments],
            (segment) => List.AnyTrue(List.Transform(
                row[ClippedAvailability],
                (window) => segment[Start] >= window[Start]
                    and segment[End] <= window[End]
            ))
        ), type list
    ),

    AddedAllDurationsPositiveFlag = Table.AddColumn(
        AddedInsideAvailabilityChecks, "AllDurationsPositive",
        each List.AllTrue([PositiveDurationChecks]), type logical
    ),

    AddedAllSegmentsInsideShiftFlag = Table.AddColumn(
        AddedAllDurationsPositiveFlag, "AllSegmentsInsideShift",
        each List.AllTrue([InsideShiftChecks]), type logical
    ),

    AddedAllSegmentsAvoidExclusionsFlag = Table.AddColumn(
        AddedAllSegmentsInsideShiftFlag, "AllSegmentsAvoidExclusions",
        each List.AllTrue([AvoidsExclusionChecks]), type logical
    ),

    AddedAllSegmentsInsideAvailabilityFlag = Table.AddColumn(
        AddedAllSegmentsAvoidExclusionsFlag, "AllSegmentsInsideAvailability",
        each List.AllTrue([InsideAvailabilityChecks]), type logical
    ),

    // The original flag passes only when every segment passes all four conditions.
    // An empty segment list passes: zero availability is not a validation error.
    AddedSegmentsValidFlag = Table.AddColumn(
        AddedAllSegmentsInsideAvailabilityFlag, "SegmentsValid",
        each [AllDurationsPositive] and [AllSegmentsInsideShift]
            and [AllSegmentsAvoidExclusions] and [AllSegmentsInsideAvailability],
        type logical
    ),

    // Subtraction preserves order. Compare each segment's end with the next start.
    AddedSegmentEndTimes = Table.AddColumn(
        AddedSegmentsValidFlag, "SegmentEndTimes",
        each List.Transform([RemainingSegments], (segment) => segment[End]),
        type list
    ),

    AddedSegmentStartTimes = Table.AddColumn(
        AddedSegmentEndTimes, "SegmentStartTimes",
        each List.Transform([RemainingSegments], (segment) => segment[Start]),
        type list
    ),

    AddedPreviousSegmentEnds = Table.AddColumn(
        AddedSegmentStartTimes, "PreviousSegmentEnds",
        each if [SegmentCount] < 2 then {} else List.RemoveLastN([SegmentEndTimes], 1),
        type list
    ),

    AddedNextSegmentStarts = Table.AddColumn(
        AddedPreviousSegmentEnds, "NextSegmentStarts",
        each if [SegmentCount] < 2 then {} else List.Skip([SegmentStartTimes], 1),
        type list
    ),

    AddedAdjacentEndpointPairs = Table.AddColumn(
        AddedNextSegmentStarts, "AdjacentEndpointPairs",
        each List.Zip({[PreviousSegmentEnds], [NextSegmentStarts]}), type list
    ),

    AddedNoOverlapChecks = Table.AddColumn(
        AddedAdjacentEndpointPairs, "NoOverlapChecks",
        each List.Transform([AdjacentEndpointPairs], (pair) => pair{0} <= pair{1}),
        type list
    ),

    AddedDisjointFlag = Table.AddColumn(
        AddedNoOverlapChecks, "Disjoint",
        each List.AllTrue([NoOverlapChecks]), type logical
    )

    // Keep all intermediate columns and zero-hour rows.
in
    Table.Buffer(AddedDisjointFlag)
