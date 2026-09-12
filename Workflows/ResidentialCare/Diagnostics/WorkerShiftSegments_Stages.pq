// Inspection-only stages; requires WorkerAvailabilityRules with the corrected daily baseline.
section WorkerShiftSegmentsStages;

// Query: WSS_WorkerShifts
// Purpose: Match each worker to the configured shifts for their role.
shared WSS_WorkerShifts = let
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
    WorkerShiftRows;

// Query: WSS_AvailabilityWindows
// Purpose: Determine which available intervals fall inside each shift.
shared WSS_AvailabilityWindows = let
    Source = WSS_WorkerShifts,

    // WorkerAvailabilityRules already resets the baseline per calendar day.
    // AvailableIntervals includes full days without AVAIL, even across midnight.
    StartingWindows = Table.AddColumn(
        Source, "StartingAvailability",
        each [AvailableIntervals],
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
    ClippedWindows;

// Query: WSS_RemainingSegments
// Purpose: Subtract exclusions from the available parts of each shift.
shared WSS_RemainingSegments = let
    Source = WSS_AvailabilityWindows,

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
    CountedSegments;

// Query: WSS_Hours
// Purpose: Calculate available clock hours and identify fully available shifts.
shared WSS_Hours = let
    Source = WSS_RemainingSegments,

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
    FullShift;

// Query: WSS_SegmentChecks
// Purpose: Verify every segment respects shift boundaries, availability and exclusions.
shared WSS_SegmentChecks = let
    Source = WSS_Hours,

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
    SegmentsValid;

// Query: WSS_DisjointChecks
// Purpose: Finish the original measurements by checking that remaining segments do not overlap.
shared WSS_DisjointChecks = let
    Source = WSS_SegmentChecks,

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
    Table.Buffer(Disjoint);
