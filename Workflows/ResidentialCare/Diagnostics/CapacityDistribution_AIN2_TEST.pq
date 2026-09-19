section CapacityDistributionAIN2Tests;

// Synthetic fixtures only: this section never opens, inspects, or refreshes a workbook.
// Source-contract checks that bind these behaviours to A1, A2, and B live in
// scripts/test-capacity-distribution-ain2.ps1.

// Query: fnAIN2TestResult
// Purpose: Evaluate one synthetic invariant and return a consistent test-result row.
shared fnAIN2TestResult = (name as text, category as text, evaluate as function) as record =>
    let
        Attempt = try evaluate(),
        Passed = not Attempt[HasError] and Attempt[Value] = true,
        Detail =
            if Attempt[HasError] then
                Record.FieldOrDefault(Attempt[Error], "Message", "Fixture evaluation failed.")
            else if Passed then
                "Invariant satisfied."
            else
                "Fixture returned false."
    in
        [Fixture = name, Category = category, Passed = Passed, Detail = Detail];

// Query: fnAIN2TestRoleContext
// Purpose: Model the approved AIN2 development and AIN promoted role tuples.
// Output: RoleFolder, SourceRole, OutputRole, and the compatibility Role alias.
shared fnAIN2TestRoleContext = (roleFolder as text) as record =>
    let
        Folder = Text.Upper(Text.Trim(roleFolder)),
        ValidatedFolder =
            if List.Contains({"AIN2", "AIN"}, Folder) then Folder
            else error Error.Record("CapacityDistribution.RoleContext", "Expected role folder AIN2 or AIN.", [RoleFolder = roleFolder]),
        SourceRole = "AIN",
        OutputRole = if ValidatedFolder = "AIN2" then "AIN2" else "AIN"
    in
        [RoleFolder = ValidatedFolder, SourceRole = SourceRole, OutputRole = OutputRole, Role = OutputRole];

// Query: fnAIN2TestPreviousCluster
// Purpose: Resolve a preceding cluster using Resource plus the preceding positional index.
// Notes: ResIndex intentionally restarts for each Resource, so an index-only lookup is invalid.
shared fnAIN2TestPreviousCluster = (rows as table) as table =>
    let
        Lookup = Table.RenameColumns(
            Table.SelectColumns(rows, {"Resource", "ResIndex", "ClusterNumber"}),
            {{"ResIndex", "LookupIndex"}, {"ClusterNumber", "PreviousCluster"}}
        ),
        Joined = Table.NestedJoin(rows, {"Resource", "PrevIndex"}, Lookup, {"Resource", "LookupIndex"}, "Previous", JoinKind.LeftOuter),
        Expanded = Table.ExpandTableColumn(Joined, "Previous", {"PreviousCluster"}, {"PreviousCluster"})
    in
        Expanded;

// Query: fnAIN2TestLinearTotals
// Purpose: Model A2's linear running totals after converting one-based group starts to its two-based row index.
// Inputs: IndexTwoBased must be globally contiguous; StartOneBased is the existing A2 group-start convention.
shared fnAIN2TestLinearTotals = (rows as table) as table =>
    let
        RequiredColumns = {"IndexTwoBased", "StartOneBased", "Count", "Value"},
        MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(rows)),
        CheckedColumns =
            if List.IsEmpty(MissingColumns) then rows
            else error Error.Record("CapacityDistribution.RunningTotal", "Fixture is missing positional columns.", [MissingColumns = MissingColumns]),
        Buffered = Table.Buffer(CheckedColumns),
        RowCount = Table.RowCount(Buffered),
        Indices = List.Buffer(Buffered[IndexTwoBased]),
        Starts = List.Buffer(List.Transform(Buffered[StartOneBased], each if _ = null then null else _ + 1)),
        Counts = List.Buffer(Buffered[Count]),
        Values = List.Buffer(Buffered[Value]),
        ExpectedIndices = List.Numbers(2, RowCount),
        ControlsValid = Value.Equals(Indices, ExpectedIndices)
            and List.AllTrue(List.Transform(List.Positions(Indices), (position) =>
                let
                    CurrentIndex = Indices{position},
                    CurrentStart = Starts{position},
                    CurrentCount = Counts{position},
                    PreviousStart = if position = 0 then null else Starts{position - 1}
                in
                    Value.Is(CurrentIndex, Int64.Type)
                        and Value.Is(CurrentStart, Int64.Type)
                        and Value.Is(CurrentCount, Int64.Type)
                        and CurrentStart >= 2
                        and CurrentStart <= CurrentIndex
                        and CurrentCount = CurrentIndex - CurrentStart + 1
                        and (position = 0 or CurrentStart = PreviousStart or CurrentStart = CurrentIndex)
            )),
        CheckedControls =
            if ControlsValid then true
            else error Error.Record("CapacityDistribution.RunningTotal", "Invalid or unnormalised positional controls.", null),
        FinalState = List.Accumulate(
            List.Positions(Values),
            [PreviousStart = null, Running = null, Totals = {}],
            (state as record, position as number) as record =>
                let
                    CurrentStart = Starts{position},
                    CurrentValue = Values{position},
                    NextRunning =
                        if position = 0 or CurrentStart <> state[PreviousStart] then CurrentValue
                        else if CurrentValue = null then state[Running]
                        else if state[Running] = null then CurrentValue
                        else state[Running] + CurrentValue
                in
                    [PreviousStart = CurrentStart, Running = NextRunning, Totals = state[Totals] & {NextRunning}]
        ),
        Result = if CheckedControls then
            Table.FromColumns(Table.ToColumns(Buffered) & {FinalState[Totals]}, Table.ColumnNames(Buffered) & {"RunningTotal"})
            else error "Unreachable positional validation state."
    in
        Result;

// Query: fnAIN2TestSkeletonRole
// Purpose: Preserve Role from the complete Resource-period skeleton, including missing availability rows.
shared fnAIN2TestSkeletonRole = (skeleton as table, availability as table) as table =>
    let
        AvailabilityMeasures = Table.SelectColumns(availability, {"Resource", "Period", "Availability"}),
        Joined = Table.NestedJoin(skeleton, {"Resource", "Period"}, AvailabilityMeasures, {"Resource", "Period"}, "AvailabilityRows", JoinKind.LeftOuter),
        Expanded = Table.ExpandTableColumn(Joined, "AvailabilityRows", {"Availability"}, {"Availability"})
    in
        Expanded;

// Query: fnAIN2TestStablePriority
// Purpose: Make priority ties deterministic by finishing the sort with Resource and Period.
shared fnAIN2TestStablePriority = (rows as table) as table =>
    Table.Sort(rows, {{"Priority", Order.Ascending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}});

// Query: CapacityDistribution_AIN2_TEST
// Purpose: Exercise role translation, resource-keyed clusters, index normalisation, skeleton Role, and tie ordering.
// Output: One result row per independent synthetic fixture.
shared CapacityDistribution_AIN2_TEST =
    let
        DevelopmentRole = fnAIN2TestRoleContext("ain2"),
        ProductionRole = fnAIN2TestRoleContext("AIN"),
        // Force the validated folder field; SourceRole is deliberately constant and would not force validation under M's lazy evaluation.
        UnexpectedRole = try fnAIN2TestRoleContext("RN")[RoleFolder],

        ClusterRows = #table(
            type table [Resource = Int64.Type, ResIndex = Int64.Type, PrevIndex = nullable number, ClusterNumber = text],
            {{101, 2, null, "R101-C1"}, {101, 3, 2, "R101-C2"}, {202, 2, null, "R202-C1"}, {202, 3, 2, "R202-C2"}}
        ),
        ClusterActual = Table.SelectRows(fnAIN2TestPreviousCluster(ClusterRows), each [PrevIndex] <> null),
        ClusterExpected = {
            [Resource = 101, ResIndex = 3, PrevIndex = 2, ClusterNumber = "R101-C2", PreviousCluster = "R101-C1"],
            [Resource = 202, ResIndex = 3, PrevIndex = 2, ClusterNumber = "R202-C2", PreviousCluster = "R202-C1"]
        },

        RunningRows = #table(
            type table [IndexTwoBased = Int64.Type, StartOneBased = Int64.Type, Count = Int64.Type, Value = nullable number],
            {{2, 1, 1, 1}, {3, 1, 2, 2}, {4, 1, 3, null}, {5, 4, 1, 4}, {6, 4, 2, -1}}
        ),
        RunningActual = fnAIN2TestLinearTotals(RunningRows)[RunningTotal],
        InvalidRunningRows = Table.TransformColumns(RunningRows, {{"Count", each if _ = 2 then 3 else _, Int64.Type}}),
        InvalidRunningAttempt = try List.Sum(fnAIN2TestLinearTotals(InvalidRunningRows)[RunningTotal]),

        Skeleton = #table(type table [Role = text, Resource = Int64.Type, Period = Int64.Type], {{"AIN2", 101, 1}, {"AIN2", 101, 2}, {"AIN2", 202, 1}}),
        Availability = #table(type table [Role = text, Resource = Int64.Type, Period = Int64.Type, Availability = nullable number], {{"WRONG", 101, 1, 1}, {"WRONG", 202, 1, null}}),
        SkeletonActual = fnAIN2TestSkeletonRole(Skeleton, Availability),

        TiedRows = #table(
            type table [RowId = text, Priority = Int64.Type, Resource = Int64.Type, Period = Int64.Type],
            {{"r2p2", 1, 2, 2}, {"r1p2", 1, 1, 2}, {"r2p1", 1, 2, 1}, {"r1p1", 1, 1, 1}}
        ),
        TiedExpected = {"r1p1", "r1p2", "r2p1", "r2p2"},
        Results = {
            fnAIN2TestResult("role-development-tuple", "Role context", () => Value.Equals(DevelopmentRole, [RoleFolder = "AIN2", SourceRole = "AIN", OutputRole = "AIN2", Role = "AIN2"])),
            fnAIN2TestResult("role-production-tuple", "Role context", () => Value.Equals(ProductionRole, [RoleFolder = "AIN", SourceRole = "AIN", OutputRole = "AIN", Role = "AIN"])),
            fnAIN2TestResult("role-unexpected-folder-rejected", "Role context", () => UnexpectedRole[HasError]),
            fnAIN2TestResult("cluster-index-is-resource-scoped", "A1 cluster", () => Value.Equals(Table.ToRecords(ClusterActual), ClusterExpected)),
            fnAIN2TestResult("a2-one-based-start-normalised", "A2 running total", () => Value.Equals(RunningActual, {1, 3, 3, 4, 3})),
            fnAIN2TestResult("a2-invalid-count-rejected", "A2 running total", () => InvalidRunningAttempt[HasError]),
            fnAIN2TestResult("skeleton-owns-role", "A2 Role", () => List.AllTrue(List.Transform(SkeletonActual[Role], each _ = "AIN2"))),
            fnAIN2TestResult("skeleton-retains-missing-availability", "A2 Role", () => Table.RowCount(SkeletonActual) = 3 and SkeletonActual{1}[Availability] = null),
            fnAIN2TestResult("priority-ties-use-resource-period", "Deterministic order", () => Value.Equals(fnAIN2TestStablePriority(TiedRows)[RowId], TiedExpected)),
            fnAIN2TestResult("priority-order-independent-of-input", "Deterministic order", () => Value.Equals(fnAIN2TestStablePriority(Table.ReverseRows(TiedRows))[RowId], TiedExpected))
        }
    in
        Table.FromRecords(Results);

// Query: CapacityDistribution_AIN2_CHECK
// Purpose: Fail native validation if an AIN2 architecture fixture fails or fixture names are duplicated.
// Output: The complete result table after all fixtures pass.
shared CapacityDistribution_AIN2_CHECK =
    let
        Results = Table.Buffer(Table.TransformColumnTypes(
            CapacityDistribution_AIN2_TEST,
            {{"Fixture", type text}, {"Category", type text}, {"Passed", type logical}, {"Detail", type text}}
        )),
        Failures = Table.SelectRows(Results, each [Passed] <> true),
        DuplicateNames = Table.RowCount(Results) <> List.Count(List.Distinct(Results[Fixture]))
    in
        if DuplicateNames then
            error Error.Record("CapacityDistribution.TestFailure", "AIN2 fixture names must be unique.", null)
        else if not Table.IsEmpty(Failures) then
            error Error.Record("CapacityDistribution.TestFailure", "AIN2 synthetic fixture validation failed.", Failures)
        else
            Results;
