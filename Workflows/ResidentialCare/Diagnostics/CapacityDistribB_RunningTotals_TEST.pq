section CapacityDistribB_RunningTotalsTests;

// These fixtures require fnAddIndexedRunningTotal from the approved B source.
// Evaluate the TEST and CHECK queries in the same M environment as that helper.
// This file contains synthetic data only and never reads or refreshes a workbook.

// Query: fnCapacityDistribBTestRows
// Purpose: Build small typed fixtures with the legacy two-based positional controls.
// Notes: Row labels run in reverse order and priorities tie, so accidental sorting changes the fixture.
shared fnCapacityDistribBTestRows = (groups as list, cap as number) as table =>
    let
        GroupLengths = List.Transform(groups, each List.Count(_)),
        TotalRows = List.Sum({0} & GroupLengths),
        GroupRows = List.Transform(
            List.Positions(groups),
            (groupPosition) =>
                let
                    PriorRows = List.Sum({0} & List.FirstN(GroupLengths, groupPosition)),
                    Values = groups{groupPosition}
                in
                    List.Transform(
                        List.Positions(Values),
                        (valuePosition) => {
                            "row-" & Text.PadStart(Text.From(TotalRows - PriorRows - valuePosition), 3, "0"),
                            1,
                            Values{valuePosition},
                            PriorRows + valuePosition + 2,
                            PriorRows + 2,
                            valuePosition + 1,
                            cap,
                            if Number.Mod(PriorRows + valuePosition, 3) = 2 then "Remove" else "Keep"
                        }
                    )
        ),
        FixtureRows = #table(
            type table [RowId = text, Priority = number, Value = nullable number, Index = number, Start = number, Count = number, Cap = number, IncomingKeep = text],
            List.Combine(GroupRows)
        ),
        TypedControls = Table.TransformColumnTypes(FixtureRows, {{"Index", Int64.Type}, {"Start", Int64.Type}, {"Count", Int64.Type}})
    in
        TypedControls;

// Query: CapacityDistribB_RunningTotals_Fixtures
// Purpose: Declare valid arithmetic, ordering, boundary and invalid-control cases independently of the optimized helper.
// Notes: ExpectedTotals = null requests exact comparison with the legacy M expression for floating-point stress cases.
shared CapacityDistribB_RunningTotals_Fixtures =
    let
        BasicRows = fnCapacityDistribBTestRows({{1, 2, 3}}, 3),
        SeparateGroups = fnCapacityDistribBTestRows({{1}, {2}, {3}}, 3),
        ReopenedGroup = Table.FromRecords(
            List.Transform(Table.ToRecords(SeparateGroups), each if [Index] = 4 then Record.TransformFields(_, {{"Start", (prior) => 2}, {"Count", (prior) => 3}}) else _),
            Value.Type(SeparateGroups)
        ),
        TwentyDayValues = List.Transform(List.Numbers(0, 60), each if Number.Mod(_, 7) = 0 then null else if Number.Mod(_, 5) = 0 then 0 else 0.25),
        TwentyDayExpected = List.Transform(List.Numbers(1, 60), each List.Sum(List.FirstN(TwentyDayValues, _))),
        ValidCases = {
            [Name = "empty", Input = fnCapacityDistribBTestRows({}, 1), ExpectedTotals = {}, ExpectError = false],
            [Name = "singleton", Input = fnCapacityDistribBTestRows({{5}}, 5), ExpectedTotals = {5}, ExpectError = false],
            [Name = "singleton-null", Input = fnCapacityDistribBTestRows({{null}}, 0), ExpectedTotals = {null}, ExpectError = false],
            [Name = "multiple-groups", Input = fnCapacityDistribBTestRows({{1, 2, 3}, {4}, {2, 0, 5}}, 5), ExpectedTotals = {1, 3, 6, 4, 2, 2, 7}, ExpectError = false],
            [Name = "all-null-groups", Input = fnCapacityDistribBTestRows({{null, null}, {null}}, 1), ExpectedTotals = {null, null, null}, ExpectError = false],
            [Name = "null-leading-interior-trailing", Input = fnCapacityDistribBTestRows({{null, 2, null, 0, null}, {null, -1, null}}, 2), ExpectedTotals = {null, 2, 2, 2, 2, null, -1, -1}, ExpectError = false],
            [Name = "zero-is-not-null", Input = fnCapacityDistribBTestRows({{0, null, 0}, {null, 0}}, 0), ExpectedTotals = {0, 0, 0, null, 0}, ExpectError = false],
            [Name = "negative-values", Input = fnCapacityDistribBTestRows({{3, -5, 2, -1}, {-2, 1}}, 0), ExpectedTotals = {3, -2, 0, -1, -2, -1}, ExpectError = false],
            [Name = "binary-fractions", Input = fnCapacityDistribBTestRows({{0.25, 0.5, 0.125}, {1.5, -0.25}}, 0.75), ExpectedTotals = {0.25, 0.75, 0.875, 1.5, 1.25}, ExpectError = false],
            [Name = "tied-priority-input-order", Input = fnCapacityDistribBTestRows({{2, 1, 4}, {3, 2}}, 3), ExpectedTotals = {2, 3, 7, 3, 5}, ExpectError = false],
            [Name = "exact-cap-and-next-row", Input = fnCapacityDistribBTestRows({{0.25, 0.75, 0, 0.125}}, 1), ExpectedTotals = {0.25, 1, 1, 1.125}, ExpectError = false],
            [Name = "decimal-cap-boundary", Input = fnCapacityDistribBTestRows({{0.1, 0.2, 0.3, -0.1}}, 0.3), ExpectedTotals = null, ExpectError = false],
            [Name = "decimal-cancellation", Input = fnCapacityDistribBTestRows({{10000000000000000, 1, -10000000000000000, 0.1, 0.2}}, 0.3), ExpectedTotals = null, ExpectError = false],
            [Name = "twenty-days-three-shifts", Input = fnCapacityDistribBTestRows({TwentyDayValues, TwentyDayValues, TwentyDayValues}, 8), ExpectedTotals = TwentyDayExpected & TwentyDayExpected & TwentyDayExpected, ExpectError = false]
        },
        InvalidCases = {
            [Name = "missing-value-column", Input = Table.RemoveColumns(BasicRows, {"Value"}), ExpectedTotals = null, ExpectError = true],
            [Name = "missing-index-column", Input = Table.RemoveColumns(BasicRows, {"Index"}), ExpectedTotals = null, ExpectError = true],
            [Name = "missing-start-column", Input = Table.RemoveColumns(BasicRows, {"Start"}), ExpectedTotals = null, ExpectError = true],
            [Name = "missing-count-column", Input = Table.RemoveColumns(BasicRows, {"Count"}), ExpectedTotals = null, ExpectError = true],
            [Name = "empty-missing-control-column", Input = Table.RemoveColumns(Table.FirstN(BasicRows, 0), {"Start"}), ExpectedTotals = null, ExpectError = true],
            [Name = "existing-output-column", Input = Table.AddColumn(BasicRows, "RunningTotal", each 99), ExpectedTotals = null, ExpectError = true],
            [Name = "wrong-index-base", Input = Table.TransformColumns(BasicRows, {{"Index", each _ - 1, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "duplicate-index-after-join", Input = Table.TransformColumns(BasicRows, {{"Index", each if _ = 3 then 2 else _, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "index-gap", Input = Table.TransformColumns(BasicRows, {{"Index", each if _ > 2 then _ + 1 else _, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "wrong-input-order", Input = Table.ReverseRows(BasicRows), ExpectedTotals = null, ExpectError = true],
            [Name = "null-index", Input = Table.TransformColumns(BasicRows, {{"Index", each if _ = 3 then null else _, type nullable number}}), ExpectedTotals = null, ExpectError = true],
            [Name = "text-index", Input = Table.TransformColumns(BasicRows, {{"Index", each if _ = 3 then "3" else _, type any}}), ExpectedTotals = null, ExpectError = true],
            [Name = "wrong-first-start", Input = Table.TransformColumns(BasicRows, {{"Start", each 1, Int64.Type}, {"Count", each _ + 1, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "noncontiguous-reopened-start", Input = ReopenedGroup, ExpectedTotals = null, ExpectError = true],
            [Name = "null-start", Input = Table.TransformColumns(BasicRows, {{"Start", each null, type nullable number}}), ExpectedTotals = null, ExpectError = true],
            [Name = "start-after-current-index", Input = Table.TransformColumns(BasicRows, {{"Start", each 5, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "wrong-count", Input = Table.TransformColumns(BasicRows, {{"Count", each _ + 1, Int64.Type}}), ExpectedTotals = null, ExpectError = true],
            [Name = "fractional-count", Input = Table.TransformColumns(BasicRows, {{"Count", each _ + 0.5, type number}}), ExpectedTotals = null, ExpectError = true],
            [Name = "null-count", Input = Table.TransformColumns(BasicRows, {{"Count", each null, type nullable number}}), ExpectedTotals = null, ExpectError = true],
            [Name = "value-error-propagates", Input = Table.TransformColumns(BasicRows, {{"Value", each if _ = 2 then error Error.Record("Fixture.ValueError", "Deliberate fixture cell error.", null) else _, type nullable number}}), ExpectedTotals = null, ExpectError = true]
        }
    in
        ValidCases & InvalidCases;

// Query: fnCapacityDistribBTestDecisions
// Purpose: Capture the existing cap decisions, including errors from null comparisons, without hiding a behavioral change.
// Notes: Each try evaluates a scalar decision; a try around a lazy record or table alone would not force its cells.
shared fnCapacityDistribBTestDecisions = (rows as table, totalColumn as text) as list =>
    let
        CaptureDecision = (decision as function) as record =>
            let
                Attempt = try decision()
            in
                if Attempt[HasError] then
                    [HasError = true, Reason = Record.FieldOrDefault(Attempt[Error], "Reason", null)]
                else
                    [HasError = false, Value = Attempt[Value]],
        Decisions = List.Transform(
            Table.ToRecords(rows),
            (row) =>
                let
                    Total = Record.Field(row, totalColumn),
                    Cap = row[Cap],
                    ResourceKeep = () => if Total <= Cap then "Keep" else "Remove"
                in
                    {
                        CaptureDecision(() => if Total > Cap then "Remove" else if Total <= Cap then "Keep" else "Split"),
                        CaptureDecision(ResourceKeep),
                        CaptureDecision(() => if row[IncomingKeep] <> "Keep" then null else if ResourceKeep() <> "Keep" then null else "Keep")
                    }
        )
    in
        Decisions;

// Query: CapacityDistribB_RunningTotals_TEST
// Purpose: Compare the optimized helper with the original prefix formula and explicit expected totals on synthetic cases.
// Output: One row per fixture, with separate arithmetic, ordering, schema and decision checks.
shared CapacityDistribB_RunningTotals_TEST =
    let
        RunFixture = (fixture as record) as record =>
            let
                Input = fixture[Input],
                Actual = fnAddIndexedRunningTotal(Input, "Value", "Index", "Start", "Count", "RunningTotal"),
                // The reference retains the original per-row prefix formula, independently of the new accumulator.
                Legacy = Table.AddColumn(Input, "RunningTotal", each List.Sum(List.Range(Input[Value], [Start] - 2, [Count]))),
                // Summation forces the running-total cells, so invalid late rows cannot pass through a shallow try.
                EvaluationAttempt = try List.Sum(Table.Column(Actual, "RunningTotal")),
                ActualTotals = Table.Column(Actual, "RunningTotal"),
                TotalsMatch = Value.Equals(ActualTotals, Legacy[RunningTotal]),
                ExplicitTotalsMatch = if fixture[ExpectedTotals] = null then true else Value.Equals(ActualTotals, fixture[ExpectedTotals]),
                ColumnsMatch = Value.Equals(Table.ColumnNames(Actual), Table.ColumnNames(Legacy)),
                RowsMatch = Value.Equals(Table.ToRows(Table.SelectColumns(Actual, Table.ColumnNames(Input))), Table.ToRows(Input)),
                SchemaMatch = Value.Equals(Table.ToRows(Table.Schema(Actual)), Table.ToRows(Table.Schema(Legacy))),
                DecisionsMatch = Value.Equals(fnCapacityDistribBTestDecisions(Actual, "RunningTotal"), fnCapacityDistribBTestDecisions(Legacy, "RunningTotal")),
                ComparisonAttempt = try TotalsMatch and ExplicitTotalsMatch and ColumnsMatch and RowsMatch and SchemaMatch and DecisionsMatch,
                Passed = if fixture[ExpectError] then EvaluationAttempt[HasError] else if EvaluationAttempt[HasError] then false else if ComparisonAttempt[HasError] then false else ComparisonAttempt[Value],
                Detail =
                    if fixture[ExpectError] then
                        if EvaluationAttempt[HasError] then "Expected invalid input was rejected." else "Expected an error; helper accepted invalid input."
                    else if EvaluationAttempt[HasError] then
                        Record.FieldOrDefault(EvaluationAttempt[Error], "Message", "Helper evaluation failed.")
                    else if ComparisonAttempt[HasError] then
                        Record.FieldOrDefault(ComparisonAttempt[Error], "Message", "Comparison failed.")
                    else if Passed then
                        "Exact totals, original rows, column schema and decision outcomes match."
                    else
                        "One or more equivalence checks failed; inspect the individual columns.",
                CanCompare = not fixture[ExpectError] and not EvaluationAttempt[HasError] and not ComparisonAttempt[HasError]
            in
                [
                    Fixture = fixture[Name],
                    ExpectedError = fixture[ExpectError],
                    Passed = Passed,
                    TotalsMatch = if CanCompare then TotalsMatch else null,
                    ExplicitTotalsMatch = if CanCompare then ExplicitTotalsMatch else null,
                    ColumnsMatch = if CanCompare then ColumnsMatch else null,
                    RowsMatch = if CanCompare then RowsMatch else null,
                    SchemaMatch = if CanCompare then SchemaMatch else null,
                    DecisionsMatch = if CanCompare then DecisionsMatch else null,
                    Detail = Detail
                ],
        Results = Table.FromRecords(List.Transform(CapacityDistribB_RunningTotals_Fixtures, RunFixture)),
        TypedResults = Table.TransformColumnTypes(Results, {{"Fixture", type text}, {"ExpectedError", type logical}, {"Passed", type logical}, {"TotalsMatch", type nullable logical}, {"ExplicitTotalsMatch", type nullable logical}, {"ColumnsMatch", type nullable logical}, {"RowsMatch", type nullable logical}, {"SchemaMatch", type nullable logical}, {"DecisionsMatch", type nullable logical}, {"Detail", type text}})
    in
        TypedResults;

// Query: CapacityDistribB_RunningTotals_CHECK
// Purpose: Fail validation when any synthetic running-total fixture fails or fixture names are duplicated.
// Output: The full fixture result table only after all required checks pass.
shared CapacityDistribB_RunningTotals_CHECK =
    let
        Results = Table.Buffer(CapacityDistribB_RunningTotals_TEST),
        Failures = Table.SelectRows(Results, each [Passed] <> true),
        DuplicateNames = Table.RowCount(Results) <> List.Count(List.Distinct(Results[Fixture])),
        CheckedResults =
            if DuplicateNames then
                error Error.Record("CapacityDistribB.TestFailure", "Fixture names must be unique.", null)
            else if not Table.IsEmpty(Failures) then
                error Error.Record("CapacityDistribB.TestFailure", "Running-total fixture validation failed.", Table.SelectColumns(Failures, {"Fixture", "Detail"}))
            else
                Results
    in
        CheckedResults;
