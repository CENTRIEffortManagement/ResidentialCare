section CapacityDistribB_InputKeysTests;

// Synthetic fixtures only: no workbook is opened, inspected or refreshed.
// The native test harness supplies wrappers extracted from the production B source:
// fnCapacityDistribBTestInputCheck(demand, allocation, capped, priority),
// fnCapacityDistribBTestPriorityJoin(left, right), and
// fnCapacityDistribBTestOutputGate(checks, outputThunk).

// Query: CapacityDistribB_InputKeys_Fixtures
// Purpose: Define independent expected diagnostics for the four production join inputs.
// Notes: Keys use type any so null, blank and erroneous cells reach the production checker unchanged.
shared CapacityDistribB_InputKeys_Fixtures =
    let
        Inputs = [
            Demand = #table(type table [Period = any, D = nullable number], {{11, 1}, {22, null}}),
            Allocation = #table(type table [Resource = any, Period = any, Allocation = nullable number], {{1, 11, null}, {2, 22, 0.5}}),
            Capped = #table(type table [Resource = any, Period = any, #"C#" = nullable number], {{1, 11, 1}, {2, 22, null}}),
            Priority = #table(type table [Resource = any, Period = any, NWDPriority = nullable number], {{1, 11, 8}, {2, 22, null}})
        ],
        Expected = {
            [Input = "PeriodDemandTABLE", Keys = "Period", MissingColumns = "", ErrorKeyRows = 0, MissingKeyRows = 0, DuplicateKeyGroups = 0, Passed = true],
            [Input = "ResPeriodAllocationTABLE", Keys = "Resource, Period", MissingColumns = "", ErrorKeyRows = 0, MissingKeyRows = 0, DuplicateKeyGroups = 0, Passed = true],
            [Input = "ResPeriodAvailabilityCapped(C#)TABLE", Keys = "Resource, Period", MissingColumns = "", ErrorKeyRows = 0, MissingKeyRows = 0, DuplicateKeyGroups = 0, Passed = true],
            [Input = "IMPORT ResPeriodNWDTABLE", Keys = "Resource, Period", MissingColumns = "", ErrorKeyRows = 0, MissingKeyRows = 0, DuplicateKeyGroups = 0, Passed = true]
        },
        InputSpecs = {
            [Field = "Demand", Position = 0, Keys = {"Period"}],
            [Field = "Allocation", Position = 1, Keys = {"Resource", "Period"}],
            [Field = "Capped", Position = 2, Keys = {"Resource", "Period"}],
            [Field = "Priority", Position = 3, Keys = {"Resource", "Period"}]
        },
        KeySpecs = List.Combine(List.Transform(InputSpecs, (inputSpec) =>
            List.Transform(inputSpec[Keys], (key) => [Field = inputSpec[Field], Position = inputSpec[Position], Key = key])
        )),
        MakeFixture = (name as text, spec as record, transform as function, expectedChange as record) as record => [
            Name = name,
            Inputs = Record.TransformFields(Inputs, {{spec[Field], transform}}),
            Expected = List.Transform(List.Positions(Expected), (position) =>
                if position = spec[Position] then Record.Combine({Expected{position}, expectedChange}) else Expected{position}
            )
        ],
        ValidCases = {
            [Name = "valid-four-inputs-nullable-measures", Inputs = Inputs, Expected = Expected],
            [Name = "valid-empty-required-schemas", Inputs = Record.TransformFields(Inputs, List.Transform(Record.FieldNames(Inputs), (field) => {field, each Table.FirstN(_, 0)})), Expected = Expected]
        },
        MissingColumnCases = List.Transform(KeySpecs, (spec) => MakeFixture(
            "missing-column-" & spec[Field] & "-" & spec[Key],
            spec,
            (rows) => Table.RemoveColumns(rows, {spec[Key]}),
            [MissingColumns = spec[Key], ErrorKeyRows = null, MissingKeyRows = null, DuplicateKeyGroups = null, Passed = false]
        )),
        MissingBothColumnsCase = MakeFixture(
            "missing-both-composite-key-columns",
            InputSpecs{1},
            (rows) => Table.RemoveColumns(rows, {"Resource", "Period"}),
            [MissingColumns = "Resource, Period", ErrorKeyRows = null, MissingKeyRows = null, DuplicateKeyGroups = null, Passed = false]
        ),
        NullKeyCases = List.Transform(KeySpecs, (spec) => MakeFixture(
            "null-key-" & spec[Field] & "-" & spec[Key],
            spec,
            (rows) => Table.TransformColumns(rows, {{spec[Key], each null, type any}}),
            [MissingKeyRows = 2, Passed = false]
        )),
        BlankKeyCases = List.Transform(KeySpecs, (spec) => MakeFixture(
            "blank-key-" & spec[Field] & "-" & spec[Key],
            spec,
            (rows) => Table.TransformColumns(rows, {{spec[Key], each " #(tab) ", type any}}),
            [MissingKeyRows = 2, Passed = false]
        )),
        EmptyTextKeyCase = MakeFixture(
            "empty-text-period-key",
            InputSpecs{0},
            (rows) => Table.TransformColumns(rows, {{"Period", each "", type any}}),
            [MissingKeyRows = 2, Passed = false]
        ),
        KeyErrorCases = List.Transform(KeySpecs, (spec) => MakeFixture(
            "key-cell-error-" & spec[Field] & "-" & spec[Key],
            spec,
            (rows) => Table.TransformColumns(rows, {{spec[Key], each error Error.Record("Fixture.KeyCellError", "Deliberate synthetic key error.", null), type any}}),
            [ErrorKeyRows = 2, Passed = false]
        )),
        DuplicateCases = List.Transform(InputSpecs, (spec) => MakeFixture(
            "duplicate-key-" & spec[Field],
            spec,
            (rows) => Table.Combine({rows, Table.FirstN(rows, 1)}),
            [DuplicateKeyGroups = 1, Passed = false]
        )),
        RetainedNullAvailabilityCase = MakeFixture(
            "duplicate-capped-key-retained-null-availability",
            InputSpecs{2},
            // The protected query retains zero-to-null C# rows; their keys still participate in joins.
            (rows) => Table.Combine({rows, Table.LastN(rows, 1)}),
            [DuplicateKeyGroups = 1, Passed = false]
        ),
        MultipleDuplicateGroupsCase = MakeFixture(
            "duplicate-group-count-distinct-from-extra-row-count",
            InputSpecs{1},
            (rows) => Table.Combine({rows, rows, Table.FirstN(rows, 1)}),
            [DuplicateKeyGroups = 2, Passed = false]
        ),
        MixedDefectsCase = MakeFixture(
            "key-errors-missing-and-duplicates-counted-separately",
            InputSpecs{0},
            (rows) => #table(type table [Period = any, D = nullable number], {
                {11, 1}, {11, null}, {null, 2}, {" ", null},
                {error Error.Record("Fixture.KeyCellError", "Deliberate synthetic key error.", null), 3}
            }),
            [ErrorKeyRows = 1, MissingKeyRows = 2, DuplicateKeyGroups = 1, Passed = false]
        )
    in
        ValidCases & MissingColumnCases & {MissingBothColumnsCase} & NullKeyCases & BlankKeyCases &
        {EmptyTextKeyCase} & KeyErrorCases & DuplicateCases &
        {RetainedNullAvailabilityCase, MultipleDuplicateGroupsCase, MixedDefectsCase};

// Query: CapacityDistribB_InputKeys_TEST
// Purpose: Compare exact production diagnostics with independently declared expected rows for every fixture.
// Output: One scalar result per fixture; unexpected checker errors are test failures.
shared CapacityDistribB_InputKeys_TEST =
    let
        RunFixture = (fixture as record) as record =>
            let
                Inputs = fixture[Inputs],
                Actual = fnCapacityDistribBTestInputCheck(Inputs[Demand], Inputs[Allocation], Inputs[Capped], Inputs[Priority]),
                // Comparing complete records forces every diagnostic cell, including error and missing-key counts.
                Attempt = try Value.Equals(Table.ToRecords(Actual), fixture[Expected]),
                Passed = if Attempt[HasError] then false else Attempt[Value],
                Detail = if Attempt[HasError] then Record.FieldOrDefault(Attempt[Error], "Message", "Checker evaluation failed.")
                    else if Passed then "All four diagnostic records match expected keys, counts and pass status."
                    else "Production diagnostics differ from the independently expected records."
            in
                [Fixture = fixture[Name], Category = "Input keys", Passed = Passed, Detail = Detail]
    in
        Table.FromRecords(List.Transform(CapacityDistribB_InputKeys_Fixtures, RunFixture));

// Query: CapacityDistribB_PriorityJoin_TEST
// Purpose: Prove that the exact production priority join pairs Resource with Resource and Period with Period.
// Notes: Reversed numeric pairs deliberately exist on the right, so a positional key reversal cannot pass accidentally.
shared CapacityDistribB_PriorityJoin_TEST =
    let
        RightRows = #table(type table [Resource = number, Period = number, NWDPriority = nullable number], {
            {1, 11, 101}, {11, 1, 1101}, {2, 22, null}
        }),
        Fixtures = {
            [Name = "priority-join-distinct-reversed-key-pairs", Left = #table(type table [RowId = text, Resource = number, Period = number], {{"first", 1, 11}, {"second", 11, 1}}),
                Expected = {[RowId = "first", Resource = 1, Period = 11, NWDPriority = 101], [RowId = "second", Resource = 11, Period = 1, NWDPriority = 1101]}],
            [Name = "priority-join-preserves-unmatched-left-row", Left = #table(type table [RowId = text, Resource = number, Period = number], {{"unmatched", 3, 33}}),
                Expected = {[RowId = "unmatched", Resource = 3, Period = 33, NWDPriority = null]}],
            [Name = "priority-join-allows-null-priority-measure", Left = #table(type table [RowId = text, Resource = number, Period = number], {{"nullable", 2, 22}}),
                Expected = {[RowId = "nullable", Resource = 2, Period = 22, NWDPriority = null]}]
        },
        RunFixture = (fixture as record) as record =>
            let
                Attempt = try Value.Equals(Table.ToRecords(fnCapacityDistribBTestPriorityJoin(fixture[Left], RightRows)), fixture[Expected]),
                Passed = if Attempt[HasError] then false else Attempt[Value]
            in
                [Fixture = fixture[Name], Category = "Priority join", Passed = Passed,
                    Detail = if Attempt[HasError] then Record.FieldOrDefault(Attempt[Error], "Message", "Priority join evaluation failed.")
                        else if Passed then "Matched priority and left-row preservation are correct."
                        else "Expanded rows or matched priority differ from expected."]
    in
        Table.FromRecords(List.Transform(Fixtures, RunFixture));

// Query: CapacityDistribB_OutputGate_TEST
// Purpose: Verify that failed input checks block the output branch and successful checks evaluate that branch.
// Notes: A uniquely identified throwing thunk detects unwanted branch evaluation without external side effects.
shared CapacityDistribB_OutputGate_TEST =
    let
        GoodFixture = CapacityDistribB_InputKeys_Fixtures{0},
        GoodInputs = GoodFixture[Inputs],
        GoodChecks = fnCapacityDistribBTestInputCheck(GoodInputs[Demand], GoodInputs[Allocation], GoodInputs[Capped], GoodInputs[Priority]),
        ExpectedOutput = #table(type table [Marker = number], {{42}}),
        ThrowingOutput = () => error Error.Record("Fixture.OutputEvaluated", "The synthetic output branch was evaluated.", null),
        InvalidFixtures = List.Select(CapacityDistribB_InputKeys_Fixtures, each List.AnyTrue(List.Transform([Expected], (row) => not row[Passed]))),
        RejectedCases = List.Transform(InvalidFixtures, (fixture) =>
            let
                Inputs = fixture[Inputs]
            in [
                Name = "gate-rejects-" & fixture[Name],
                Checks = fnCapacityDistribBTestInputCheck(Inputs[Demand], Inputs[Allocation], Inputs[Capped], Inputs[Priority]),
                Output = ThrowingOutput,
                ExpectedReason = "CapacityDistribB.InputValidation"
            ]
        ),
        Fixtures = {
            [Name = "gate-valid-output-passes", Checks = GoodChecks, Output = () => ExpectedOutput, ExpectedReason = null],
            [Name = "gate-valid-checks-evaluate-output-thunk", Checks = GoodChecks, Output = ThrowingOutput, ExpectedReason = "Fixture.OutputEvaluated"]
        } & RejectedCases,
        RunFixture = (fixture as record) as record =>
            let
                // Scalar comparison forces the selected output branch; merely constructing a lazy table is insufficient.
                Attempt = try Value.Equals(Table.ToRecords(fnCapacityDistribBTestOutputGate(fixture[Checks], fixture[Output])), Table.ToRecords(ExpectedOutput)),
                ActualReason = if Attempt[HasError] then Record.FieldOrDefault(Attempt[Error], "Reason", null) else null,
                Passed = if fixture[ExpectedReason] = null then not Attempt[HasError] and Attempt[Value]
                    else Attempt[HasError] and ActualReason = fixture[ExpectedReason],
                Detail = if Passed then
                        if fixture[ExpectedReason] = "CapacityDistribB.InputValidation" then "Input validation failed before the throwing output branch was evaluated."
                        else if fixture[ExpectedReason] = null then "Valid checks returned the expected output unchanged."
                        else "Valid checks evaluated the output branch and propagated its deliberate error."
                    else if Attempt[HasError] then "Unexpected error reason: " & Text.From(ActualReason)
                    else "Gate returned an unexpected successful result."
            in
                [Fixture = fixture[Name], Category = "Output gate", Passed = Passed, Detail = Detail]
    in
        Table.FromRecords(List.Transform(Fixtures, RunFixture));

// Query: CapacityDistribB_InputKeys_CHECK
// Purpose: Fail native validation when any key, corrected-join or output-gate fixture fails.
// Output: The complete typed result table only after all synthetic fixtures pass and fixture names are unique.
shared CapacityDistribB_InputKeys_CHECK =
    let
        Results = Table.Buffer(Table.TransformColumnTypes(
            Table.Combine({CapacityDistribB_InputKeys_TEST, CapacityDistribB_PriorityJoin_TEST, CapacityDistribB_OutputGate_TEST}),
            {{"Fixture", type text}, {"Category", type text}, {"Passed", type logical}, {"Detail", type text}}
        )),
        Failures = Table.SelectRows(Results, each [Passed] <> true),
        DuplicateNames = Table.RowCount(Results) <> List.Count(List.Distinct(Results[Fixture]))
    in
        if DuplicateNames then
            error Error.Record("CapacityDistribB.TestFailure", "Fixture names must be unique.", null)
        else if not Table.IsEmpty(Failures) then
            error Error.Record("CapacityDistribB.TestFailure", "Input-key, priority-join or output-gate fixture validation failed.", Failures)
        else
            Results;
