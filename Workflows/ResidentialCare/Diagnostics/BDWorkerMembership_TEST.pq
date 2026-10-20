section BDWorkerMembershipFixtures;

// Query: BDWorkerMembershipFixtures
// Purpose: Test the BD membership adapter in isolation from availability records, contracts and shift calculations.
// Inputs: Actual production evaluators and a read-only snapshot of the current Employees_TABLE.
// Output: Focused pass/fail cases for the reported interface failure and preserved membership meanings.
shared BDWorkerMembershipFixtures = (evaluate as function, evaluateLegacy as function, invalidMembershipRows as function, currentWorkers as table) as table =>
let
    Navigation = (workers as table) as table =>
        #table(type table [Item = text, Kind = text, Data = table], {{"Employees_TABLE", "Table", workers}}),
    EvaluateRows = (workers as table) as record => evaluate([Navigation = Navigation(workers)]),
    Current = EvaluateRows(currentWorkers),
    Prepared = Table.Buffer(Current[Prepared]),
    Grouped = Table.Buffer(Current[Grouped]),
    LegacyAttempt = try Table.RowCount(evaluateLegacy(Navigation(currentWorkers))),
    SourcePairs = List.Zip({Table.ToRecords(currentWorkers), Table.ToRecords(Prepared)}),
    CorrectFlags = List.AllTrue(List.Transform(SourcePairs, (pair) =>
        pair{0}[In Roster] = pair{1}[In Allocation]
        and pair{0}[In Availabilities] = pair{1}[In Records]
        and pair{0}[In Records] = pair{1}[In HR Records]
        and pair{0}[In Leave Balances] = pair{1}[In Leave Balances])),
    ExpectedWorkers = #table(type table [EmployeeID = text, Name = text], {
        {"24344", "Alexis Griffin"}, {"26705", "Ainura Borbueva"},
        {"25161", "Deeanne Tracey"}, {"25184", "Shelby Bon"}, {"23245", "Majal Dias"}}),
    CorrectWorkerIdentities = List.AllTrue(List.Transform(Table.ToRecords(ExpectedWorkers), (expected) =>
        let
            Matches = Table.SelectRows(Prepared, each [EmployeeID] = expected[EmployeeID] and [#"Facility-Abbrev"] = "BD"),
            IsCorrect = if Table.RowCount(Matches) <> 1 then false else
                Matches{0}[Employee Roster Name] = expected[Name]
                and Matches{0}[Role] = "Assistant in Nursing"
        in
            IsCorrect)),
    SourceLeaveOnly = Table.SelectRows(currentWorkers, each not [In Roster] and not [In Availabilities] and [In Leave Balances]),
    GroupedLeaveOnly = Table.SelectRows(Grouped, each [Worker Membership] = "Leave balances only"),
    SyntheticWorkers = #table(type table [EmployeeID = text, #"Facility-Abbrev" = text, #"Employee Roster Name" = text,
        EmploymentType = text, Role = text, #"In Roster" = logical, #"In Availabilities" = logical,
        #"In Leave Balances" = logical, #"In Records" = logical], {
        {"00017", "BD", "Roster Worker", "Casual", "Assistant in Nursing AM", true, false, false, true},
        {"2", "BD", "Availability Worker", "Casual", "Assistant in Nursing", false, true, false, false},
        {"3", "BD", "Both Worker", "Part-Time", "Assistant in Nursing", true, true, false, false},
        {"4", "BD", "Leave Worker", "Part-Time", "Registered Nurse", false, false, true, true}}),
    Synthetic = EvaluateRows(SyntheticWorkers),
    SyntheticPrepared = Table.Buffer(Synthetic[Prepared]),
    SyntheticGrouped = Table.Buffer(Synthetic[Grouped]),
    ExpectedMemberships = #table(type table [EmployeeID = text, Membership = text], {
        {"00017", "Allocation only"}, {"2", "Records only"}, {"3", "Both"}, {"4", "Leave balances only"}}),
    CorrectMemberships = List.AllTrue(List.Transform(Table.ToRecords(ExpectedMemberships), (expected) =>
        let
            Matches = Table.SelectRows(SyntheticGrouped, each [EmployeeID] = expected[EmployeeID]),
            IsCorrect = if Table.RowCount(Matches) <> 1 then false else Matches{0}[Worker Membership] = expected[Membership]
        in
            IsCorrect)),
    MissingFlagAttempt = try Table.RowCount(EvaluateRows(Table.RemoveColumns(currentWorkers, {"In Availabilities"}))[Prepared]),
    LeadingZeroWorker = Table.SelectRows(SyntheticPrepared, each [EmployeeID] = "00017"),
    MakeResult = (name as text, passed as logical) as record => [Check = name, Passed = passed],
    Results = {
        MakeResult("Original reader reproduces the missing-membership error", LegacyAttempt[HasError]
            and Text.Contains(LegacyAttempt[Error][Message], "worker-membership columns")),
        MakeResult("Adapter and grouping preserve every captured worker", Table.RowCount(Prepared) = Table.RowCount(currentWorkers)
            and Table.RowCount(Grouped) = Table.RowCount(currentWorkers)),
        MakeResult("Roster, availability, HR and leave flags retain their meanings", CorrectFlags),
        MakeResult("Four unmatched workers and Majal retain current identities and full roles", CorrectWorkerIdentities),
        MakeResult("Leave-only workers remain in the model", Table.RowCount(SourceLeaveOnly) > 0
            and Table.RowCount(GroupedLeaveOnly) = Table.RowCount(SourceLeaveOnly)),
        MakeResult("Adapter introduces no duplicate employee/facility keys", Table.RowCount(Prepared)
            = Table.RowCount(Table.Distinct(Prepared, {"EmployeeID", "Facility-Abbrev"}))),
        MakeResult("Existing membership categories and leave-only membership evaluate correctly", CorrectMemberships),
        MakeResult("Employee code leading zeros and complete role value survive", Table.RowCount(LeadingZeroWorker) = 1
            and LeadingZeroWorker{0}[Role] = "Assistant in Nursing AM"),
        MakeResult("Missing availability membership is still rejected", MissingFlagAttempt[HasError]),
        MakeResult("Existing membership validation accepts the retained leave-only category", Table.IsEmpty(invalidMembershipRows(Grouped))
            and Table.IsEmpty(invalidMembershipRows(SyntheticGrouped)))
    },
    Output = Table.FromRecords(Results, type table [Check = text, Passed = logical])
in
    Output;
