section OvertimeShiftsIntegrationFixtures;

// Query: ShiftsIntegrationFixtures
// Purpose: Exercise the actual old/new Shifts expressions and unchanged interval-allocation consumer with synthetic inputs.
// Inputs: Native evaluator functions generated from the approved source extraction and canonical source.
// Output: Case/Passed/Details rows. No connector, workbook or SharePoint access.
shared ShiftsIntegrationFixtures = (EvaluateBaseline as function, EvaluateNew as function, EvaluateConsumer as function) as table =>
let
    RosterShift = (code as number, name as text, role as text, start as datetime, finish as datetime, shift as text) as record =>
        [Code = code, Name = name, Role = role, Date = Date.From(start), Start = Number.From(Time.From(start)),
            End = Number.From(Time.From(finish)), StartDateTime = start, EndDateTime = finish,
            Duration = Duration.TotalDays(finish - start), Shift = shift,
            EffectiveDuration = Duration.TotalDays(finish - start) - 0.5 / 24],
    RosterRecords = {
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 18, 0, 0), "AM"),
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 3, 14, 0, 0), #datetime(2026, 11, 3, 22, 0, 0), "PM"),
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 4, 22, 0, 0), #datetime(2026, 11, 5, 6, 0, 0), "NIGHT"),
        RosterShift(202, "Ben", "RN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(900, "Master unmatched", "AIN", #datetime(2026, 11, 4, 6, 0, 0), #datetime(2026, 11, 4, 14, 0, 0), "AM")
    },
    Roster = Table.FromRecords(RosterRecords),
    IntervalTable = (roster as table) as table =>
        let
            Boundaries = List.Sort(List.Distinct(List.Combine(List.Transform(Table.ToRecords(roster), each {[StartDateTime], [EndDateTime]})))),
            IntervalRecords = List.Combine(List.Transform({0..List.Count(Boundaries) - 2}, (index) =>
                List.Transform({"AIN", "RN"}, (role) =>
                    [StartInterval = Boundaries{index}, EndPrecise = Boundaries{index + 1},
                        EndInterval = Boundaries{index + 1} - #duration(0, 0, 1, 0),
                        DayDate = Date.From(Boundaries{index}), IntervalID = index + 1,
                        Duration = Duration.TotalDays(Boundaries{index + 1} - Boundaries{index}),
                        Role = role, ShiftPeriod = "AM", IntervalAssociatedShift = "AM"])))
        in Table.FromRecords(IntervalRecords),
    CheckRecord = (row as record, ordinal as number) as record =>
        [SourceUnit = "BD", AllocationCode = Text.From(row[Code]), Role = row[Role],
            ShiftStart = row[StartDateTime], ShiftFinish = row[EndDateTime], SourceShiftID = "S" & Text.From(ordinal),
            EmployeeID = Text.From(row[Code]), RosterShiftPeriod = row[Shift],
            DailyStatus = if ordinal = 1 then "Flag" else "Pass", DailyExcessHours = if ordinal = 1 then 1.5 else 0,
            MinimumBreakFlag = if ordinal = 3 then null else false,
            GapStatus = if ordinal = 3 then "Unknown" else "Known",
            CoverageStatus = if ordinal = 4 then "Unknown" else "Confirmed",
            SourceFile = "synthetic-source", SourceRow = ordinal, SourceName = row[Name], SourceDate = row[Date]],
    CheckRecords = List.Transform({0..3}, each CheckRecord(RosterRecords{_}, _ + 1)),
    Checks = Table.FromRecords(CheckRecords),
    Inputs = (checks as table, optional suppliedRoster as nullable table) as record => [
        Roster = if suppliedRoster = null then Roster else suppliedRoster,
        Checks = checks,
        PathTable = #table({"Variable Name", "Value"}, {{"Unit", "BD"}, {"Root Path", "fixture/2. Calculations"}}),
        IntervalsNavigation = #table({"Item", "Kind", "Data"}, {
            {"Table_StartTime_Duration", "Table", if suppliedRoster = null then Roster else suppliedRoster},
            {"Intervals", "Table", IntervalTable(if suppliedRoster = null then Roster else suppliedRoster)}}),
        SettingsNavigation = #table({"Item", "Kind", "Data"}, {
            {"Meals", "Table", #table({"MealBreakTime", "MealBreakstart"}, {{0.5, 5}})},
            {"ShiftGap", "Table", #table({"MinGap"}, {{8}})},
            {"ShiftLent", "Table", #table({"MinHrs", "ShortHrs", "StdHrs", "1.5OT", "2.0OT"}, {{4, 6, 8, 1.5, 2}})},
            {"ShiftPeriod", "Table", #table({"ShiftPeriod", "StartTime", "StartDay"}, {{"AM", 6 / 24, 6 / 24}, {"PM", 14 / 24, 14 / 24}, {"NIGHT", 22 / 24, 22 / 24}})}}),
        Roles = #table({"ROLES"}, {{"AIN"}, {"RN"}})
    ],
    Baseline = EvaluateBaseline(Inputs(Checks)),
    Current = EvaluateNew(Inputs(Checks)),
    PointColumns = {"Name", "Role", "Duration", "Shift", "IntervalPoint", "DateTime", "StaffCount", "Double Shift", "Effective Duration", "RealDuration", "ShiftType"},
    PointSort = {{"Name", Order.Ascending}, {"DateTime", Order.Ascending}, {"IntervalPoint", Order.Ascending}, {"Role", Order.Ascending}},
    IntervalSort = {{"Name", Order.Ascending}, {"Role", Order.Ascending}, {"TimeDate", Order.Ascending}, {"AllocatedIntervals", Order.Ascending}, {"Attribute", Order.Ascending}, {"Role1", Order.Ascending}},
    EqualProjection = (left as table, right as table, columns as list, sort as list) as logical =>
        Table.ToRows(Table.Sort(Table.SelectColumns(left, columns), sort)) = Table.ToRows(Table.Sort(Table.SelectColumns(right, columns), sort)),
    OldIntervals = Baseline[Intervals],
    NewIntervals = Current[Intervals],
    OldConsumer = EvaluateConsumer(OldIntervals),
    NewConsumer = EvaluateConsumer(NewIntervals),
    Metadata = {"SourceUnit", "AllocationCode", "ShiftStart", "ShiftFinish", "SourceShiftID", "EmployeeID", "RosterShiftPeriod",
        "DailyStatus", "DailyExcessHours", "MinimumBreakFlag", "GapStatus", "CoverageStatus", "OvertimeMatchCount", "OvertimeMatchStatus", "OvertimeMatchReason"},
    SourcePoints = (sourceID as text) as table => Table.SelectRows(Current[Points], each [SourceShiftID] = sourceID),
    SourceIntervals = (sourceID as text) as table => Table.SelectRows(NewIntervals, each [SourceShiftID] = sourceID),
    WithAmbiguity = Table.FromRecords(CheckRecords & {Record.TransformFields(CheckRecords{0}, {{"SourceShiftID", each "duplicate"}})}),
    Ambiguous = EvaluateNew(Inputs(WithAmbiguity)),
    OverlapRows = {
        RosterShift(301, "Shared Name", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(302, "Shared Name", "AIN", #datetime(2026, 11, 2, 7, 0, 0), #datetime(2026, 11, 2, 15, 0, 0), "AM")},
    SameStartRows = {
        RosterShift(401, "Identical Start", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(402, "Identical Start", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 15, 0, 0), "AM")},
    AdjacentRows = {
        RosterShift(501, "Adjacent worker", "AIN", #datetime(2026, 11, 6, 6, 0, 0), #datetime(2026, 11, 6, 10, 0, 0), "AM"),
        RosterShift(501, "Adjacent worker", "AIN", #datetime(2026, 11, 6, 11, 0, 0), #datetime(2026, 11, 6, 15, 0, 0), "AM")},
    OverlapChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(OverlapRows{_}, _ + 1))),
    SameStartChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(SameStartRows{_}, _ + 1))),
    Overlap = EvaluateNew(Inputs(OverlapChecks, Table.FromRecords(OverlapRows))),
    SameStart = EvaluateNew(Inputs(SameStartChecks, Table.FromRecords(SameStartRows))),
    AdjacentChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(AdjacentRows{_}, _ + 1))),
    Adjacent = EvaluateNew(Inputs(AdjacentChecks, Table.FromRecords(AdjacentRows))),
    AdjacentBaseline = EvaluateBaseline(Inputs(AdjacentChecks, Table.FromRecords(AdjacentRows))),
    CanonicalPointCounts = (result as record) as logical =>
        let grouped = Table.Group(result[Published], {"SourceShiftID"}, {{"Count", each Table.RowCount(_), Int64.Type}})
        in Table.RowCount(result[Published]) = 8 and Table.RowCount(grouped) = 2 and List.Distinct(grouped[Count]) = {4},
    IntervalsStayWithinOwnShift = (result as record, sourceRows as list) as logical =>
        let
            Actual = EvaluateConsumer(result[Intervals]),
            WithinBounds = Table.RowCount(Table.SelectRows(Actual, each [TimeDate] < [ShiftStart] or [TimeDate] >= [ShiftFinish])) = 0,
            Starts = Table.SelectRows(Actual, each [Attribute] = "IntervalStart" and [Role1] = [Role]),
            Totals = Table.Group(Starts, {"SourceShiftID"}, {{"ElapsedHours", each List.Sum([IntervalDuration]) * 24, type number}}),
            OriginalHours = List.Transform(sourceRows, each Duration.TotalHours([EndDateTime] - [StartDateTime]))
        in WithinBounds and Table.RowCount(Totals) = 2 and List.Sort(Totals[ElapsedHours]) = List.Sort(OriginalHours),
    AlterFirstKey = (key as text, replacement as any) as table => Table.FromRecords({Record.TransformFields(CheckRecords{0}, {{key, each replacement}})} & List.Skip(CheckRecords)),
    Check = (name as text, run as function, predicate as function) as record =>
        let
            Attempt = try run(),
            Test = if Attempt[HasError] then [HasError = true, Error = Attempt[Error]] else try predicate(Attempt[Value]),
            Passed = not Test[HasError] and Test[Value] = true,
            Details = if Test[HasError] then try Test[Error][Message] otherwise "Evaluation failed"
                else if Passed then null else "Unexpected fixture result"
        in [Case = name, Passed = Passed, Details = Details],
    Cases = {
        Check("Five source shifts retain exactly four plotting points each", () => Current[Points], each Table.RowCount(_) = 20),
        Check("All eleven legacy plotting columns preserve exact baseline values", () => Current[Points], each EqualProjection(Baseline[Points], _, PointColumns, PointSort)),
        Check("SourceShiftID and checks copied to all four points", () => {1..4}, each List.AllTrue(List.Transform(_, (index) => Table.RowCount(SourcePoints("S" & Text.From(index))) = 4))),
        Check("Different shifts of same employee retain independent statuses", () => {SourcePoints("S1"), SourcePoints("S2")}, each
            List.Distinct(_{0}[DailyStatus]) = {"Flag"} and List.Distinct(_{1}[DailyStatus]) = {"Pass"}
            and List.Distinct(_{0}[DailyExcessHours]) = {1.5} and List.Distinct(_{1}[DailyExcessHours]) = {0}),
        Check("Same-time different employees retain independent status", () => SourcePoints("S4"), each List.Distinct(_[CoverageStatus]) = {"Unknown"} and List.Distinct(_[DailyStatus]) = {"Pass"}),
        Check("Nullable flags are preserved without fill-down", () => SourcePoints("S3"), each List.Distinct(_[MinimumBreakFlag]) = {null} and List.Distinct(_[GapStatus]) = {"Unknown"}),
        Check("Unmatched Master row stays as four unknown plotting points", () => Table.SelectRows(Current[Points], each [Name] = "Master unmatched"), each
            Table.RowCount(_) = 4 and List.Distinct(_[OvertimeMatchStatus]) = {"Unmatched"}
            and List.Distinct(_[SourceShiftID]) = {null} and List.Distinct(_[DailyStatus]) = {"Unknown"}
            and List.Distinct(_[DailyExcessHours]) = {null}),
        Check("Ambiguous match retains one prepared row with unknown checks", () => Table.SelectRows(Ambiguous[Prepared], each [AllocationCode] = "101" and [ShiftStart] = #datetime(2026, 11, 2, 6, 0, 0)), each
            Table.RowCount(_) = 1 and _{0}[OvertimeMatchCount] = 2 and _{0}[OvertimeMatchStatus] = "Ambiguous"
            and _{0}[SourceShiftID] = null and _{0}[DailyStatus] = "Unknown" and _{0}[DailyExcessHours] = null),
        Check("Ambiguous match causes no plotting or interval row growth", () => Ambiguous, each
            Table.RowCount(_[Points]) = Table.RowCount(Baseline[Points]) and Table.RowCount(_[Intervals]) = Table.RowCount(OldIntervals)),
        Check("NAMES+INTERVALSLIST legacy keys and values preserve baseline", () => NewIntervals, each
            EqualProjection(OldIntervals, _, Table.ColumnNames(OldIntervals), IntervalSort)),
        Check("NAMES+INTERVALSLIST carries all new checks before endpoint and Role1 expansion", () => NewIntervals, each Table.HasColumns(_, Metadata)),
        Check("Interval rows preserve status by source shift without fill-down", () => {SourceIntervals("S1"), SourceIntervals("S2"), SourceIntervals("S3")}, each
            List.AllTrue(List.Transform(_, (rows) => Table.RowCount(rows) > 0))
            and List.Distinct(_{0}[DailyStatus]) = {"Flag"} and List.Distinct(_{1}[DailyStatus]) = {"Pass"}
            and List.Distinct(_{2}[MinimumBreakFlag]) = {null}),
        Check("Actual ResourceIntervalAllocation keeps all old metrics and row grain", () => NewConsumer, each
            EqualProjection(OldConsumer, _, Table.ColumnNames(OldConsumer), IntervalSort)),
        Check("Actual ResourceIntervalAllocation retains added check columns", () => NewConsumer, each Table.HasColumns(_, Metadata)),
        Check("Published OvertimeRosterChecks exposes five plotting fields and metadata", () => Current[Published], each
            List.FirstN(Table.ColumnNames(_), 5) = {"Name", "Role", "DateTime", "StaffCount", "ShiftType"}
            and Table.RowCount(_) = 20 and Table.HasColumns(_, Metadata)),
        Check("Wide source provenance is not repeated on plotting or interval rows", () => Current, each
            List.IsEmpty(List.Intersect({Table.ColumnNames(_[Points]), {"SourceFile", "SourceRow", "SourceName", "SourceDate"}}))
            and List.IsEmpty(List.Intersect({Table.ColumnNames(_[Intervals]), {"SourceFile", "SourceRow", "SourceName", "SourceDate"}}))),
        Check("Same display name overlapping distinct employees retain four canonical points each", () => Overlap, CanonicalPointCounts),
        Check("Same display name overlapping intervals retain each source shift's own bounds and elapsed hours", () => Overlap, each IntervalsStayWithinOwnShift(_, OverlapRows)),
        Check("Same display name legacy classifications remain visibly unknown", () => Overlap[Published], each
            List.Distinct(_[LegacyShiftStatus]) = {"Unknown"} and List.Distinct(_[ShiftType]) = {null}),
        Check("Same display name with identical starts causes no canonical point multiplication", () => SameStart, CanonicalPointCounts),
        Check("Same display name with identical starts retains exact original interval bounds", () => SameStart, each IntervalsStayWithinOwnShift(_, SameStartRows)),
        Check("Same-name distinct employee check results are not cross-attributed", () => Overlap[Published], each
            List.Distinct(Table.SelectRows(_, (row) => row[AllocationCode] = "301")[DailyStatus]) = {"Flag"}
            and List.Distinct(Table.SelectRows(_, (row) => row[AllocationCode] = "302")[DailyStatus]) = {"Pass"}),
        Check("Adjacent double-shift legs retain known classification", () => Adjacent[Published], each
            Table.RowCount(_) = 8 and List.Distinct(_[LegacyShiftStatus]) = {"Known"}),
        Check("Adjacent double-shift legs preserve baseline plotting values and bands", () => Adjacent[Published], each
            EqualProjection(AdjacentBaseline[Points], _, {"Name", "Role", "DateTime", "StaffCount", "ShiftType"},
                {{"Name", Order.Ascending}, {"Role", Order.Ascending}, {"DateTime", Order.Ascending}}))
    },
    KeyCases = List.Transform({
        [Key = "SourceUnit", Value = "TE"], [Key = "AllocationCode", Value = "other"], [Key = "Role", Value = "RN"],
        [Key = "ShiftStart", Value = #datetime(2026, 11, 2, 6, 1, 0)], [Key = "ShiftFinish", Value = #datetime(2026, 11, 2, 18, 1, 0)]},
        (change) => Check("Exact five-part key rejects changed " & change[Key],
            () => EvaluateNew(Inputs(AlterFirstKey(change[Key], change[Value])))[Prepared],
            each let first = Table.SelectRows(_, (row) => row[Name] = "Ava" and row[StartDateTime] = #datetime(2026, 11, 2, 6, 0, 0)){0}
                in first[OvertimeMatchCount] = 0 and first[OvertimeMatchStatus] = "Unmatched" and first[DailyStatus] = "Unknown")),
    Output = Table.FromRecords(Cases & KeyCases)
in
    Output;
