section DemandRoleMappingFixtures;

// Query: DemandRoleFixtures
// Purpose: Exercise the actual Demand Extract queries with synthetic imports and no external data access.
// Inputs: Evaluate is assembled from the authoritative M source by the native test runner.
// Output: One pass/fail row per mapping, coverage, reconciliation or path scenario.
shared DemandRoleFixtures = (Evaluate as function) as table =>
let
    Definitions = #table({"Roster Roles", "DC Role", "Role Group", "DC Category", "Direct Care %"}, {
        {"Nurse roster", "RN", "NURSE", "RN", 1},
        {"Nurse in charge", " rn ", "NURSE", "RN", 1},
        {"Assistant roster", "AIN", "CARE", "OTHER", 1},
        {"Enrolled roster", "EN", "CARE", "OTHER", 0.5},
        {"Manager roster", "CSM", "MGR", "OTHER", 0.25},
        {"Administration", null, null, "NA", null}
    }),
    Analysis = #table({"Roles", "Effort Management Analysis", "Leave Balance Analysis"}, {
        {"NURSE", true, false}, {"CARE", true, false}, {"MGR", false, true}, {"EN", true, false}
    }),
    WellbeingDefinitions = #table(Table.ColumnNames(Definitions), {
        {"Therapy Assistant", "WELLBEING", "TA", "OTHER", 0.6},
        {"Wellbeing & Lifestyle Coordinator", "WELLBEING", "WLC", "OTHER", 0.6},
        {"Wellbeing & Lifestyle Officer", "WELLBEING", "WLO", "OTHER", 0.6}
    }),
    WellbeingAnalysis = #table(Table.ColumnNames(Analysis), {
        {"TA", false, true}, {"WLC", false, true}, {"WLO", false, true}
    }),
    SourceDefinitions = {
        [Role="RN", Category="RN", Percentage=1, FTE=1],
        [Role="AIN", Category="OTHER", Percentage=1, FTE=2],
        [Role="EN", Category="OTHER", Percentage=0.5, FTE=3],
        [Role="CSM", Category="OTHER", Percentage=0.25, FTE=100]
    },
    Profile = Table.FromRecords(List.Combine(List.Transform(SourceDefinitions, (R) =>
        List.Combine(List.Transform({1..14}, (D) => List.Transform({"AM", "PM", "NS"}, (S) =>
            let
                Zero = R[Role] = "EN" and D = 3 and S = "PM",
                FTE = if Zero then 0 else R[FTE] * (if D <= 7 then 1 else 2)
            in [Facility="BD", Role=R[Role], #"DC Role"=R[Role], #"DC Category"=R[Category],
                MinuteCategory=if R[Category] = "RN" then "RN" else "OTHERS",
                #"Direct Care %"=R[Percentage], #"Week No"=30 + Number.IntegerDivide(D - 1, 7),
                FortnightWeek=1 + Number.IntegerDivide(D - 1, 7), DayOfWeek=1 + Number.Mod(D - 1, 7),
                FortnightDayIndex=D, Shift=S, HistoricalRosterFTE=FTE, RedistributedRosterFTE=FTE,
                RedistributionMatchCount=if Zero then 0 else 1, HistoricalCoverageStatus="PASS",
                ProfileAlignmentStatus=if Zero then "PASS ZERO" else "PASS", FTE=FTE]
        )))))),
    Allocation = Table.SelectRows(Profile, each [RedistributionMatchCount] = 1),
    Checks = #table({"Severity", "Check", "Facility", "Role", "Message"}, {{"Pass", "Fixture", "BD", null, "Synthetic"}}),
    Groups = {"NURSE", "CARE", "MGR", "EN"},
    PlanningCalendar = (DayCount as number) as table => Table.FromRecords(List.Combine(List.Transform(List.Numbers(0, DayCount), (D) =>
        List.Combine(List.Transform({0..2}, (S) => List.Transform(Groups, (R) =>
            [Date=Date.AddDays(#date(2026, 7, 20), D), Day=D + 1,
                Shifts={"AM", "PM", "NIGHT"}{S}, Period=D * 3 + S + 1, RolesList=R]
        )))))),
    Calendar = PlanningCalendar(28),
    Timing = Table.FromRecords(List.Combine(List.Transform(Groups, (R) => List.Transform({0..2}, (S) =>
        [Role=R, ShiftPeriod={"AM", "PM", "NIGHT"}{S}, StartTime={0.25, 0.5833333333333334, 0.875}{S},
            StartDay={0.25, 0.5833333333333334, 0.875}{S}, EndDay={0.59375, 0.8854166666666666, 0.25}{S},
            DurationOfShifts={8.25, 7.25, 9}{S}]
    )))),
    Settings = (C as table, T as table) as table => #table({"Item", "Kind", "Data"}, {
        {"ShiftDuration", "Table", #table({"ShiftDuration"}, {{7.6}})},
        {"ShiftStart", "Table", #table({"StartAM", "StartPM", "StartNIGHT"}, {{0.25, 0.5833333333333334, 0.875}})},
        {"PermutationDimensions", "Table", C}, {"ShiftPeriod", "Table", T}
    }),
    PathMappings = #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {
        {"https://fixture.sharepoint.com/sites/Care", "Z:\ResidentialCare"},
        {"https://fixture.sharepoint.com/sites/Care/Sub", "Y:\ResidentialCare"}
    }),
    PathNavigation = #table({"Item", "Kind", "Data"}, {{"CentriSyncPaths", "Table", PathMappings}}),
    Base = [Roles=Definitions, Analysis=Analysis, Allocation=Allocation, Profile=Profile, Checks=Checks,
        Settings=Settings(Calendar, Timing), PathNavigation=PathNavigation,
        PathInput=#table({"FilePath"}, {{"Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\[2-DemandExtract.xlsx]Sheet1"}})],
    Change = (Fields as record) as record => Record.Combine({Base, Fields}),
    Near = (A as number, B as number) as logical => Number.Abs(A - B) < 0.0000001,
    Reject = (Inputs as record) as logical => (try Table.RowCount(Evaluate(Inputs)[Demand]))[HasError],
    Cell = (T as table, R as text, D as number, S as text) as record =>
        Table.SelectRows(T, each [Role] = R and [Day] = D and [Shift] = S){0},
    Check = (Name as text, Work as function) as record =>
        let Result = try Work()
        in [Check=Name, Passed=if Result[HasError] then false else Result[Value] = true,
            Detail=if Result[HasError] then Result[Error][Message] else if Result[Value] = true then "Pass" else "Assertion failed"],
    Result = Evaluate(Base),
    Demand = Table.Buffer(Result[Demand]),
    ModifiedRole = (R as record, Field as text, Value as any) as record => Record.TransformFields(R, {{Field, (Old) => Value}}),
    UpdatedDefinition = (RoleName as text, Field as text, Value as any) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(Definitions), each
            if [DC Role] = RoleName then ModifiedRole(_, Field, Value) else _)),
    RoleAnalysis = (RoleName as text, Value as any) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(Analysis), each
            if [Roles] = RoleName then ModifiedRole(_, "Effort Management Analysis", Value) else _)),
    // A single published DC role can represent several groups that are all outside
    // effort management. Its hours must be excluded once, without inventing a group split.
    WellbeingProfile = Table.TransformColumns(Table.SelectRows(Profile, each [Role] = "CSM"), {
        {"Role", each "WELLBEING"}, {"DC Role", each "WELLBEING"}, {"Direct Care %", each 0.6}
    }),
    WithWellbeing = (Controls as table) as record => Change([
        Roles=Table.Combine({Definitions, WellbeingDefinitions}),
        Analysis=Table.Combine({Analysis, Controls}),
        Profile=Table.Combine({Profile, WellbeingProfile}),
        Allocation=Table.Combine({Allocation, WellbeingProfile})
    ]),
    MappingRejects = (Inputs as record) as logical => (try Evaluate(Inputs)[RoleCode]("WELLBEING"))[HasError],
    // Compare variable horizons with the original synthetic DC-role FTE cells, including
    // the uneven source weeks and sparse EN zero. Partial cycles must not be averaged.
    HorizonPasses = (DayCount as number) as logical =>
        let
            Changed = Evaluate(Change([Settings=Settings(PlanningCalendar(DayCount), Timing)])),
            Output = Changed[Demand],
            Reconciliation = Changed[Reconciliation],
            ExpectedHours = List.Sum(List.Transform(List.Numbers(0, DayCount), (D) =>
                List.Sum(Table.SelectRows(Profile, each List.Contains({"RN", "AIN", "EN"}, [Role])
                    and [FortnightDayIndex] = Number.Mod(D, 14) + 1)[FTE]) * 7.6)),
            ValidCycleCounts = List.AllTrue(List.Transform(Table.ToRecords(Reconciliation), each
                [CellCount] = List.Min({14, DayCount - ([PlanningFortnight] - 1) * 14}) * 3
                    and [CellCount] = [ExpectedCellCount] and [Status] = "Pass"))
        in Table.RowCount(Output) = 2 * DayCount * 3
            and List.Sort(List.Distinct(Output[Date])) = List.Dates(#date(2026, 7, 20), DayCount, #duration(1, 0, 0, 0))
            and List.Sort(List.Distinct(Output[Period])) = {1..DayCount * 3}
            and Table.RowCount(Reconciliation) = 2 * Number.RoundUp(DayCount / 14)
            and ValidCycleCounts and Near(List.Sum(Output[DemandHRS]), ExpectedHours)
            and List.AllTrue(List.Transform(Changed[Checks][Severity], each _ = "Pass")),
    Rows = {
        Check("Linked groups replace hard-coded abbreviations", () => Result[Roles] = {"CARE", "NURSE"}
            and Result[RoleCode](" en ") = "CARE" and Table.RowCount(Demand) = 168),
        Check("Roster aliases do not multiply DC-role demand", () =>
            Near(Cell(Demand, "NURSE", 1, "AM")[DemandHRS], 7.6)),
        Check("DC-role groups sum roster hours after source validation", () =>
            Near(Cell(Demand, "CARE", 1, "AM")[DemandHRS], 38)
            and Near(Cell(Demand, "CARE", 8, "AM")[DemandHRS], 76)),
        Check("Productive hours retain separate source percentages", () =>
            let Pattern = Table.SelectRows(Result[Pattern], each [Role] = "CARE" and [FortnightDayIndex] = 1 and [Shift] = "AM"){0}
            in Near(Pattern[ProductiveHRS], 26.6) and Near(Pattern[DemandHRS], 38)),
        Check("Both fortnight totals reconcile directly to DC-role cells", () =>
            Table.RowCount(Result[Reconciliation]) = 4 and List.AllTrue(List.Transform(Result[Reconciliation][Status], each _ = "Pass"))),
        Check("Full pipeline check passes", () => List.AllTrue(List.Transform(Result[Checks][Severity], each _ = "Pass"))),
        Check("Calendar diagnostics show complete coverage for each enabled group", () =>
            Table.RowCount(Result[CalendarCheck]) = 2 and List.AllTrue(List.Transform(Table.ToRecords(Result[CalendarCheck]), each
                [Severity] = "Pass" and [ExpectedPeriods] = 84 and [ActualRows] = 84
                    and [MissingPeriods] = 0 and [DuplicateRows] = 0))),
        Check("Calendar diagnostics identify an enabled group absent from Settings", () =>
            let Inputs = Change([Settings=Settings(Table.SelectRows(Calendar, each [RolesList] <> "CARE"), Timing)]),
                Diagnostics = Evaluate(Inputs)[CalendarCheck],
                Missing = Table.SelectRows(Diagnostics, each [Role] = "CARE"){0}
            in Missing[Severity] = "Error" and Missing[ExpectedPeriods] = 84
                and Missing[ActualRows] = 0 and Missing[MissingPeriods] = 84 and Missing[DuplicateRows] = 0
                and Text.Contains(Missing[Message], "Settings Roles query") and Reject(Inputs)),
        Check("Calendar diagnostics name a missing period without masking it", () =>
            let Inputs = Change([Settings=Settings(Table.SelectRows(Calendar, each not ([RolesList] = "CARE" and [Period] = 2)), Timing)]),
                Missing = Table.SelectRows(Evaluate(Inputs)[CalendarCheck], each [Role] = "CARE"){0}
            in Missing[MissingPeriods] = 1 and Missing[MissingPeriodIDs] = "2"
                and Missing[DuplicateRows] = 0 and Missing[ActualRows] = 83 and Reject(Inputs)),
        Check("Calendar diagnostics distinguish duplicates from missing periods", () =>
            let Inputs = Change([Settings=Settings(Table.Combine({Calendar,
                    Table.SelectRows(Calendar, each [RolesList] = "CARE" and [Period] = 1)}), Timing)]),
                Duplicate = Table.SelectRows(Evaluate(Inputs)[CalendarCheck], each [Role] = "CARE"){0}
            in Duplicate[DuplicateRows] = 1 and Duplicate[DuplicatePeriodIDs] = "1"
                and Duplicate[MissingPeriods] = 0 and Duplicate[ActualRows] = 85 and Reject(Inputs)),
        Check("Calendar diagnostics show both errors when total row counts still match", () =>
            let Inputs = Change([Settings=Settings(Table.Combine({
                    Table.SelectRows(Calendar, each not ([RolesList] = "CARE" and [Period] = 2)),
                    Table.SelectRows(Calendar, each [RolesList] = "CARE" and [Period] = 1)}), Timing)]),
                Invalid = Table.SelectRows(Evaluate(Inputs)[CalendarCheck], each [Role] = "CARE"){0}
            in Invalid[ActualRows] = Invalid[ExpectedPeriods] and Invalid[MissingPeriodIDs] = "2"
                and Invalid[DuplicatePeriodIDs] = "1" and Invalid[Severity] = "Error" and Reject(Inputs)),
        Check("Calendar diagnostic counts follow the Settings horizon", () =>
            let Diagnostics = Evaluate(Change([Settings=Settings(PlanningCalendar(14), Timing)]))[CalendarCheck]
            in List.AllTrue(List.Transform(Table.ToRecords(Diagnostics), each [ExpectedPeriods] = 42
                and [ActualRows] = 42 and [Severity] = "Pass"))),
        Check("Calendar error names the affected role and available Settings labels", () =>
            let Inputs = Change([Settings=Settings(Table.SelectRows(Calendar, each [RolesList] <> "CARE"), Timing)]),
                Attempt = try Table.RowCount(Evaluate(Inputs)[Calendar])
            in Attempt[HasError] and Text.Contains(Attempt[Error][Message], "CARE: expected 84 periods, found 0")
                and Text.Contains(Attempt[Error][Message], "[NURSE]")
                and Text.Contains(Attempt[Error][Message], "Inspect Demand Calendar CHECK")
                and Table.RowCount(Attempt[Error][Detail]) = 1),
        Check("Publication retains twelve columns, period IDs and fractional attendance", () =>
            Table.ColumnNames(Demand) = {"Date", "Day", "Shift", "Period", "Role", "StartTime", "EndTime", "Unit",
                "Facility", "DemandFTE", "DemandHRS", "DurationOfShifts"}
            and List.Sort(List.Distinct(Demand[Period])) = {1..84}
            and List.AllTrue(List.Transform(Table.ToRecords(Demand), each Near([DemandFTE] * [DurationOfShifts], [DemandHRS])))),
        Check("Fortnight pattern repeats without weekday averaging", () =>
            Near(Cell(Demand, "CARE", 1, "AM")[DemandHRS], Cell(Demand, "CARE", 15, "AM")[DemandHRS])
            and Date.From(Cell(Demand, "NURSE", 28, "NIGHT")[EndTime]) = #date(2026, 8, 17)),
        Check("One-day Settings horizon reconciles the actual source day", () => HorizonPasses(1)),
        Check("Seven-day Settings horizon retains the first source week", () => HorizonPasses(7)),
        Check("Fourteen-day Settings horizon reconciles one complete cycle", () => HorizonPasses(14)),
        Check("Fifteen-day Settings horizon reconciles the partial second cycle", () => HorizonPasses(15)),
        Check("Twenty-one-day Settings horizon preserves unequal source weeks", () => HorizonPasses(21)),
        Check("Thirty-five-day Settings horizon reconciles a partial third cycle", () => HorizonPasses(35)),
        Check("Forty-two-day Settings horizon reconciles three complete cycles", () => HorizonPasses(42)),
        Check("Missing internal Settings date still fails", () => Reject(Change([
            Settings=Settings(Table.SelectRows(Calendar, each [Day] <> 7), Timing)]))),
        Check("Missing shift across all roles still fails", () => Reject(Change([
            Settings=Settings(Table.SelectRows(Calendar, each [Shifts] <> "NIGHT"), Timing)]))),
        Check("Excluded groups preserve the expected horizon when enabled roles lose a date", () => Reject(Change([
            Settings=Settings(Table.SelectRows(Calendar, each [Day] <> 28 or not List.Contains({"CARE", "NURSE"}, [RolesList])), Timing)]))),
        Check("Duplicate role period still fails", () => Reject(Change([
            Settings=Settings(Table.Combine({Calendar, Table.FirstN(Calendar, 1)}), Timing)]))),
        Check("One missing role shift cannot be replaced by a duplicate", () => Reject(Change([
            Settings=Settings(Table.Combine({Table.SelectRows(Calendar, each not ([RolesList] = "CARE" and [Period] = 2)),
                Table.SelectRows(Calendar, each [RolesList] = "CARE" and [Period] = 1)}), Timing)]))),
        Check("A role with a different shared period ID fails", () => Reject(Change([
            Settings=Settings(Table.FromRecords(List.Transform(Table.ToRecords(Calendar), each
                if [RolesList] = "CARE" and [Period] = 1 then ModifiedRole(_, "Period", 2) else _)), Timing)]))),
        Check("A shared period ID cannot represent two shifts", () => Reject(Change([
            Settings=Settings(Table.TransformColumns(Calendar, {{"Period", each if _ = 2 then 1 else _}}), Timing)]))),
        Check("Null Settings date fails", () => Reject(Change([
            Settings=Settings(Table.TransformColumns(Calendar, {{"Date", each null}}), Timing)]))),
        Check("Null Settings day fails", () => Reject(Change([
            Settings=Settings(Table.TransformColumns(Calendar, {{"Day", each null}}), Timing)]))),
        Check("Null Settings period fails", () => Reject(Change([
            Settings=Settings(Table.TransformColumns(Calendar, {{"Period", each null}}), Timing)]))),
        Check("Empty Settings horizon fails", () => Reject(Change([
            Settings=Settings(Table.FirstN(Calendar, 0), Timing)]))),
        Check("Non-Monday start still fails to protect source-week alignment", () => Reject(Change([
            Settings=Settings(Table.TransformColumns(Calendar, {{"Date", each Date.AddDays(_, 1)}}), Timing)]))),
        Check("Proven sparse zero contributes zero, not a missing group", () =>
            Near(Cell(Demand, "CARE", 3, "PM")[DemandHRS], 15.2)),
        Check("False effort flag excludes despite true leave flag", () =>
            Table.RowCount(Result[Exclusions]) = 1 and Result[Exclusions]{0}[SourceRole] = "CSM"),
        Check("Null effort flag excludes", () =>
            Table.RowCount(Evaluate(Change([Analysis=RoleAnalysis("MGR", null)]))[Demand]) = 168),
        Check("TA WLC WLO sharing an excluded DC role do not block demand", () =>
            let Changed = Evaluate(WithWellbeing(WellbeingAnalysis))
            in Changed[Roles] = Result[Roles] and Table.RowCount(Changed[Demand]) = 168
                and Near(List.Sum(Changed[Demand][DemandHRS]), List.Sum(Demand[DemandHRS]))),
        Check("Shared excluded DC-role hours are counted once in exclusions", () =>
            let Changed = Evaluate(WithWellbeing(WellbeingAnalysis)),
                Excluded = Table.SelectRows(Changed[Exclusions], each [SourceRole] = "WELLBEING")
            in Table.RowCount(Excluded) = 1
                and Near(Excluded{0}[FortnightRosterHours], List.Sum(WellbeingProfile[FTE]) * 7.6)
                and Near(Excluded{0}[FortnightProductiveHours], List.Sum(WellbeingProfile[FTE]) * 7.6 * 0.6)),
        Check("Shared excluded DC role is not assigned an arbitrary group", () =>
            Evaluate(WithWellbeing(WellbeingAnalysis))[RoleCode]("WELLBEING") = null),
        Check("False and null groups can share an excluded DC role", () =>
            let Controls = Table.FromRecords(List.Transform(Table.ToRecords(WellbeingAnalysis), each
                    if [Roles] = "TA" then ModifiedRole(_, "Effort Management Analysis", null) else _))
            in Table.RowCount(Evaluate(WithWellbeing(Controls))[Demand]) = 168),
        Check("One enabled group cannot absorb excluded groups in the same DC role", () =>
            let Controls = Table.FromRecords(List.Transform(Table.ToRecords(WellbeingAnalysis), each
                    if [Roles] = "TA" then ModifiedRole(_, "Effort Management Analysis", true) else _))
            in MappingRejects(WithWellbeing(Controls))),
        Check("Several enabled groups still cannot share one DC role", () =>
            MappingRejects(WithWellbeing(Table.TransformColumns(WellbeingAnalysis, {{"Effort Management Analysis", each true}})))),
        Check("A missing analysis row is not treated as an excluded group", () =>
            MappingRejects(WithWellbeing(Table.SelectRows(WellbeingAnalysis, each [Roles] <> "TA")))),
        Check("Enabling another group dynamically increases output", () =>
            Table.RowCount(Evaluate(Change([Analysis=RoleAnalysis("MGR", true)]))[Demand]) = 252),
        Check("Disabling a group preserves remaining demand", () =>
            let Changed = Evaluate(Change([Analysis=RoleAnalysis("CARE", false)]))[Demand]
            in Table.RowCount(Changed) = 84 and Near(List.Sum(Changed[DemandHRS]), List.Sum(Table.SelectRows(Demand, each [Role] = "NURSE")[DemandHRS]))),
        Check("EN uses the linked code without forced AINC4 substitution", () =>
            let Changed = Evaluate(Change([Roles=UpdatedDefinition("EN", "Role Group", "EN")]))
            in Changed[RoleCode]("EN") = "EN" and Table.RowCount(Changed[Demand]) = 252),
        Check("Duplicate effort-analysis keys fail", () => Reject(Change([Analysis=Table.Combine({Analysis, Table.FirstN(Analysis, 1)})]))),
        Check("Text effort flags fail instead of coercing", () => Reject(Change([Analysis=RoleAnalysis("CARE", "true")]))),
        Check("Missing effort-analysis row fails", () => Reject(Change([Analysis=Table.SelectRows(Analysis, each [Roles] <> "CARE")]))),
        Check("No enabled groups fails explicitly", () => Reject(Change([Analysis=Table.TransformColumns(Analysis, {{"Effort Management Analysis", each false}})]))),
        Check("One DC role cannot map to competing groups", () =>
            Reject(Change([Roles=Table.Combine({Definitions, #table(Table.ColumnNames(Definitions), {{"Other EN", "EN", "EN", "OTHER", 0.5}})})]))),
        Check("Missing DC role mapping fails", () => Reject(Change([Roles=Table.SelectRows(Definitions, each [DC Role] <> "EN")]))),
        Check("Blank output group fails", () => Reject(Change([Roles=UpdatedDefinition("EN", "Role Group", " ")]))),
        Check("Conflicting direct-care percentage fails", () => Reject(Change([Roles=UpdatedDefinition("EN", "Direct Care %", 0.75)]))),
        Check("Category validation uses definition, not output label", () =>
            Reject(Change([Roles=UpdatedDefinition("RN", "DC Category", "OTHER")]))),
        Check("Published DC-role joins ignore case and surrounding spaces", () =>
            let Changed = Evaluate(Change([
                Profile=Table.TransformColumns(Profile, {{"Role", each " " & Text.Lower(_) & " "}}),
                Allocation=Table.TransformColumns(Allocation, {{"Role", each Text.Lower(_)}})]))
            in Near(List.Sum(Changed[Demand][DemandHRS]), List.Sum(Demand[DemandHRS]))),
        Check("QFR schema retains explicit linked-role mapping", () =>
            let AsQFR = (T as table) as table => Table.AddColumn(Table.RemoveColumns(T, {"DC Role", "DC Category"}),
                    "QFR Category", each [MinuteCategory]),
                Changed = Evaluate(Change([Profile=AsQFR(Profile), Allocation=AsQFR(Allocation)]))
            in Near(List.Sum(Changed[Demand][DemandHRS]), List.Sum(Demand[DemandHRS]))),
        Check("A group cannot hide a missing source-role cell", () =>
            Reject(Change([Profile=Table.SelectRows(Profile, each not ([Role] = "EN" and [FortnightDayIndex] = 14 and [Shift] = "NS"))]))),
        Check("Duplicate source allocation fails", () => Reject(Change([Allocation=Table.Combine({Allocation, Table.FirstN(Allocation, 1)})]))),
        Check("Unproven missing allocation fails", () => Reject(Change([Allocation=Table.Skip(Allocation, 1)]))),
        Check("Paired FTE mismatch fails", () => Reject(Change([Allocation=Table.TransformColumns(Allocation, {{"FTE", each _ + 1}})]))),
        Check("Upstream error prevents publication", () => Reject(Change([Checks=Table.TransformColumns(Checks, {{"Severity", each "Error"}})]))),
        Check("Enabled group missing from profile fails", () => Reject(Change([
            Profile=Table.SelectRows(Profile, each [Role] <> "RN"), Allocation=Table.SelectRows(Allocation, each [Role] <> "RN")]))),
        Check("Missing group calendar fails", () => Reject(Change([
            Settings=Settings(Table.SelectRows(Calendar, each [RolesList] <> "CARE"), Timing)]))),
        Check("Missing group timing fails", () => Reject(Change([
            Settings=Settings(Calendar, Table.SelectRows(Timing, each [Role] <> "CARE"))]))),
        Check("Duplicate timing fails", () => Reject(Change([
            Settings=Settings(Calendar, Table.Combine({Timing, Table.FirstN(Timing, 1)}))]))),
        Check("Other facilities do not enter modeled demand", () =>
            let Changed = Evaluate(Change([
                Profile=Table.Combine({Profile, Table.TransformColumns(Profile, {{"Facility", each "OTHER"}})}),
                Allocation=Table.Combine({Allocation, Table.TransformColumns(Allocation, {{"Facility", each "OTHER"}})})]))
            in Near(List.Sum(Changed[Demand][DemandHRS]), List.Sum(Demand[DemandHRS]))),
        Check("Local CELL filename resolves the unit root", () => Result[UnitPath] = "Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1"),
        Check("SharePoint Shared Documents path resolves", () => Evaluate(Change([
            PathInput=#table({"FilePath"}, {{"https://FIXTURE.sharepoint.com/sites/Care/Shared Documents/CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/[2-DemandExtract.xlsx]Path"}})
        ]))[UnitPath] = "Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1"),
        Check("Longest matching SharePoint prefix wins", () => Evaluate(Change([
            PathInput=#table({"Column1"}, {{"https://fixture.sharepoint.com/sites/Care/Sub/CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx"}})
        ]))[UnitPath] = "Y:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1"),
        Check("Resolver rejects partial-prefix match", () => (try Evaluate(Change([
            PathInput=#table({"FilePath"}, {{"https://fixture.sharepoint.com/sites/Careful/CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx"}})
        ]))[UnitPath])[HasError]),
        Check("Resolver rejects another workbook", () => (try Evaluate(Change([
            PathInput=#table({"FilePath"}, {{"Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\Other.xlsx"}})
        ]))[UnitPath])[HasError])
    }
in
    Table.FromRecords(Rows);
