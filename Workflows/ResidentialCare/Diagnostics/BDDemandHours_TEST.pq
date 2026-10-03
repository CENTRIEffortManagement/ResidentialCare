section BDDemandHoursFixtures;

// Query: BDDemandHoursFixtures
// Purpose: Evaluate the three approved native M sources with synthetic demand and timing inputs.
// Inputs: Evaluators are assembled by scripts/test-bd-demand-hours-native.ps1 with external access disabled.
// Output: One assertion per hour-conservation, reporting-independence or validation scenario.
// Notes: Master begins at prepared historical cells and target inputs; its distribution and publication gates are real queries.
shared BDDemandHoursFixtures = (Master as function, Extract as function, Intervals as function) as table =>
let
    Near = (A as number, B as number) as logical => Number.Abs(A - B) <= 0.0000001,
    Check = (Name as text, Work as function) as record =>
        let Result = try Work()
        in [Check=Name, Passed=if Result[HasError] then false else Result[Value] = true,
            Detail=if Result[HasError] then Result[Error][Message] else if Result[Value] = true then "Pass" else "Assertion failed"],
    Change = (Base as record, Fields as record) as record => Record.Combine({Base, Fields}),
    Set = (Row as record, Field as text, Value as any) as record => Record.TransformFields(Row, {{Field, each Value}}),
    UpdateFirst = (T as table, Field as text, Value as any) as table =>
        Table.FromRecords({Set(T{0}, Field, Value)} & List.Skip(Table.ToRecords(T))),
    Roles = #table({"Roster Roles", "DC Role", "Role Group", "DC Category", "Direct Care %"}, {
        {"Nurse roster", "RN", "NURSE", "RN", 1}, {"Nurse in charge", " rn ", "NURSE", "RN", 1},
        {"Assistant roster", "AIN", "CARE", "OTHER", 1}, {"Enrolled roster", "EN", "CARE", "OTHER", 0.5},
        {"Manager roster", "CSM", "MGR", "OTHER", 0.25}
    }),
    Analysis = #table({"Roles", "Effort Management Analysis", "Leave Balance Analysis"}, {
        {"NURSE", true, false}, {"CARE", true, false}, {"MGR", false, true}
    }),
    RoleValues = {[Role="RN", Category="RN", Percentage=1, Hours=7.6],
        [Role="AIN", Category="OTHERS", Percentage=1, Hours=15.2],
        [Role="EN", Category="OTHERS", Percentage=0.5, Hours=22.8],
        [Role="CSM", Category="OTHERS", Percentage=0.25, Hours=760]},
    History = Table.Buffer(Table.FromRecords(List.Combine(List.Transform(RoleValues, (R) =>
        List.Combine(List.Transform({1..14}, (D) => List.Transform({0..2}, (S) =>
            let H = if R[Role] = "EN" and D = 3 and S = 1 then 0 else R[Hours] * (if D <= 7 then 1 else 2),
                W = 1 + Number.IntegerDivide(D - 1, 7),
                Weekday = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}{Number.Mod(D - 1, 7)}
            in [Facility="BD", MinuteCategory=R[Category], Role=R[Role], #"DC Role"=R[Role],
                #"DC Category"=if R[Category] = "RN" then "RN" else "OTHER", #"Direct Care %"=R[Percentage],
                #"Week No"=29 + W, FortnightWeek=W, DayOfWeek=Number.Mod(D - 1, 7) + 1, #"Week Day"=Weekday,
                FortnightDayIndex=D, FortnightDay=Text.From(W) & "-" & Weekday,
                Shift={"AM", "PM", "NS"}{S}, ShiftIndex=S + 1, HistoricalRosterHours=H,
                HistoricalRosterFTE=H / 7.6, HistoricalProductiveHours=H * R[Percentage], HistoricalCoverageStatus="PASS"]
        ))))))),
    CategoryTotals = Table.Group(History, {"Facility", "MinuteCategory"}, {{"Hours", each List.Sum([HistoricalProductiveHours]), type number}}),
    AllTargetMinutes = List.Sum(CategoryTotals[Hours]) * 1.25 * 60,
    Targets = Table.FromRecords(List.Transform(Table.ToRecords(CategoryTotals), each
        [Facility=[Facility], MinuteCategory=[MinuteCategory], CategoryTargetMinutes=[Hours] * 1.25 * 60,
            CategoryDailyTargetMinutes=[Hours] * 1.25 * 60 / 14, ALLFortnightTargetMinutes=AllTargetMinutes])),
    SourceTargets = #table({"Facility", "MinuteType", "TargetFacilityMappingCount", "TargetFortnightHours", "TargetFacilityKey"}, {
        {"BD", "RN", 1, Table.SelectRows(CategoryTotals, each [MinuteCategory] = "RN"){0}[Hours] * 1.25, "BD"},
        {"BD", "ALL", 1, AllTargetMinutes / 60, "BD"}
    }),
    MasterInputs = [History=History, Targets=Targets, SourceTargets=SourceTargets],
    MasterResult = Master(MasterInputs),
    Distribution = Table.Buffer(MasterResult[Distribution]),
    Profile = Table.Buffer(MasterResult[Profile]),
    MasterChecks = Table.Buffer(MasterResult[Checks]),
    Navigation = (D as table, P as table, C as table) as table => #table({"Item", "Kind", "Data"}, {
        {"MinuteWorkersFTE_TABLE", "Table", D}, {"MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE", "Table", P},
        {"MinuteWorkersFTE_CHECK", "Table", C}
    }),
    Groups = {"NURSE", "CARE", "MGR"},
    Calendar = (Days as number) as table => Table.FromRecords(List.Combine(List.Transform(List.Numbers(0, Days), (D) =>
        List.Combine(List.Transform({0..2}, (S) => List.Transform(Groups, (R) =>
            [Date=Date.AddDays(#date(2026, 7, 20), D), Day=D + 1, Shifts={"AM", "PM", "NIGHT"}{S},
                Period=D * 3 + S + 1, RolesList=R]
        )))))),
    Timing = Table.FromRecords(List.Combine(List.Transform(Groups, (R) => List.Transform({0..2}, (S) =>
        [Role=R, ShiftPeriod={"AM", "PM", "NIGHT"}{S}, StartTime={0.25, 14 / 24, 21 / 24}{S},
            StartDay={0.25, 14 / 24, 21 / 24}{S}, EndDay={14.25 / 24, 21.25 / 24, 0.25}{S},
            DurationOfShifts={8.25, 7.25, 9}{S}]
    )))),
    Settings = (Days as number, Standard as any) as table => #table({"Item", "Kind", "Data"}, {
        {"ShiftDuration", "Table", #table({"ShiftDuration"}, {{Standard}})},
        {"ShiftStart", "Table", #table({"StartAM", "StartPM", "StartNIGHT"}, {{0.25, 14 / 24, 21 / 24}})},
        {"PermutationDimensions", "Table", Calendar(Days)}, {"ShiftPeriod", "Table", Timing}
    }),
    ExtractInputs = [Roles=Roles, Analysis=Analysis,
        DemandNavigation=Navigation(Distribution, Profile, MasterChecks), Settings=Settings(15, 7.6),
        PathNavigation=#table({"Item", "Kind", "Data"}, {{"CentriSyncPaths", "Table",
            #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {{"https://fixture.invalid/Care", "Z:\ResidentialCare"}})}}),
        PathInput=#table({"FilePath"}, {{"Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\1. Input\[2-DemandExtract.xlsx]Path"}})],
    ExtractWith = (D as table, P as table, C as table) as record =>
        Extract(Change(ExtractInputs, [DemandNavigation=Navigation(D, P, C)])),
    ExtractResult = Extract(ExtractInputs),
    Demand = Table.Buffer(ExtractResult[Demand]),
    Cell = (T as table, R as text, D as number, S as text) as record =>
        Table.SelectRows(T, each [Role] = R and [Day] = D and [Shift] = S){0},
    RejectExtract = (D as table, P as table, C as table) as logical =>
        (try Table.RowCount(Table.Buffer(ExtractWith(D, P, C)[Demand])))[HasError],
    RejectMaster = (Fields as record) as logical =>
        (try Table.RowCount(Table.Buffer(Master(Change(MasterInputs, Fields))[Distribution])))[HasError],
    SameDemand = (Other as table) as logical => Table.SelectColumns(Other, {"Period", "Role", "DemandHRS"}) =
        Table.SelectColumns(Demand, {"Period", "Role", "DemandHRS"}),
    HorizonPasses = (Days as number) as logical =>
        let Changed = Extract(Change(ExtractInputs, [Settings=Settings(Days, 7.6)])),
            Expected = List.Sum(List.Transform(List.Numbers(0, Days), (D) =>
                List.Sum(Table.SelectRows(Profile, each List.Contains({"RN", "AIN", "EN"}, [Role])
                    and [FortnightDayIndex] = Number.Mod(D, 14) + 1)[RedistributedDemandHours])))
        in Table.RowCount(Changed[Demand]) = Days * 3 * 2 and Near(List.Sum(Changed[Demand][DemandHRS]), Expected)
            and List.AllTrue(List.Transform(Changed[Reconciliation][Status], each _ = "Pass")),
    // Quarter-hour boundaries include all three unequal shifts and the next-day night-shift endpoint.
    MakeIntervalInputs = (D as table) as record =>
        let First = DateTime.From(List.Min(D[Date])), Last = List.Max(D[EndTime]),
            Count = Number.Round(Duration.TotalHours(Last - First) * 4),
            IDs = Table.FromRecords(List.Transform({0..Count}, each
                [DateTime=First + #duration(0, 0, 15 * _, 0),
                    DayDate=Date.From(First + #duration(0, 0, 15 * _, 0)), IntervalID=_ + 1])),
            Rows = Table.FromRecords(List.Combine(List.Transform(List.Distinct(D[Role]), (R) =>
                List.Transform({0..Count - 1}, (I) => [Role=R, IntervalID=I + 1,
                    StartInterval=First + #duration(0, 0, 15 * I, 0),
                    EndInterval=First + #duration(0, 0, 15 * (I + 1) - 1, 0),
                    EndPrecise=First + #duration(0, 0, 15 * (I + 1), 0), Duration=1 / 96, ShiftPeriod="Fixture"]))))
        in [Demand=D, IntervalIDs=IDs, Intervals=Rows],
    IntervalInputs = MakeIntervalInputs(Demand),
    IntervalResult = Intervals(IntervalInputs),
    IntervalDemand = Table.Buffer(IntervalResult[Intervals]),
    SmallDemand = Table.FirstN(Demand, 1),
    SmallInputs = MakeIntervalInputs(SmallDemand),
    ActiveIntervals = Table.SelectRows(SmallInputs[Intervals], each [StartInterval] >= SmallDemand{0}[StartTime]
        and [EndPrecise] <= SmallDemand{0}[EndTime]),
    FirstActiveID = ActiveIntervals{0}[IntervalID],
    EditActive = (Field as text, F as function) as table => Table.FromRecords(List.Transform(
        Table.ToRecords(SmallInputs[Intervals]), each if [IntervalID] = FirstActiveID then Set(_, Field, F(Record.Field(_, Field))) else _)),
    RejectIntervals = (Fields as record) as logical =>
        (try Table.RowCount(Table.Buffer(Intervals(Change(SmallInputs, Fields))[Intervals])))[HasError],
    MinuteStart = SmallDemand{0}[StartTime],
    MinuteEnd = MinuteStart + #duration(0, 0, 1, 0),
    MinuteDemand = UpdateFirst(UpdateFirst(SmallDemand, "EndTime", MinuteEnd), "DurationOfShifts", 1 / 60),
    MinuteInputs = [Demand=MinuteDemand,
        IntervalIDs=#table({"DateTime", "DayDate", "IntervalID"}, {{MinuteStart, Date.From(MinuteStart), 1}, {MinuteEnd, Date.From(MinuteEnd), 2}}),
        Intervals=#table({"Role", "IntervalID", "StartInterval", "EndInterval", "EndPrecise", "Duration", "ShiftPeriod"},
            {{SmallDemand{0}[Role], 1, MinuteStart, MinuteStart, MinuteEnd, 1 / 1440, "Fixture"}})],
    ExpectedIntervalColumns = {"Shift", "Role", "Facility", "StartTime", "EndTime", "Date", "DateTime", "DayDate", "DateTime.1", "IntervalListx",
        "StartInterval", "EndInterval", "Duration", "ShiftPeriod", "Unit", "DemandFTE", "DemandHRS", "ShiftDuration",
        "EffectiveDuration", "IntervalEffectiveRatio", "EffectiveIntervalAttendance", "DemandEffort"},
    Rows = {
        Check("Master publishes minutes divided by 60 and preserves legacy display FTE", () =>
            List.AllTrue(List.Transform(Table.ToRecords(Distribution), each Near([RedistributedDemandHours], [WeekdayShiftRosterMinutes] / 60)
                and Near([FTE], [WeekdayShiftRosterMinutes] / 456)))
            and List.AllTrue(List.Transform(Table.ToRecords(Profile), each Near([RedistributedDemandHours], [HistoricalRosterHours] * 1.25)
                and Near([RedistributedRosterFTE], [HistoricalRosterFTE] * 1.25)))),
        Check("Actual Master hour, daily and profile publication gates pass", () =>
            Table.RowCount(MasterChecks) > 0 and List.AllTrue(List.Transform(MasterChecks[Severity], each _ = "Pass"))),
        Check("Master retains distinct unequal weeks and proven sparse zero", () =>
            Table.RowCount(Distribution) = 167 and Table.RowCount(Profile) = 168
                and Table.RowCount(Table.SelectRows(Profile, each [ProfileAlignmentStatus] = "PASS ZERO" and [RedistributedDemandHours] = 0)) = 1),
        Check("Master display FTE changes cannot alter authoritative hours", () =>
            let Changed = Master(Change(MasterInputs, [History=Table.TransformColumns(History, {{"HistoricalRosterFTE", each 999}}),
                DistributionTransform=(T) => Table.TransformColumns(T, {{"FTE", each 12345}})]))
            in Changed[Distribution][RedistributedDemandHours] = Distribution[RedistributedDemandHours]
                and Changed[Profile][RedistributedDemandHours] = Profile[RedistributedDemandHours]
                and List.AllTrue(List.Transform(Changed[Checks][Severity], each _ = "Pass"))),
        Check("Master missing positive-demand cell blocks publication", () => RejectMaster([DistributionTransform=(T) => Table.Skip(T, 1)])),
        Check("Master duplicated demand cell blocks publication", () => RejectMaster([DistributionTransform=(T) => Table.Combine({T, Table.FirstN(T, 1)})])),
        Check("Master changed hours block publication", () => RejectMaster([DistributionTransform=(T) => UpdateFirst(T, "RedistributedDemandHours", 999)])),
        Check("Master null hours block publication", () => RejectMaster([DistributionTransform=(T) => UpdateFirst(T, "RedistributedDemandHours", null)])),
        Check("Master target-hour mismatch blocks publication", () => RejectMaster([SourceTargets=UpdateFirst(SourceTargets, "TargetFortnightHours", 999)])),
        Check("Master-to-Extract totals reconcile by source role with exclusions", () => HorizonPasses(15)
            and ExtractResult[Roles] = {"CARE", "NURSE"} and Table.RowCount(ExtractResult[Exclusions]) = 1
            and ExtractResult[Exclusions]{0}[SourceRole] = "CSM"
            and Near(ExtractResult[Exclusions]{0}[FortnightRosterHours], List.Sum(Table.SelectRows(Profile, each [Role] = "CSM")[RedistributedDemandHours]))),
        Check("Grouped roles preserve per-source Direct Care percentages", () =>
            let Pattern = Table.SelectRows(ExtractResult[Pattern], each [Role] = "CARE" and [FortnightDayIndex] = 1 and [Shift] = "AM"){0}
            in Near(Pattern[DemandHRS], 47.5) and Near(Pattern[ProductiveHRS], 33.25)),
        Check("Sparse zero is preserved without dropping its role group", () => Near(Cell(Demand, "CARE", 3, "PM")[DemandHRS], 19)),
        Check("Fractional attendance uses actual shift duration with no additional meal deduction", () =>
            Near(Cell(Demand, "NURSE", 1, "AM")[DemandHRS], 9.5)
                and Near(Cell(Demand, "NURSE", 1, "AM")[DemandFTE], 9.5 / 8.25)),
        Check("Extract retains the twelve-column publication interface", () => Table.ColumnNames(Demand) =
            {"Date", "Day", "Shift", "Period", "Role", "StartTime", "EndTime", "Unit", "Facility", "DemandFTE", "DemandHRS", "DurationOfShifts"}),
        Check("Published display FTE is not an Extract input", () => SameDemand(ExtractWith(
            Table.TransformColumns(Distribution, {{"FTE", each 999}}),
            Table.TransformColumns(Profile, {{"RedistributedRosterFTE", each -99}, {"HistoricalRosterFTE", each null}}), MasterChecks)[Demand])),
        Check("Absent published display FTE does not prevent Extract output", () => SameDemand(ExtractWith(
            Table.RemoveColumns(Distribution, {"FTE"}), Table.RemoveColumns(Profile, {"RedistributedRosterFTE", "HistoricalRosterFTE"}), MasterChecks)[Demand])),
        Check("7.5, 7.6 and 8 standard hours cannot alter Extract hours", () => List.AllTrue(List.Transform({7.5, 7.6, 8},
            each SameDemand(Extract(Change(ExtractInputs, [Settings=Settings(15, _)]))[Demand])))),
        Check("Missing standard duration cannot alter Extract hours", () => SameDemand(Extract(Change(ExtractInputs,
            [Settings=Table.SelectRows(Settings(15, null), each [Item] <> "ShiftDuration")]))[Demand])),
        Check("Missing distribution hours fail without an FTE fallback", () => RejectExtract(Table.RemoveColumns(Distribution, {"RedistributedDemandHours"}), Profile, MasterChecks)),
        Check("Missing profile hours fail without an FTE fallback", () => RejectExtract(Distribution, Table.RemoveColumns(Profile, {"RedistributedDemandHours"}), MasterChecks)),
        Check("Missing historical hours cannot prove a sparse zero", () => RejectExtract(Distribution, Table.RemoveColumns(Profile, {"HistoricalRosterHours"}), MasterChecks)),
        Check("Invalid authoritative hour values fail", () => List.AllTrue(List.Transform({null, "7.6", -1, #nan, #infinity},
            each RejectExtract(UpdateFirst(Distribution, "RedistributedDemandHours", _), Profile, MasterChecks)))),
        Check("Paired distribution/profile hours must agree", () => RejectExtract(UpdateFirst(Distribution, "RedistributedDemandHours", 999), Profile, MasterChecks)),
        Check("A role group cannot mask a missing source-role cell", () => RejectExtract(Distribution, Table.Skip(Profile, 1), MasterChecks)),
        Check("Duplicate distribution rows fail", () => RejectExtract(Table.Combine({Distribution, Table.FirstN(Distribution, 1)}), Profile, MasterChecks)),
        Check("Unproven missing distribution row fails", () => RejectExtract(Table.Skip(Distribution, 1), Profile, MasterChecks)),
        Check("Saved Master error rows block Extract publication", () => RejectExtract(Distribution, Profile, UpdateFirst(MasterChecks, "Severity", "Error"))),
        Check("One, seven, fourteen, twenty-one and thirty-five day horizons retain actual pattern", () =>
            List.AllTrue(List.Transform({1, 7, 14, 21, 35}, HorizonPasses))),
        Check("Night shifts retain the next-day endpoint", () => Date.From(Cell(Demand, "NURSE", 15, "NIGHT")[EndTime]) = #date(2026, 8, 4)),
        Check("All actual interval coverage and hour reconciliation checks pass", () =>
            List.AllTrue(IntervalResult[Checks][Passed]) and List.AllTrue(IntervalResult[InputChecks][Passed])),
        Check("Master to Extract to intervals conserves total hours", () => Near(List.Sum(IntervalDemand[DemandEffort]), List.Sum(Demand[DemandHRS]))),
        Check("Interval output preserves existing columns and fractional attendance", () => Table.ColumnNames(IntervalDemand) = ExpectedIntervalColumns
            and List.AllTrue(List.Transform(Table.ToRecords(IntervalDemand), each
                Near([DemandEffort], [DemandHRS] * ([Duration] * 24) / [ShiftDuration])
                and Near([EffectiveIntervalAttendance], [DemandHRS] / [ShiftDuration])))),
        Check("Interval display FTE changes cannot alter demand effort", () =>
            let Changed = Intervals(Change(SmallInputs, [Demand=Table.TransformColumns(SmallDemand, {{"DemandFTE", each 12345}})]))[Intervals],
                Baseline = Intervals(SmallInputs)[Intervals]
            in Changed[DemandEffort] = Baseline[DemandEffort] and Changed[EffectiveIntervalAttendance] = Baseline[EffectiveIntervalAttendance]),
        Check("Zero demand hours publish zero interval effort and attendance", () =>
            let Zero = Intervals(Change(SmallInputs, [Demand=UpdateFirst(SmallDemand, "DemandHRS", 0)]))
            in List.AllTrue(Zero[Checks][Passed])
                and List.AllTrue(List.Transform(Zero[Intervals][DemandEffort], each _ = 0))
                and List.AllTrue(List.Transform(Zero[Intervals][EffectiveIntervalAttendance], each _ = 0))),
        Check("One-minute interval preserves the inclusive end marker", () =>
            let Minute = Intervals(MinuteInputs)
            in List.AllTrue(Minute[Checks][Passed]) and Table.RowCount(Minute[Intervals]) = 1
                and Minute[Intervals]{0}[EndInterval] = Minute[Intervals]{0}[StartInterval]
                and Near(Minute[Intervals]{0}[DemandEffort], SmallDemand{0}[DemandHRS])),
        Check("Final-demand start/end point interface retains both endpoints", () =>
            let Points = Intervals(SmallInputs)[Points]
            in Table.RowCount(Points) = 2 * Table.RowCount(Intervals(SmallInputs)[Intervals])
                and List.Sort(List.Distinct(Points[PointType])) = {"EndInterval", "StartInterval"}),
        Check("Duplicate imported demand rows block intervals", () => RejectIntervals([Demand=Table.Combine({SmallDemand, SmallDemand})])),
        Check("Invalid imported demand hours block intervals", () => List.AllTrue(List.Transform({null, "9.5", -1, #nan, #infinity},
            each RejectIntervals([Demand=UpdateFirst(SmallDemand, "DemandHRS", _)])))),
        Check("Nonpositive or mismatched shift duration blocks intervals", () => List.AllTrue(List.Transform({0, -1, null, 8},
            each RejectIntervals([Demand=UpdateFirst(SmallDemand, "DurationOfShifts", _)])))),
        Check("Missing start endpoint blocks intervals", () => RejectIntervals([IntervalIDs=Table.SelectRows(SmallInputs[IntervalIDs], each [DateTime] <> SmallDemand{0}[StartTime])])),
        Check("Duplicate endpoint blocks intervals", () => RejectIntervals([IntervalIDs=Table.Combine({SmallInputs[IntervalIDs], Table.SelectRows(SmallInputs[IntervalIDs], each [DateTime] = SmallDemand{0}[StartTime])})])),
        Check("Missing role interval blocks publication", () => RejectIntervals([Intervals=Table.SelectRows(SmallInputs[Intervals], each [IntervalID] <> FirstActiveID)])),
        Check("Duplicate role interval blocks publication", () => RejectIntervals([Intervals=Table.Combine({SmallInputs[Intervals], Table.FirstN(ActiveIntervals, 1)})])),
        Check("Interval gap blocks publication", () => RejectIntervals([Intervals=EditActive("EndPrecise", each _ - #duration(0, 0, 1, 0))])),
        Check("Interval overlap blocks publication", () => RejectIntervals([Intervals=EditActive("EndPrecise", each _ + #duration(0, 0, 1, 0))])),
        Check("Conserved-total gap and overlap still block publication", () => RejectIntervals([
            Intervals=Table.FromRecords(List.Transform(Table.ToRecords(SmallInputs[Intervals]), each
                if [IntervalID] = FirstActiveID + 1 then
                    Set(Set(Set(_, "StartInterval", [StartInterval] + #duration(0, 0, 1, 0)),
                        "EndPrecise", [EndPrecise] + #duration(0, 0, 1, 0)), "EndInterval", [EndInterval] + #duration(0, 0, 1, 0))
                else _))])),
        Check("Incorrect interval duration blocks publication", () => RejectIntervals([Intervals=EditActive("Duration", each _ * 2)])),
        Check("Zero interval duration blocks publication", () => RejectIntervals([Intervals=EditActive("Duration", each 0)]))
    }
in
    Table.FromRecords(Rows);
