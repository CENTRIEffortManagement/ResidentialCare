section OvertimeRosterChecksFixturesSection;

// Query: OvertimeRosterChecksFixtures
// Purpose: Exercise the actual OvertimeRosterChecks definitions with synthetic landed inputs and independent expected results.
// Inputs: Native evaluator assembled from production shared definitions with reviewed landing overrides only.
// Output: One Case/Passed/Details row per independently specified boundary, failure or reconciliation case.
// Notes: No workbook, connector, live roster, SharePoint or production-output dependencies are used here.
shared OvertimeRosterChecksFixtures = (Evaluate as function) as table =>
let
    Anchor = #date(2026, 11, 2),
    WindowStart = #datetime(2026, 11, 2, 0, 0, 0),
    WindowEnd = #datetime(2026, 11, 16, 0, 0, 0),
    RosterColumns = {"Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time",
        "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "Published", "Non Attended", "Status", "MealAvailabilityRequired", "MealStart", "MealEnd",
        "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    WorkerColumns = {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours",
        "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    AsRawSheet = (rows as list, columns as list) as table => Table.DemoteHeaders(Table.FromRecords(rows, columns, MissingField.UseNull)),
    Worker = (employee as text, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = employee, #"Position Description" = "AIN", #"Employee Status Description" = "Permanent Full-Time",
         #"Start Date" = #date(2020, 1, 1), #"Termination Date" = null, #"Contracted FN Hours" = 76,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Shift = (day as number, start as time, finish as time, net as number, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = "E1", #"Employee_Code" = "A1", #"Pay Company" = "Synthetic employer", #"Employee Roster Name" = "Synthetic worker",
         Role = "AIN", Location = "BD", Date = Date.AddDays(Anchor, day), #"Start Time" = start, #"End Time" = finish,
         #"Break Length Minutes" = 0, #"Shift Net Length" = net, #"Shift Type" = null, #"Employment Type" = "Permanent Full-Time",
         Published = true, #"Non Attended" = false, Status = "Published",
         MealAvailabilityRequired = false, MealStart = null, MealEnd = null, CoverageStart = WindowStart, CoverageEnd = WindowEnd,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Carrier = (unit as text) as record => Shift(0, #time(1, 0, 0), #time(2, 0, 0), 1,
        [#"Employee Code" = "C" & unit, #"Employee_Code" = "AC" & unit, #"Employee Roster Name" = "Coverage carrier " & unit, Location = unit]),
    Thresholds = #table(type table [Title = text, Value = any, Unit = text, Notes = nullable text], {
        {"DailyHoursLimit", 10, "hours", null}, {"CSHospitalityHoursLimit", 8, "hours", null}, {"CSPersonalCareHoursLimit", 10, "hours", null},
        {"MinimumBreakHours", 8, "hours", null}, {"ShortGapReviewHours", 10, "hours", null}, {"MinimumFortnightDutyFreeDays", 4, "days", null},
        {"FortnightHoursLimit", 76, "hours", null}, {"FortnightLengthDays", 14, "days", null}}),
    Periods = #table(type table [ShiftPeriod = text, StartDay = number], {{"AM", 6 / 24}, {"PM", 14 / 24}, {"NIGHT", 22 / 24}}),
    Dates = (start as date, count as number) as table => #table(type table [Date = date],
        List.Transform(List.Numbers(0, count), each {Date.AddDays(start, _)})),
    Settings = (dates as table) as table => #table(type table [Item = text, Kind = text, Data = table], {
        {"ShiftPeriod", "Table", Periods}, {"PermutationDimensions", "Table", dates}}),
    Roles = #table(type table [#"Roster Roles" = text, #"Role Group" = text], {{"AIN", "AIN"}, {"EN", "EN"}, {"RN", "RN"}, {"CS", "CS"}}),
    Facilities = #table(type table [Location = text, #"Facility-Abbrev" = text], {{"BD", "BD"}, {"TE", "TE"}, {"JH-RY", "JH-RY"}}),
    Inputs = (rows as list, optional changes as nullable record) as record =>
        let
            Base = [
                PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\OvertimeRosterChecks.xlsx"}}),
                PathMapping = #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {{"https://synthetic.invalid/residentialcare", "C:\SyntheticResidentialCare"}}),
                RosterBD = AsRawSheet(rows & {Carrier("BD")}, RosterColumns),
                RosterTE = AsRawSheet({Carrier("TE")}, RosterColumns),
                RosterJHRY = AsRawSheet({Carrier("JH-RY")}, RosterColumns),
                WorkersBD = AsRawSheet({Worker("E1"), Worker("CBD")}, WorkerColumns),
                WorkersTE = AsRawSheet({Worker("E1"), Worker("CTE")}, WorkerColumns),
                WorkersJHRY = AsRawSheet({Worker("E1"), Worker("CJH-RY")}, WorkerColumns),
                SettingsBD = Settings(Dates(Anchor, 14)), SettingsTE = Settings(Dates(Anchor, 14)), SettingsJHRY = Settings(Dates(Anchor, 14)),
                Thresholds = Thresholds, Roles = Roles, Facilities = Facilities],
            Changed = Record.Combine({Base, if changes = null then [] else changes}),
            // A worker report is a current employee fact. Keep the same synthetic
            // employee fact across unit reports unless a case explicitly varies it.
            ChangedWorkers = if changes <> null and Record.HasFields(changes, "WorkersBD") then Table.PromoteHeaders(changes[WorkersBD]) else null,
            TargetWorkerRows = if ChangedWorkers = null then {} else Table.ToRecords(Table.SelectRows(ChangedWorkers, each [Employee Code] = "E1")),
            ConsistentTE = if ChangedWorkers <> null and not Record.HasFields(changes, "WorkersTE") then
                [WorkersTE = AsRawSheet(TargetWorkerRows & {Worker("CTE")}, WorkerColumns)] else [],
            ConsistentJHRY = if ChangedWorkers <> null and not Record.HasFields(changes, "WorkersJHRY") then
                [WorkersJHRY = AsRawSheet(TargetWorkerRows & {Worker("CJH-RY")}, WorkerColumns)] else [],
            Result = Record.Combine({Changed, ConsistentTE, ConsistentJHRY})
        in Result,
    WithCoverage = (input as record, start as datetime, finish as datetime) as record =>
        let SetCoverage = (raw as table) as table => Table.DemoteHeaders(Table.TransformColumns(Table.PromoteHeaders(raw),
                {{"CoverageStart", each start}, {"CoverageEnd", each finish}})),
            Calendar = Settings(Dates(Date.From(start), Duration.Days(Date.From(finish) - Date.From(start))))
        in Record.Combine({Record.TransformFields(input, {{"RosterBD", SetCoverage}, {"RosterTE", SetCoverage}, {"RosterJHRY", SetCoverage}}),
            [SettingsBD = Calendar, SettingsTE = Calendar, SettingsJHRY = Calendar]}),
    WithRosterDatesOffset = (input as record, days as number) as record =>
        let OffsetDates = (raw as table) as table => Table.DemoteHeaders(Table.TransformColumns(Table.PromoteHeaders(raw),
                {{"Date", each Date.AddDays(Date.From(_), days)}}))
        in Record.TransformFields(input, {{"RosterBD", OffsetDates}, {"RosterTE", OffsetDates}, {"RosterJHRY", OffsetDates}}),
    ChangeThreshold = (title as text, value as any) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(Thresholds), each if [Title] = title then
            Record.Combine({_, [Value = value]}) else _), Value.Type(Thresholds)),
    Near = (actual as any, expected as number) as logical => Value.Is(actual, type number) and Number.Abs(actual - expected) < 0.00000001,
    TargetShifts = (result as record) as table => Table.SelectRows(result[Shifts], each [SourceName] = "Synthetic worker"),
    TargetFortnights = (result as record) as table => Table.SelectRows(result[Fortnights], each [EmployeeID] = "E1"),
    OneShift = (result as record) as record => TargetShifts(result){0},
    LaterShift = (result as record) as record => Table.SelectRows(TargetShifts(result), each [GapHours] <> null){0},
    Unknown = (value as any) as logical => Value.Is(value, type text) and Text.Contains(Text.Upper(value), "UNKNOWN"),
    Check = (name as text, input as record, predicate as function) as record =>
        let
            Attempt = try Evaluate(input),
            Assert = if Attempt[HasError] then null else try predicate(Attempt[Value]),
            Passed = if Attempt[HasError] then false else if Assert[HasError] then false else Assert[Value] = true,
            Details = if Attempt[HasError] then try Attempt[Error][Message] otherwise "Production evaluation error"
                else if Assert[HasError] then try Assert[Error][Message] otherwise "Assertion evaluation error"
                else if Passed then null else "Actual production result differs from the independent fixture expectation"
        in [Case = name, Passed = Passed, Details = Details],
    Reject = (name as text, input as record, force as function, expectedSetting as text) as record =>
        let
            Attempt = try force(Evaluate(input)),
            Message = if Attempt[HasError] then try Attempt[Error][Message] otherwise "Evaluation rejected" else null,
            Passed = Attempt[HasError] and Message <> null and Text.Contains(Message, expectedSetting, Comparer.OrdinalIgnoreCase)
        in [Case = name, Passed = Passed, Details = if Passed then null else if Attempt[HasError] then Message else "Invalid configuration was not rejected"],
    GapInputs = (laterDay as number, laterStart as time, laterFinish as time, laterNet as number, optional changes as nullable record) as record =>
        Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(laterDay, laterStart, laterFinish, laterNet)}, changes),
    DayShifts = (count as number) as list => List.Transform(List.Numbers(0, count), each Shift(_, #time(6, 0, 0), #time(14, 0, 0), 8)),
    CompleteTenDays = Inputs(DayShifts(10)),
    CompleteElevenDays = Inputs(DayShifts(11)),
    SeventySixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(10, 0, 0), 4)},
    SeventyFourPlusSixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(8, 0, 0), 2), Shift(10, #time(6, 0, 0), #time(12, 0, 0), 6)},
    FiftyFiveHours = Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9), Shift(2, #time(22, 0, 0), #time(6, 0, 0), 8)}),
    DailyFortnightOverlap = Inputs(DayShifts(8) & {Shift(8, #time(0, 0, 0), #time(14, 0, 0), 14)}),
    Cases = {
        Check("Daily exactly ten hours has no excess", Inputs({Shift(0, #time(6, 0, 0), #time(16, 0, 0), 10)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Daily eleven net hours has one excess hour", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and OneShift(r)[DailyHoursFlag] = true),
        Check("CS Hospitality exactly eight hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [Role = "CS", AgreementStream = "Hospitality"])},
                [WorkersBD = AsRawSheet({Worker("E1", [AgreementStream = "Hospitality"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("CS Hospitality eight and one-quarter hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 15, 0), 8.25, [Role = "CS", AgreementStream = "Hospitality"])},
                [WorkersBD = AsRawSheet({Worker("E1", [AgreementStream = "Hospitality"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Near(OneShift(r)[DailyExcessHours], 0.25) and OneShift(r)[DailyHoursFlag] = true),
        Check("Conflicting roster and worker agreement streams remain unknown", Inputs({Shift(0, #time(6, 0, 0), #time(14, 15, 0), 8.25,
                [Role = "CS", AgreementStream = "Hospitality"])}),
            (r) => Unknown(OneShift(r)[PolicyStatus]) and OneShift(r)[DailyHoursFlag] = null),
        Check("CS Personal Care uses ten-hour limit", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = "Personal Care"])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Unpaid meal is excluded once", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0)),
        Check("Required client availability during meal counts as work", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealAvailabilityRequired = true, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1)),
        Check("Known overtime clock segments never include an unpaid meal", Inputs({Shift(0, #time(6, 0, 0), #time(21, 0, 0), 14,
                [#"Break Length Minutes" = 60, MealStart = #time(17, 0, 0), MealEnd = #time(18, 0, 0)])}),
            (r) => Table.RowCount(r[Segments]) = 2 and Near(List.Sum(r[Segments][SegmentHours]), 4)
                and Near(List.Sum(List.Transform(Table.ToRecords(r[Segments]), each Duration.TotalHours([SegmentFinish] - [SegmentStart]))), 4)
                and Table.IsEmpty(Table.SelectRows(r[Segments], each [SegmentStart] < #datetime(2026, 11, 2, 18, 0, 0) and [SegmentFinish] > #datetime(2026, 11, 2, 17, 0, 0)))),
        Check("Unknown meal timing preserves known daily excess quantity", Inputs({Shift(0, #time(6, 0, 0), #time(18, 0, 0), 11,
                [#"Break Length Minutes" = 60])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and Table.RowCount(Table.SelectRows(r[Segments], each Unknown([TimingStatus]))) > 0),
        Check("Seven hours fifty-nine minutes fails minimum rest", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 7 + 59 / 60) and LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly eight hours passes minimum rest but needs short-gap review", GapInputs(0, #time(22, 0, 0), #time(23, 0, 0), 1),
            (r) => Near(LaterShift(r)[GapHours], 8) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly ten hours has no short-gap trigger", GapInputs(1, #time(0, 0, 0), #time(2, 0, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 10) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = false),
        Check("First observed shift does not invent a predecessor", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => OneShift(r)[GapHours] = null and Unknown(OneShift(r)[GapStatus])),
        Check("Cross-unit duty participates in rest chronology", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(22, 0, 0), #time(23, 0, 0), 1, [Location = "TE"])}, RosterColumns)]),
            (r) => Table.RowCount(TargetShifts(r)) = 2 and Near(LaterShift(r)[GapHours], 8)),
        Check("Casual short rest remains a roster breach", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Employment Type" = "Casual"]),
                Shift(0, #time(21, 59, 0), #time(23, 59, 0), 2, [#"Employment Type" = "Casual"])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Employee Status Description" = "Casual", #"Contracted FN Hours" = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Documented broken-shift exception affects overtime review only", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8),
                Shift(0, #time(21, 0, 0), #time(23, 0, 0), 2, [BrokenShiftException = true])}),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Four complete duty-free dates passes days-off check", CompleteTenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 4) and Near(TargetFortnights(r){0}[DaysShort], 0)),
        Check("Three complete duty-free dates flags one day short", CompleteElevenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3) and Near(TargetFortnights(r){0}[DaysShort], 1)),
        Check("Missing one unit's declared coverage leaves days off unknown", Inputs(DayShifts(10),
                [RosterTE = AsRawSheet({Record.Combine({Carrier("TE"), [CoverageStart = null, CoverageEnd = null]})}, RosterColumns)]),
            (r) => Unknown(TargetFortnights(r){0}[DaysOffStatus]) and TargetFortnights(r){0}[DutyFreeDays] = null),
        Check("Overnight duty occupies both touched calendar dates", Inputs({Shift(0, #time(22, 0, 0), #time(6, 0, 0), 8)} &
                List.Transform({2..10}, each Shift(_, #time(6, 0, 0), #time(7, 0, 0), 1))),
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3)),
        Check("Monday fifteen to Wednesday twenty-two gives fifty-five elapsed gap hours", FiftyFiveHours,
            (r) => Near(LaterShift(r)[GapHours], 55)),
        Check("Removed pre-night fields and trigger are absent from published outputs", FiftyFiveHours,
            (r) => List.IsEmpty(List.Intersect({Table.ColumnNames(r[Shifts]), {"PreviousRosterShiftPeriod", "NightTransitionFlag", "NightDutyFreeDays",
                    "NightDaysFlag", "NightGapStatus", "NightStatus", "NightTransitionStatus"}}))
                and List.AllTrue(List.Transform(r[Shifts][TriggerSummary], each not Text.Contains(_, "BeforeNight")))
                and List.AllTrue(List.Transform(r[Segments][TriggerSummary], each not Text.Contains(_, "BeforeNight")))),
        Check("Exactly seventy-six net hours has no fortnight excess", Inputs(SeventySixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 76) and Near(TargetFortnights(r){0}[FortnightExcessHours], 0)),
        Check("Seventy-four plus six identifies four excess hours", Inputs(SeventyFourPlusSixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)
                and Near(List.Sum(TargetShifts(r)[FortnightExcessHours]), 4)),
        Check("Daily overtime hours still enter the net fortnight counter once", Inputs(List.Transform({0..5}, each Shift(_, #time(0, 0, 0), #time(10, 0, 0), 10)) &
                {Shift(6, #time(0, 0, 0), #time(14, 0, 0), 14), Shift(7, #time(0, 0, 0), #time(6, 0, 0), 6)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)),
        Check("Daily and fortnight overlapping excess is counted once", DailyFortnightOverlap,
            (r) => Near(TargetFortnights(r){0}[FortnightExcessHours], 2) and Near(List.Sum(r[Segments][SegmentHours]), 4)
                and Table.RowCount(Table.SelectRows(r[Segments], each [DailyTrigger] and [FortnightTrigger])) > 0),
        Check("A known boundary-crossing shift allocates two and six hours to adjacent fortnights", WithCoverage(Inputs({Shift(13, #time(22, 0, 0), #time(6, 0, 0), 8)}),
                WindowStart, #datetime(2026, 11, 30, 0, 0, 0)),
            (r) => List.Sort(TargetFortnights(r)[NetRosterHours]) = {2, 6}),
        Check("Unknown meal placement across a fortnight boundary is not prorated", WithCoverage(Inputs({Shift(13, #time(20, 0, 0), #time(8, 0, 0), 11,
                [#"Break Length Minutes" = 60])}), WindowStart, #datetime(2026, 11, 30, 0, 0, 0)),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and Near(List.Sum(r[Segments][SegmentHours]), 1)),
        Check("Invalid equal start and finish record cannot leave a mixed fortnight clear", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(1, #time(6, 0, 0), #time(6, 0, 0), 8)}),
            (r) => List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and List.AllTrue(List.Transform(TargetFortnights(r)[DaysOffStatus], Unknown))),
        Check("Missing start and finish record cannot disappear from mixed fortnight coverage", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8),
                Shift(1, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Start Time" = null, #"End Time" = null])}),
            (r) => List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and List.AllTrue(List.Transform(TargetFortnights(r)[DaysOffStatus], Unknown))),
        Check("Exact duplicate source rows never double known net hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 8) or Unknown(TargetFortnights(r){0}[FortnightStatus])),
        Check("Overlapping cross-unit duties are explicitly unknown", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(10, 0, 0), #time(18, 0, 0), 8, [Location = "TE"])}, RosterColumns)]),
            (r) => List.AllTrue(List.Transform(TargetShifts(r)[OverlapStatus], each Unknown(_) or _ = "Overlap"))),
        Check("Conflicting worker rows cannot silently choose employment type", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [WorkersBD = AsRawSheet({Worker("E1"), Worker("E1", [#"Employee Status Description" = "Casual"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[IdentityStatus]) or Unknown(OneShift(r)[DailyStatus])),
        Check("Missing CS agreement stream remains unknown", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Position Description" = "CS", AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) and OneShift(r)[DailyHoursFlag] = null),
        Reject("Invalid nonblank daily threshold is explicitly rejected", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2,
                [Thresholds = ChangeThreshold("DailyHoursLimit", "ten")]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Reject("Duplicate threshold title is explicitly rejected", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = Table.Combine({Thresholds, #table(type table [Title = text, Value = any, Unit = text, Notes = nullable text], {{"DailyHoursLimit", 12, "hours", null}})})]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Reject("Missing required threshold title is explicitly rejected", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = Table.SelectRows(Thresholds, each [Title] <> "DailyHoursLimit")]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Check("Empty Settings dates leave fortnight unknown without a roster-date fallback", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [SettingsBD = Settings(Dates(Anchor, 0)), SettingsTE = Settings(Dates(Anchor, 0)), SettingsJHRY = Settings(Dates(Anchor, 0))]),
            (r) => r[Anchor] = null and Near(OneShift(r)[DailyExcessHours], 1) and Unknown(OneShift(r)[FortnightStatus])),
        Check("Settings minimum anchors the fortnight when every published shift starts later", WithRosterDatesOffset(
                Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}), 3),
            (r) => r[Anchor] = Anchor and Date.From(OneShift(r)[ShiftStart]) = Date.AddDays(Anchor, 3)
                and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Check("Unordered duplicate Settings dates still derive the earliest date", Inputs({Shift(3, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [SettingsBD = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)})),
                 SettingsTE = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)})),
                 SettingsJHRY = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)}))]),
            (r) => r[Anchor] = Anchor and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Check("Anchor is the minimum across unit Settings even when one unit has no dates", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [SettingsBD = Settings(Dates(Date.AddDays(Anchor, 5), 14)), SettingsTE = Settings(Dates(Anchor, 0)), SettingsJHRY = Settings(Dates(Anchor, 14))]),
            (r) => r[Anchor] = Anchor and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Reject("Invalid Settings calendar date is explicitly rejected", Inputs({},
                [SettingsBD = Settings(#table({"Date"}, {{"not-a-date"}}))]),
            (r) => r[Anchor], "Date"),
        Reject("Missing Settings calendar Date column is explicitly rejected", Inputs({},
                [SettingsBD = Settings(#table({"Day"}, {{1}}))]),
            (r) => r[Anchor], "Date"),
        Check("Missing payroll eligibility remains explicit", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11,
                [OvertimeExempt = null, Grade = null, AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [OvertimeExempt = null, Grade = null, AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) or Unknown(OneShift(r)[OvertimeStatus])),
        Check("Unresolved DST offset does not publish a known elapsed rest result", WithCoverage(Inputs({
                Shift(-30, #time(15, 0, 0), #time(23, 0, 0), 8), Shift(-29, #time(9, 0, 0), #time(11, 0, 0), 2)}),
                #datetime(2026, 10, 1, 0, 0, 0), WindowEnd),
            (r) => Unknown(Table.SelectRows(TargetShifts(r), each [GapWallHours] <> null){0}[GapStatus])
                and Table.SelectRows(TargetShifts(r), each [GapWallHours] <> null){0}[MinimumBreakFlag] = null),
        Check("Source shift keys are stable under reordered roster rows", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(2, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => List.Sort(TargetShifts(r)[SourceShiftID]) = List.Sort(TargetShifts(Evaluate(Inputs({
                Shift(2, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)})))[SourceShiftID])),
        Check("Duplicated roster publications do not create extra actual shift identities", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}, RosterColumns)]),
            (r) => List.Count(List.Distinct(TargetShifts(r)[SourceShiftID])) = 1 and
                (Near(TargetFortnights(r){0}[NetRosterHours], 8) or Unknown(TargetFortnights(r){0}[FortnightStatus]))),
        Check("Worker and allocation codes retain separate identifier meanings", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => OneShift(r)[EmployeeID] = "E1" and OneShift(r)[AllocationCode] = "A1" and not Unknown(OneShift(r)[IdentityStatus])),
        Check("Allocation code alone is not guessed to be a worker identifier", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Employee Code" = null])}),
            (r) => Unknown(Table.SelectRows(r[Shifts], each [AllocationCode] = "A1"){0}[IdentityStatus])),
        Check("Overtime exemption does not suppress the minimum-rest roster check", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [OvertimeExempt = true]),
                Shift(0, #time(21, 0, 0), #time(23, 0, 0), 2, [OvertimeExempt = true])}),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Off-role duty still occupies a target employee's calendar day", Inputs(DayShifts(10) & {
                Shift(10, #time(6, 0, 0), #time(7, 0, 0), 1, [Role = "Other role"])},
                [Roles = Table.Combine({Roles, #table(type table [#"Roster Roles" = text, #"Role Group" = text], {{"Other role", "OTHER"}})})]),
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3)),
        Check("Resolver accepts CELL filename brackets at the approved nested target", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\[OvertimeRosterChecks.xlsx]Sheet1"}})]),
            (r) => Comparer.OrdinalIgnoreCase(r[Path][ClientPath], "C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon") = 0),
        Check("Resolver uses case-insensitive longest SharePoint prefix", Inputs({}, [
                PathInput = #table({"FilePath"}, {{"https://SYNTHETIC.invalid/residentialcare/client/datexx-whiddon/2. calculations/overtimerosterchecks/overtimerosterchecks.xlsx"}}),
                PathMapping = #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {
                    {"https://synthetic.invalid/residentialcare", "C:\SyntheticResidentialCare"},
                    {"https://synthetic.invalid/residentialcare/CLIENT/DATExx-Whiddon", "C:\LongestAuthoritativeWhiddon"}})]),
            (r) => Comparer.OrdinalIgnoreCase(r[Path][ClientPath], "C:\LongestAuthoritativeWhiddon") = 0),
        Reject("Resolver refuses a false local prefix boundary", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCareCollision\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\OvertimeRosterChecks.xlsx"}})]),
            (r) => r[Path][ClientPath], "CentriSyncPaths"),
        Reject("Resolver refuses an alternative workbook filename", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\RosterChecksCopy.xlsx"}})]),
            (r) => r[Path][ClientPath], "OvertimeRosterChecks.xlsx"),
        Reject("Resolver refuses the organisation calculation parent folder", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks.xlsx"}})]),
            (r) => r[Path][ClientPath], "OvertimeRosterChecks")
    },
    Results = Table.FromRecords(Cases, type table [Case = text, Passed = logical, Details = nullable text])
in
    Results;
