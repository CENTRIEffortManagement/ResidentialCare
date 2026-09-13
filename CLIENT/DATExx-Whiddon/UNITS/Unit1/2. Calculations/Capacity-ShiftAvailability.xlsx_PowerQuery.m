// Power Query from: Capacity-ShiftAvailability.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Capacity-ShiftAvailability.xlsx
// Extracted: 2026-05-18T06:14:10.401Z

section Section1;

shared #"IMPORT MultiRolesResRolesProportion" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    ResRolesProportion_Table = Source{[Item="ResRolesProportion",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResRolesProportion_Table,{{"Name", type text}, {"AINC4", type number}, {"AIN", type number}, {"Total", type number}, {"AINC4HrsAvilPref", type number}})
in
    #"Changed Type";

// Query: StdRosterDays
// Purpose: Preserve the existing ancillary roster-day cap; this does not set effective shift hours.
shared StdRosterDays = 20 meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

// Query: AvailabilityText
// Purpose: Normalize optional identifiers and names without converting text identifiers to numbers.
shared AvailabilityText = (value as any) as nullable text =>
    let Clean = try Text.Trim(Text.From(value)) otherwise null
    in if Clean = "" then null else Clean;

// Query: AvailabilityUnion
// Purpose: Merge duplicate, overlapping and touching half-open datetime intervals.
shared AvailabilityUnion = (intervals as list) as list =>
    let
        Ordered = List.Sort(intervals, (a, b) =>
            if a[Start] < b[Start] then -1 else if a[Start] > b[Start] then 1
            else if a[End] < b[End] then -1 else if a[End] > b[End] then 1 else 0),
        Merged = List.Accumulate(Ordered, {}, (state, current) =>
            // Materialize accumulators to avoid re-traversing lazy concatenations.
            List.Buffer(if List.IsEmpty(state) then {current}
            else
                let Previous = List.Last(state)
                in if current[Start] <= Previous[End] then
                    List.RemoveLastN(state, 1) &
                        {[Start = Previous[Start], End = List.Max({Previous[End], current[End]})]}
                else state & {current}))
    in Merged;

// Query: AvailabilitySubtract
// Purpose: Remove exclusion intervals from available intervals, retaining both sides of an internal gap.
shared AvailabilitySubtract = (available as list, excluded as list) as list =>
    List.Accumulate(excluded, available, (segments, cut) =>
        List.Combine(List.Transform(segments, (segment) =>
            if cut[End] <= segment[Start] or cut[Start] >= segment[End] then {segment}
            else
                (if cut[Start] > segment[Start] then
                    {[Start = segment[Start], End = cut[Start]]} else {}) &
                (if cut[End] < segment[End] then
                    {[Start = cut[End], End = segment[End]]} else {}))));

// Query: IMPORT Combined Availabilities
// Purpose: Read dated availability and leave, temporarily restricted to Beaudesert.
// Inputs: Combined Output sheet in 1. Input/whiddon_availability_leave_extraction.xlsx.
shared #"IMPORT Combined Availabilities" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\whiddon_availability_leave_extraction.xlsx"), null, true),
    PromotedHeaders = Table.PromoteHeaders(Source{[Item = "Combined Output", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    // Temporary approved facility restriction; apply the same restriction to reconciled workers.
    FilteredBeaudesert = Table.SelectRows(PromotedHeaders, each [Site Name] = "BD: Beaudesert"),
    AddedFacility = Table.AddColumn(FilteredBeaudesert, "Facility-Abbrev", each Text.Start([Site Name], 2), type text),
    SelectedColumns = Table.SelectColumns(AddedFacility,
        {"Site Name", "Employee Name", "Payroll Code", "Department", "Date From", "Date To", "Availability or Leave Reason", "Facility-Abbrev"}),
    NormalizedIdentity = Table.TransformColumns(SelectedColumns,
        {{"Payroll Code", AvailabilityText, type nullable text}, {"Employee Name", AvailabilityText, type nullable text}}),
    IndexedRecords = Table.AddIndexColumn(NormalizedIdentity, "SourceRecord", 1, 1, Int64.Type)
// Buffer the reduced scalar source shared by name resolution, parsing and diagnostics.
in Table.Buffer(IndexedRecords);

// Query: IMPORT Reconciled Workers
// Purpose: Read the authoritative worker population from the existing local reconciliation table.
shared #"IMPORT Reconciled Workers" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\Worker Reconciliation.xlsx"), null, true),
    Workers = Table.SelectColumns(Source{[Item = "Employees_TABLE", Kind = "Table"]}[Data],
        {"EmployeeID", "Facility-Abbrev", "Role", "Employee Roster Name", "EmploymentType"}),
    Normalized = Table.TransformColumns(Workers, List.Transform(Table.ColumnNames(Workers),
        each {_, AvailabilityText, type nullable text}))
in Normalized;

// Query: IMPORT Availability Settings
// Purpose: Share the Settings workbook navigation among the shift, calendar and allowance imports.
shared #"IMPORT Availability Settings" =
    let
        Navigation = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
        Required = Table.SelectRows(Navigation, each
            ([Item] = "ShiftDuration" and List.Contains({"Table", "DefinedName"}, [Kind]))
            or (List.Contains({"ShiftPeriod", "PermutationDimensions"}, [Item]) and [Kind] = "Table")),
        // Buffer only required data tables; navigation buffering alone is shallow.
        BufferedData = Table.TransformColumns(Required, {{"Data", each Table.Buffer(_), type table}})
    in Table.Buffer(BufferedData);

// Query: EXTRACT EffectiveShiftHrs
// Purpose: Import the full-shift effective-hours allowance; never derive it from roster-day caps.
// Inputs: ShiftDuration table or named range, with one ShiftDuration value.
shared #"EXTRACT EffectiveShiftHrs" = let
    Matches = Table.SelectRows(#"IMPORT Availability Settings",
        each [Item] = "ShiftDuration" and List.Contains({"Table", "DefinedName"}, [Kind])),
    Data = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error "Expected exactly one ShiftDuration table or named range in Settings Data.",
    WithHeaders = if Table.HasColumns(Data, "ShiftDuration") then Data
        else Table.PromoteHeaders(Data, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "ShiftDuration"),
    Allowance = if List.Count(Values) = 1 then Number.From(Values{0})
        else error "ShiftDuration must contain exactly one allowance.",
    Validated = if Allowance = null then error "ShiftDuration is blank."
        else if Number.IsNaN(Allowance) or Allowance <= 0 or Allowance = #infinity then
            error "ShiftDuration must be a positive finite number."
        else Allowance
in Validated;

// Query: AvailabilityRecords
// Purpose: Parse source timestamps and apply the approved case-sensitive reason classification.
// Notes: Invalid records remain visible and cannot silently become unrestricted availability.
shared AvailabilityRecords = let
    Parsed = Table.AddColumn(#"IMPORT Combined Availabilities", "Parsed", each
        let
            StartValue = try DateTime.From([Date From]),
            EndValue = try DateTime.From([Date To]),
            Classification = try (
                if [Availability or Leave Reason] = null then error "Missing reason."
                else if Text.Contains([Availability or Leave Reason], "UNAVAIL") then false
                else if Text.Contains([Availability or Leave Reason], "AVAIL") then true
                else false),
            Start = if StartValue[HasError] then null else StartValue[Value],
            End = if EndValue[HasError] then null else EndValue[Value],
            Issue = if [Payroll Code] = null then "Missing or invalid payroll code"
                else if Classification[HasError] then "Missing or invalid availability reason"
                else if Start = null or End = null then "Missing or invalid datetime"
                else if End <= Start then "End must be later than start"
                else null
        in [Start = Start, End = End,
            Available = if Classification[HasError] then null else Classification[Value],
            Issue = Issue]),
    Expanded = Table.ExpandRecordColumn(Parsed, "Parsed", {"Start", "End", "Available", "Issue"})
// Buffer parsed scalar values reused by validation and worker interval grouping.
in Table.Buffer(Table.SelectColumns(Expanded,
    {"Payroll Code", "Facility-Abbrev", "SourceRecord", "Start", "End", "Available", "Issue"}));

// Query: ReconciledWorkers_Prepare
// Purpose: Resolve nursing roles and worker names while retaining explicit eligibility diagnostics.
// Output: One row per reconciled employee/facility/source-role, including excluded rows.
shared ReconciledWorkers_Prepare = let
    SourceNames = Table.Group(#"IMPORT Combined Availabilities", {"Payroll Code", "Facility-Abbrev"},
        {{"SourceNames", each List.Distinct(List.RemoveNulls([Employee Name])), type list}}),
    WorkerGroups = Table.Group(#"IMPORT Reconciled Workers", {"EmployeeID", "Facility-Abbrev", "Role"},
        {{"RosterNames", each List.Distinct(List.RemoveNulls([Employee Roster Name])), type list},
         {"EmploymentType", each Text.Combine(List.Distinct(List.RemoveNulls([EmploymentType])), "; "), type text}}),
    RenamedRole = Table.RenameColumns(WorkerGroups, {{"Role", "SourceRole"}}),
    MappedRole = Table.AddColumn(RenamedRole, "Role", each
        if [SourceRole] = "Registered Nurse" then "RN"
        else if [SourceRole] = "Assistant in Nursing" then "AIN"
        else if [SourceRole] = "Enrolled Nurse" then "AINC4" else null, type nullable text),
    JoinedNames = Table.NestedJoin(MappedRole, {"EmployeeID", "Facility-Abbrev"},
        SourceNames, {"Payroll Code", "Facility-Abbrev"}, "Names", JoinKind.LeftOuter),
    ResolvedNames = Table.AddColumn(JoinedNames, "NameCandidates", each
        if not List.IsEmpty([RosterNames]) then [RosterNames]
        else if Table.IsEmpty([Names]) then {} else [Names]{0}[SourceNames], type list),
    NamedWorkers = Table.AddColumn(ResolvedNames, "BaseName", each
        if List.Count([NameCandidates]) = 1 then [NameCandidates]{0} else null, type nullable text),
    Eligibility = Table.AddColumn(NamedWorkers, "Issue", each
        if [EmployeeID] = null then "Missing employee ID"
        else if [#"Facility-Abbrev"] = null then "Missing facility"
        else if [#"Facility-Abbrev"] <> "BD" then "Outside temporary BD scope"
        else if [SourceRole] = null then "Missing role"
        else if [Role] = null then "Role outside approved nursing mappings"
        else if [BaseName] = null then "Missing or ambiguous worker name"
        else null, type nullable text),
    Result = Table.RemoveColumns(Eligibility, {"Names", "RosterNames", "NameCandidates"})
// Nested name lists have been removed; reuse this scalar preparation for eligibility and diagnostics.
in Table.Buffer(Result);

// Query: ReconciledWorkers_Eligible
// Purpose: Publish the eligible BD identities with the existing multiple-role display-name convention.
// Notes: Membership is independent of available hours, so wholly unavailable workers remain identifiable.
shared ReconciledWorkers_Eligible = let
    Included = Table.SelectRows(ReconciledWorkers_Prepare, each [Issue] = null),
    UniqueWorkers = Table.Distinct(Table.SelectColumns(Included,
        {"EmployeeID", "Facility-Abbrev", "Role", "BaseName", "EmploymentType"})),
    RoleCounts = Table.Group(UniqueWorkers, {"EmployeeID", "Facility-Abbrev"},
        {{"RoleCount", each List.Count(List.Distinct([Role])), Int64.Type}}),
    JoinedCounts = Table.NestedJoin(UniqueWorkers, {"EmployeeID", "Facility-Abbrev"},
        RoleCounts, {"EmployeeID", "Facility-Abbrev"}, "Roles", JoinKind.LeftOuter),
    // Use a distinct name list rather than expanding a join that could multiply worker rows.
    MultiRoleNames = List.Buffer(List.Distinct(List.RemoveNulls(
        List.Transform(#"IMPORT MultiRolesResRolesProportion"[Name], AvailabilityText)))),
    // StaffListMaster deduplicates by Name: every reconciled multi-role worker needs the same suffix convention.
    DisplayNames = Table.AddColumn(JoinedCounts, "Name", each
        if [Roles]{0}[RoleCount] > 1 or List.Contains(MultiRoleNames, [BaseName]) then [BaseName] & " (" & [Role] & ")"
        else [BaseName], type text)
in Table.Buffer(Table.RemoveColumns(DisplayNames, {"Roles"}));

// Query: WorkerAvailability_DIAGNOSTICS
// Purpose: Show excluded reconciliation rows and extraction workers absent from the selected register.
shared WorkerAvailability_DIAGNOSTICS = let
    Excluded = Table.SelectRows(ReconciledWorkers_Prepare, each [Issue] <> null),
    SourceWorkers = Table.Distinct(Table.SelectColumns(#"IMPORT Combined Availabilities", {"Payroll Code", "Facility-Abbrev"})),
    NotRegistered = Table.NestedJoin(SourceWorkers, {"Payroll Code", "Facility-Abbrev"},
        #"IMPORT Reconciled Workers", {"EmployeeID", "Facility-Abbrev"}, "Register", JoinKind.LeftAnti),
    MissingWorkers = Table.RenameColumns(Table.RemoveColumns(NotRegistered, {"Register"}), {{"Payroll Code", "EmployeeID"}}),
    MarkedMissing = Table.AddColumn(MissingWorkers, "Issue", each "Extraction worker absent from reconciliation", type text)
in Table.Combine({Excluded, MarkedMissing});

// Query: AvailabilityRecords_Invalid
// Purpose: Identify invalid intervals for included workers, plus records without a usable worker ID.
shared AvailabilityRecords_Invalid = let
    Invalid = Table.SelectRows(AvailabilityRecords, each [Issue] <> null),
    WorkerKeys = Table.Distinct(Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"})),
    Matched = Table.NestedJoin(Invalid, {"Payroll Code", "Facility-Abbrev"},
        WorkerKeys, {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.LeftOuter),
    RequiredFailures = Table.SelectRows(Matched, each [Payroll Code] = null or not Table.IsEmpty([Worker]))
in Table.RemoveColumns(RequiredFailures, {"Worker"});

// Query: AvailabilityShiftDefinitions
// Purpose: Validate the role-specific shift endpoints and durations from Settings Data.
// Notes: StartDay/EndDay are fractions of a day; EndTimeCheck is a clock time; StartTime is hours.
shared AvailabilityShiftDefinitions = let
    Data = #"IMPORT Availability Settings"{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    Selected = Table.SelectColumns(Data, {"Role", "ShiftPeriod", "StartTime", "StartDay", "DurationOfShifts", "EndDay", "EndTimeCheck"}),
    NursingRoles = Table.SelectRows(Selected, each List.Contains({"RN", "AIN", "AINC4"}, [Role])),
    Parsed = Table.AddColumn(NursingRoles, "Definition", each
        let
            Start = Number.From([StartDay]), End = Number.From([EndDay]),
            Hours = Number.From([DurationOfShifts]), StartHours = Number.From([StartTime]),
            EndClock = Time.From([EndTimeCheck]),
            ClockEnd = (Time.Hour(EndClock) * 3600 + Time.Minute(EndClock) * 60 + Time.Second(EndClock)) / 86400,
            ClockHours = (if End <= Start then End + 1 - Start else End - Start) * 24,
            Valid = List.NonNullCount({Start, End, Hours, StartHours, EndClock}) = 5
                and Start >= 0 and Start < 1 and End >= 0 and End < 1
                and Hours > 0 and Hours < 24
                and Number.Abs(StartHours - Start * 24) < 0.00000001
                and Number.Abs(ClockEnd - End) < 0.00000001
                and Number.Abs(ClockHours - Hours) < 0.00000001
                and List.Contains({"AM", "PM", "NIGHT"}, [ShiftPeriod])
        in if Valid then [StartOffset = #duration(0, 0, 0, Number.Round(Start * 86400, 6)),
            EndOffset = #duration(0, 0, 0, Number.Round((if End <= Start then End + 1 else End) * 86400, 6))]
            else error "Invalid or inconsistent role-specific ShiftPeriod definition."),
    Expanded = Table.ExpandRecordColumn(Parsed, "Definition", {"StartOffset", "EndOffset"}),
    Unique = if Table.RowCount(Expanded) = Table.RowCount(Table.Distinct(Expanded, {"Role", "ShiftPeriod"})) then Expanded
        else error "Duplicate role/shift definitions in ShiftPeriod.",
    // Force endpoint validation before a missing calendar match could conceal it.
    Validated = if List.AllTrue(List.Transform(Table.ToRecords(Unique), each [EndOffset] > [StartOffset])) then Unique
        else error "Invalid shift endpoints."
in Validated;

// Query: AvailabilityShiftWindows
// Purpose: Build one dated role/shift window for every Settings permutation, including the final night.
shared AvailabilityShiftWindows = let
    Data = #"IMPORT Availability Settings"{[Item = "PermutationDimensions", Kind = "Table"]}[Data],
    Selected = Table.SelectColumns(Data, {"Date", "Shifts", "RolesList", "Period"}),
    Renamed = Table.RenameColumns(Selected, {{"Shifts", "Shift"}, {"RolesList", "Role"}}),
    NursingRoles = Table.SelectRows(Renamed, each List.Contains({"RN", "AIN", "AINC4"}, [Role])),
    Typed = Table.TransformColumnTypes(NursingRoles, {{"Date", type date}, {"Period", Int64.Type}, {"Role", type text}, {"Shift", type text}}),
    ValidCalendar = if Table.IsEmpty(Typed) then error "Settings roster calendar is empty."
        else if Table.RowCount(Table.SelectRows(Typed, each [Date] = null or [Period] = null or [Shift] = null)) > 0 then
            error "Settings roster calendar contains missing keys."
        else if Table.RowCount(Typed) <> Table.RowCount(Table.Distinct(Typed, {"Date", "Role", "Shift"}))
            or Table.RowCount(Typed) <> Table.RowCount(Table.Distinct(Typed, {"Role", "Period"})) then
            error "Settings roster calendar contains duplicate keys."
        else Typed,
    Joined = Table.NestedJoin(ValidCalendar, {"Role", "Shift"},
        AvailabilityShiftDefinitions, {"Role", "ShiftPeriod"}, "Definition", JoinKind.LeftOuter),
    Complete = if Table.RowCount(Table.SelectRows(Joined, each Table.RowCount([Definition]) <> 1)) = 0 then Joined
        else error "A configured period lacks exactly one role-specific shift definition.",
    Expanded = Table.ExpandTableColumn(Complete, "Definition", {"StartOffset", "EndOffset"}),
    Starts = Table.AddColumn(Expanded, "ShiftStart", each DateTime.From([Date]) + [StartOffset], type datetime),
    Ends = Table.AddColumn(Starts, "ShiftEnd", each DateTime.From([Date]) + [EndOffset], type datetime),
    RosterStart = List.Min(ValidCalendar[Date]),
    Weeks = Table.AddColumn(Ends, "Week", each 1 + Number.IntegerDivide(Duration.Days([Date] - RosterStart), 7), Int64.Type),
    Days = Table.AddColumn(Weeks, "Day", each Date.ToText([Date], "ddd", "en-US"), type text),
    Result = Table.RemoveColumns(Days, {"StartOffset", "EndOffset"})
in Table.Buffer(Result);

// Query: AvailabilityDailyExclusions
// Purpose: Promote a calendar day's distinct excluded hours to a full day when they reach ShiftDuration.
// Inputs: One worker/facility's exclusion intervals and the Settings full-day hours threshold.
// Notes: Overlaps count once. Partial days retain their times; midnight endpoints are exclusive.
shared AvailabilityDailyExclusions = (intervals as list, fullDayHours as number) as list =>
    let
        DailyPieces = List.Combine(List.Transform(intervals, (interval) =>
            List.Generate(
                () => DateTime.From(Date.From(interval[Start])),
                (dayStart) => dayStart < interval[End],
                (dayStart) => dayStart + #duration(1, 0, 0, 0),
                (dayStart) => [DayStart = dayStart,
                    Start = List.Max({interval[Start], dayStart}),
                    End = List.Min({interval[End], dayStart + #duration(1, 0, 0, 0)})]))),
        DailyTable = Table.FromRecords(DailyPieces, type table [DayStart = datetime, Start = datetime, End = datetime]),
        GroupedDays = Table.Group(DailyTable, {"DayStart"}, {{"Exclusions", (rows) =>
            let
                DistinctIntervals = List.Buffer(AvailabilityUnion(Table.ToRecords(Table.SelectColumns(rows, {"Start", "End"})))),
                Hours = List.Sum(List.Transform(DistinctIntervals, each Duration.TotalHours([End] - [Start]))),
                DayStart = rows[DayStart]{0},
                DayEnd = DayStart + #duration(1, 0, 0, 0)
       I cover on three minutes cold to the list, add it to the list. Yeah, d'art for low one so it's hard to tell, but it's used by date is today's sweet it's nice. Are you using an ITX forty seventy right now and wondering should I upgrade to fifty eighty? Let's break it down fast and real first yes technically this isn't a fair fight forty seventy is meteor fifty eighty is high end but people still ask because the upgrade H is real so here's the deal the fifty eighty bring almost touchless of Juda course sixteen gigabytes of next generation GDD R seven VRAM and a way faster rate tracing NDLSS port. It is built for four K gaming heavy AI and video editing compared to your forty seventy, you are getting around forty to fifty percent more FPS at four K, and of course forgetting that kind of performance, it also eats more power and cost. So should you upgrade? Well, I would say if you are gaming on two K or do not use AI tool, stick with forty seven B But if you want to max out four K better renders and future proofing for next three or four fifty eighty is a solid step up so let's remove you would like to step up or would like to save do let us know and comment down below. Do subscribe to Nikitech for Most Such tech insights and I will catch you in the next one are you using an RTX forty seventy right now and wondering should I upgrade to fifty eighty? Let's break it down fast and real first yes technically this isn't a fair fight seventy is meteor fifty eighty is high end but people still ask because the upgrade h is real so here's the deal the fifty eighty bring almost double the Judah course sixteen gigabytes of next generation GDD R seven VRM and a way faster rate tracing NDLSS four it is built for four K gaming heavy AI and video editing compared to your forty seventy you are getting around forty to fifty percent more you know this guy swims up out of nowhere guides me up like an angel probably saved a lot of lives did you know with rough he knew what he was talking about especially down there in the ocean so what our does at least pray like you except your fate I won't thank you for new so good ministers cause I was gonna do in stress in the microwave and balling up grass that we go in another twenty minutes to dinner is already a go microwave did us ten years the podcasters appears to accidentally make the case for socialism just imagine how much less crime there would be if no one was ever starving, if no one was ever no one was ever desperate, if no one was ever worried about whether the next meal was going to come from or whether or not they're going to have a roof over their head, but if like there's no more poverty, what if the concept of poverty is like completely automatically and there is no poverty in this country that already done that already done that rich people that are in charge of industry and corporations and in government to agree to meet the world better forever telling all the wind is socialism curious right now Donald Trump has a pended decades of Republican principal letters discovered taken private prices of social price controls leading a former Ala advisor to recently say this may refollowing if you put price controls on products like pharmaceuticals if you bailed out farms, if you took private stakes in public companies, if we said to energy companies you're making too much money, if you propose credit card caps you would all go aplect ads would give your viewers that's what President Trump's doing you could argue he is the most powerful socialist in the country today's not law that is totally accurate I'm shocked that you guys didn't play imagine in the background
        Result = AvailabilityUnion(List.Combine(GroupedDays[Exclusions]))
    in List.Buffer(Result);

// Query: WorkerAvailabilityRules
// Purpose: Determine eligible worker mode over the whole BD extraction and union roster-overlapping intervals.
// Notes: Any AVAIL restricts the worker to those windows; false intervals always take precedence.
shared WorkerAvailabilityRules = let
    WorkerKeys = Table.Distinct(Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"})),
    ValidRecords = Table.SelectRows(AvailabilityRecords, each [Issue] = null),
    Registered = Table.NestedJoin(ValidRecords, {"Payroll Code", "Facility-Abbrev"},
        WorkerKeys, {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.Inner),
    RequiredRecords = Table.RemoveColumns(Registered, {"Worker"}),
    HorizonStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    FullDayHours = #"EXTRACT EffectiveShiftHrs",
    // Evaluate complete boundary dates before clipping to shift hours, so morning leave is counted.
    CalendarStart = DateTime.From(Date.From(HorizonStart)),
    CalendarEnd = if Time.From(HorizonEnd) = #time(0, 0, 0) then HorizonEnd
        else DateTime.From(Date.From(HorizonEnd)) + #duration(1, 0, 0, 0),
    RosterIntervals = (records as table, available as logical) as list =>
        let
            Overlapping = Table.SelectRows(records, each [Available] = available
                and [Start] < CalendarEnd and [End] > CalendarStart),
            CalendarIntervals = List.Transform(Table.ToRecords(Table.SelectColumns(Overlapping, {"Start", "End"})),
                each [Start = List.Max({[Start], CalendarStart}), End = List.Min({[End], CalendarEnd})]),
            // Full-day expansion applies only to exclusions; explicit AVAIL windows retain their times.
            Adjusted = if available then CalendarIntervals else AvailabilityDailyExclusions(CalendarIntervals, FullDayHours),
            WithinRoster = List.Select(Adjusted, each [Start] < HorizonEnd and [End] > HorizonStart),
            Clipped = List.Transform(WithinRoster,
                each [Start = List.Max({[Start], HorizonStart}), End = List.Min({[End], HorizonEnd})])
        in List.Buffer(AvailabilityUnion(Clipped)),
    // Determine mode from all valid history, then retain only intervals intersecting the roster.
    WorkerIntervals = Table.Group(RequiredRecords, {"Payroll Code", "Facility-Abbrev"},
        {{"HasAvailability", each List.Contains([Available], true), type logical},
         {"AvailableIntervals", each RosterIntervals(_, true), type list},
         {"ExcludedIntervals", each RosterIntervals(_, false), type list}}),
    JoinedWorkers = Table.NestedJoin(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"},
        WorkerIntervals, {"Payroll Code", "Facility-Abbrev"}, "Intervals", JoinKind.LeftOuter),
    WithRules = Table.AddColumn(JoinedWorkers, "Rule", each
        if Table.IsEmpty([Intervals]) then [HasAvailability = false, AvailableIntervals = {}, ExcludedIntervals = {}]
        else Record.SelectFields([Intervals]{0}, {"HasAvailability", "AvailableIntervals", "ExcludedIntervals"})),
    ExpandedRules = Table.ExpandRecordColumn(Table.RemoveColumns(WithRules, {"Intervals"}), "Rule",
        {"HasAvailability", "AvailableIntervals", "ExcludedIntervals"})
// Grouped interval lists are explicitly buffered before expansion to dated shifts.
in Table.Buffer(ExpandedRules);

// Query: WorkerShiftSegments
// Purpose: Intersect available windows with each configured shift and subtract exclusions.
// Output: One worker/shift row with scalar hours and validation flags; no minute-level expansion.
shared WorkerShiftSegments = let
    Windows = Table.Group(AvailabilityShiftWindows, {"Role"}, {{"Shifts", each _, type table}}),
    Joined = Table.NestedJoin(WorkerAvailabilityRules, {"Role"}, Windows, {"Role"}, "Windows", JoinKind.LeftOuter),
    ValidRoles = if Table.RowCount(Table.SelectRows(Joined, each Table.IsEmpty([Windows]))) = 0 then Joined
        else error "An eligible worker role has no configured roster shifts.",
    ShiftLists = Table.AddColumn(ValidRoles, "Shifts", each [Windows]{0}[Shifts], type table),
    Expanded = Table.ExpandTableColumn(Table.RemoveColumns(ShiftLists, {"Windows"}), "Shifts",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}),
    Measured = Table.AddColumn(Expanded, "Measurement", each
        let
            WindowStart = [ShiftStart], WindowEnd = [ShiftEnd],
            Base = if [HasAvailability] then [AvailableIntervals] else {[Start = WindowStart, End = WindowEnd]},
            Overlapping = List.Select(Base, each [Start] < WindowEnd and [End] > WindowStart),
            Clipped = List.Buffer(List.Transform(Overlapping, each
                [Start = List.Max({[Start], WindowStart}), End = List.Min({[End], WindowEnd})])),
            Cuts = List.Buffer(List.Select([ExcludedIntervals], each [Start] < WindowEnd and [End] > WindowStart)),
            Segments = List.Buffer(AvailabilitySubtract(Clipped, Cuts)),
            Hours = List.Accumulate(Segments, 0, (total, segment) => total + Duration.TotalHours(segment[End] - segment[Start])),
            SegmentCount = List.Count(Segments),
            FullShift = SegmentCount = 1 and Segments{0}[Start] = WindowStart and Segments{0}[End] = WindowEnd,
            
            // Validate while shift-local lists are in scope; retain only scalar check results.
            SegmentsValid = List.AllTrue(List.Transform(Segments, (segment) =>
                segment[Start] < segment[End] and segment[Start] >= WindowStart and segment[End] <= WindowEnd
                and not List.AnyTrue(List.Transform(Cuts, (cut) => segment[Start] < cut[End] and segment[End] > cut[Start]))
                and List.AnyTrue(List.Transform(Clipped, (window) => segment[Start] >= window[Start] and segment[End] <= window[End])))),
            
            // Subtraction preserves order, so adjacent endpoint checks replace a second interval union.
            Disjoint = if SegmentCount < 2 then true else List.AllTrue(List.Transform({1..(SegmentCount - 1)},
                (i) => Segments{i - 1}[End] <= Segments{i}[Start]))
        in [AvailableHours = Hours, FullShift = FullShift, SegmentsValid = SegmentsValid, Disjoint = Disjoint]),
    ScalarRows = Table.ExpandRecordColumn(Table.SelectColumns(Measured,
        {"EmployeeID", "Facility-Abbrev", "Role", "Name", "Date", "Shift", "Week", "Day", "ShiftStart", "ShiftEnd", "Measurement"}),
        "Measurement", {"AvailableHours", "FullShift", "SegmentsValid", "Disjoint"})
// Retain zero-hour rows for checks; historical lists do not escape this stage.
in Table.Buffer(ScalarRows);

// Query: ResDayShift_Calculated
// Purpose: Apply the Settings allowance to scalar worker/shift measurements.
// Notes: A fully available shift receives the standard allowance, not its clock duration.
shared ResDayShift_Calculated = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    EffectiveHours = Table.AddColumn(WorkerShiftSegments, "EffectiveShiftHrs", each
        if [FullShift] then Allowance else List.Min({[AvailableHours], Allowance}), type number)
in EffectiveHours;

// Query: AvailabilityCheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared AvailabilityCheckResult = (name as text, evaluate as function) as record =>
    let Result = try evaluate()
    in [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
        Failures = if Result[HasError] then null else Result[Value],
        Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

// Query: ReconciledWorkers_CHECK
// Purpose: Gate worker outputs using identity checks without evaluating shift intervals.
shared ReconciledWorkers_CHECK = let
    Workers = ReconciledWorkers_Eligible,
    Checks = {
        AvailabilityCheckResult("Unique worker keys", () => Table.RowCount(Workers) -
            Table.RowCount(Table.Distinct(Workers, {"EmployeeID", "Facility-Abbrev", "Role"}))),
        // Names alone must remain unique because the existing master list uses that key.
        AvailabilityCheckResult("Unambiguous published names", () => Table.RowCount(Workers) -
            Table.RowCount(Table.Distinct(Workers, {"Name"})))
    }
// Small scalar check rows are reused by staff and shift publication within an evaluation.
in Table.Buffer(Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]));

// Query: ResDayShift_CHECK
// Purpose: Gate shift publication using input checks and materialized scalar measurement checks.
shared ResDayShift_CHECK = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    // Scalar columns only; reuse these rows for every shift check in this evaluation.
    Calculated = Table.Buffer(ResDayShift_Calculated),
    Checks = {
        AvailabilityCheckResult("Settings and role coverage", () =>
            if Allowance > 0 and Table.RowCount(AvailabilityShiftWindows) > 0
                then Table.RowCount(Table.SelectRows(Calculated, each [ShiftEnd] <= [ShiftStart])) else 1),
        AvailabilityCheckResult("Invalid source records", () => Table.RowCount(AvailabilityRecords_Invalid)),
        AvailabilityCheckResult("Unique worker shift rows", () => Table.RowCount(Calculated) -
            Table.RowCount(Table.Distinct(Calculated, {"EmployeeID", "Facility-Abbrev", "Role", "Date", "Shift"}))),
        AvailabilityCheckResult("Segments obey availability and exclusions", () =>
            Table.RowCount(Table.SelectRows(Calculated, each not [SegmentsValid]))),
        AvailabilityCheckResult("Disjoint segments and reconciled hours", () =>
            Table.RowCount(Table.SelectRows(Calculated, each
                not [Disjoint] or [AvailableHours] < 0
                or [AvailableHours] > Duration.TotalHours([ShiftEnd] - [ShiftStart]) + 0.00000001
                or Number.Abs([EffectiveShiftHrs] - (if [FullShift] then Allowance else List.Min({[AvailableHours], Allowance}))) > 0.00000001
                or [EffectiveShiftHrs] < 0 or [EffectiveShiftHrs] > Allowance))),
        AvailabilityCheckResult("Published facility and roles", () => Table.RowCount(Table.SelectRows(Calculated,
            each [#"Facility-Abbrev"] <> "BD" or not List.Contains({"RN", "AIN", "AINC4"}, [Role]))))
    },
    Result = Table.Combine({ReconciledWorkers_CHECK,
        Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])})
// Buffer the small check output; buffers are scoped to an evaluation, not shared across refreshes.
in Table.Buffer(Result);

// Query: ResDayShift
// Purpose: Publish validated worker/date/shift availability with reconciled worker and facility identifiers.
// Output: The original seven columns followed by ID (reconciled EmployeeID) and Facility-Abbrev, both text.
// Notes: CapacityDistrib still assigns one availability unit per row; partial-hour weighting is deferred.
shared ResDayShift = let
    Failures = Table.SelectRows(ResDayShift_CHECK, each [Status] <> "Pass"),
    Checked = if Table.IsEmpty(Failures) then ResDayShift_Calculated
        else error Error.Record("ResDayShift validation", "Required availability checks failed.", Failures),
    // Preserve the original seven columns in order and append identifiers for future downstream matching.
    Positive = Table.SelectRows(Checked, each [EffectiveShiftHrs] > 0),
    Identified = Table.RenameColumns(Positive, {{"EmployeeID", "ID"}}),
    Output = Table.SelectColumns(Identified, {"Role", "Week", "Day", "Shift", "EffectiveShiftHrs", "Date", "Name", "ID", "Facility-Abbrev"}),
    Typed = Table.TransformColumnTypes(Output, {{"Role", type text}, {"Week", Int64.Type}, {"Day", type text},
        {"Shift", type text}, {"EffectiveShiftHrs", type number}, {"Date", type date}, {"Name", type text},
        {"ID", type text}, {"Facility-Abbrev", type text}})
in Table.Sort(Typed, {{"Week", Order.Ascending}, {"Date", Order.Ascending}, {"Role", Order.Ascending}, {"Name", Order.Ascending}, {"Shift", Order.Ascending}});

shared MultiRoles = let
    Source = #"IMPORT MultiRolesResRolesProportion",
    #"Removed Columns1" = Table.RemoveColumns(Source,{"Total"}),
    #"Unpivoted Other Columns1" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Name", "AINC4HrsAvilPref"}, "MultiRole", "Value"),
    #"Removed Columns" = Table.RemoveColumns(#"Unpivoted Other Columns1",{"Value"})
in
    #"Removed Columns";

shared RoleResDayAvailabilityCapped = let
    Source = ResDayShift,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Added STDDAYSAVAIL" = Table.AddColumn(#"Grouped Rows", "StdRosterDays", each StdRosterDays),
    #"Merged Queries" = Table.NestedJoin(#"Added STDDAYSAVAIL", {"Name"}, MultiRoles, {"Name"}, "IMPORT ResRolesProportion2", JoinKind.LeftOuter),
    #"Expanded IMPORT ResRolesProportion2" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ResRolesProportion2", {"AINC4HrsAvilPref", "MultiRole"}, {"AINC4HrsAvilPref", "MultiRole"}),
    #"Added ROLESPLIT" = Table.AddColumn(#"Expanded IMPORT ResRolesProportion2", "RoleSplit", each if [MultiRole] = null then null  
else if [MultiRole] = "AIN" and [MultiRole] = [Role] then 1 - [AINC4HrsAvilPref]
else if [MultiRole] = [Role] then [AINC4HrsAvilPref]
else 99),
    #"Filtered ROLESPLIT" = Table.SelectRows(#"Added ROLESPLIT", each ([RoleSplit] <> 99)),
    #"Added SPLITAVAIL" = Table.AddColumn(#"Filtered ROLESPLIT", "AvailabilityCapped", each if [Count] <= [StdRosterDays] then [Count] else [StdRosterDays] * (if [MultiRole] <> null then [RoleSplit] else 1)),
    #"Transformed Columns" = Table.TransformColumns(
        #"Added SPLITAVAIL",
        {"Name", each if try Record.Field(_, "MultiRole") <> null otherwise false then _ & "(" & Record.Field(_, "MultiRole") & ")" else _}
    ),
    #"Renamed Columns" = Table.RenameColumns(#"Transformed Columns",{{"Name", "NameX"}}),
    #"NEW NAME" = Table.AddColumn(#"Renamed Columns", "Name", each if [MultiRole] <> null then [NameX] & " (" & [MultiRole] & ")" else [NameX]),
    #"Renamed Columns1" = Table.RenameColumns(#"NEW NAME",{{"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns1", "Role", each if [MultiRole] <> null then [MultiRole] else [RoleX]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Conditional Column",{"AINC4HrsAvilPref", "MultiRole", "RoleSplit", "RoleX", "NameX"})
in
    #"Removed Columns";

shared RoleAvailabilityCapped = let
    Source = RoleResDayAvailabilityCapped,
    #"Grouped Rows" = Table.Group(Source, {"Role"}, {{"AvailabilityCapped", each List.Sum([AvailabilityCapped]), type number}})
in
    #"Grouped Rows";

// Query: Availability-StaffList
// Purpose: Preserve the staff identity interface, including eligible workers with no available shifts.
shared #"Availability-StaffList" = let
    Failures = Table.SelectRows(ReconciledWorkers_CHECK, each [Status] <> "Pass"),
    Workers = if Table.IsEmpty(Failures) then ReconciledWorkers_Eligible
        else error Error.Record("Availability staff validation", "Required worker identity checks failed.", Failures),
    Output = Table.Distinct(Table.SelectColumns(Workers, {"Name", "Role"}))
in Table.Sort(Output, {{"Name", Order.Ascending}, {"Role", Order.Ascending}});

shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    FilePathUrl =
    let
        Source = try Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content] otherwise Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        FirstColumn = Table.ColumnNames(Source){0},
        RenamedColumns = if FirstColumn = "FilePath" then Source else Table.RenameColumns(Source, {{FirstColumn, "FilePath"}}, MissingField.Ignore),
        ReplacedValue = Table.TransformColumns(RenamedColumns, {{"FilePath", each Text.Replace(Text.From(_), "/", "\"), type text}}),
        BufferedTable = Table.Buffer(ReplacedValue)
    in
        BufferedTable,

    RawFilePath = FilePathUrl{0}[FilePath],
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(RawFilePath),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    SharePointCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SharepointRootUrl], type text),
    SharePointDocumentsCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type text),
    LocalCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SyncedFolderRootPath], type text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(
        MatchCandidates_WithLength,
        each [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    BestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(LocalFullPath, "[", "]") otherwise Text.AfterDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

shared #"FilePath - 2Calculations" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"FilePath - 1Input" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;
