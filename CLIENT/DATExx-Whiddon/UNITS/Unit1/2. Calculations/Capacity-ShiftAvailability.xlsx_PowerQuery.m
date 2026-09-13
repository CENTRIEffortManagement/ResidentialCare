// Power Query from: Capacity-ShiftAvailability.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Capacity-ShiftAvailability.xlsx
// Extracted: 2026-09-12T21:57:06.510Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve the configured workbook location through the existing CentriSyncPaths mapping.
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

// Query: FilePath - 2Calculations
// Purpose: Return the calculation folder from the configured workbook location.
shared #"FilePath - 2Calculations" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: FilePath - 1Input
// Purpose: Return the input folder using the established Unit1 path convention.
shared #"FilePath - 1Input" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;

// Query: IMPORT MultiRolesResRolesProportion
// Purpose: Read the existing role-suffix source and the externally consumed MultiRoles output.
shared #"IMPORT MultiRolesResRolesProportion" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    ResRolesProportion_Table = Source{[Item="ResRolesProportion",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResRolesProportion_Table,{{"Name", type text}, {"AINC4", type number}, {"AIN", type number}, {"Total", type number}, {"AINC4HrsAvilPref", type number}})
in
    #"Changed Type";

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

// Query: AvailabilityLeaveReasons
// Purpose: List exact recognised leave descriptions; new descriptions require an explicit mapping.
// Notes: Includes source descriptions confirmed in the 197-record diagnosis. Matching ignores surrounding spaces and letter case.
shared AvailabilityLeaveReasons = #table(type table [Reason = text], {
    {"Annual Rec Lve"},
    {"Maternity Lve Unpaid"},
    {"Unpaid Leave"},
    {"Study Leave Unpaid"},
    {"Workers Comp Not Worked Sec 40"},
    {"Personal Sick Leave Unpaid"},
    {"Personal Sick Leave Paid"},
    {"WC Sec 37 13-130wks unfit"},
    {"WC Sec 36 3-13 wks unfit"},
    {"Personal Carer Leave Paid"},
    {"Personal Carer Leave Unpaid"},
    {"Long Service Lve"},
    {"Compassionate Lve Paid"},
    {"Personal Emergency Leave (PEL)"}
});

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

// Query: AvailabilityIntervalHours
// Purpose: Sum clock hours in an already disjoint interval list.
shared AvailabilityIntervalHours = (intervals as list) as number =>
    List.Accumulate(intervals, 0, (total, interval) => total + Duration.TotalHours(interval[End] - interval[Start]));

// Query: AvailabilityClipIntervals
// Purpose: Retain half-open intersections with one shift or calendar range.
shared AvailabilityClipIntervals = (intervals as list, start as datetime, end as datetime) as list =>
    List.Buffer(List.Transform(List.Select(intervals, each [Start] < end and [End] > start),
        each [Start = List.Max({[Start], start}), End = List.Min({[End], end})]));

// Query: AvailabilityShiftMeasurement
// Purpose: Measure a single shift and check interval invariants while its small lists are in scope.
shared AvailabilityShiftMeasurement = (baseline as list, unavail as list, leave as list,
    shiftStart as datetime, shiftEnd as datetime, fullDayLeave as logical) as record =>
    let
        Base = AvailabilityClipIntervals(baseline, shiftStart, shiftEnd),
        Unavail = AvailabilityClipIntervals(unavail, shiftStart, shiftEnd),
        Leave = AvailabilityClipIntervals(leave, shiftStart, shiftEnd),
        // UNION across both categories prevents overlapping leave and UNAVAIL being deducted twice.
        Cuts = List.Buffer(AvailabilityUnion(Unavail & Leave)),
        Segments = if fullDayLeave then {} else List.Buffer(AvailabilitySubtract(Base, Cuts)),
        Count = List.Count(Segments),
        FullShift = Count = 1 and Segments{0}[Start] = shiftStart and Segments{0}[End] = shiftEnd,
        SegmentsValid = List.AllTrue(List.Transform(Segments, (segment) =>
            segment[Start] < segment[End] and segment[Start] >= shiftStart and segment[End] <= shiftEnd
            and not List.AnyTrue(List.Transform(Cuts, (cut) => segment[Start] < cut[End] and segment[End] > cut[Start]))
            and List.AnyTrue(List.Transform(Base, (window) => segment[Start] >= window[Start] and segment[End] <= window[End])))),
        Disjoint = if Count < 2 then true else List.AllTrue(List.Transform({1..(Count - 1)},
            (i) => Segments{i - 1}[End] <= Segments{i}[Start]))
    in [BaselineHours = AvailabilityIntervalHours(Base), LeaveHoursInShift = AvailabilityIntervalHours(Leave),
        UnavailHoursInShift = AvailabilityIntervalHours(Unavail), CombinedAbsenceHoursInShift = AvailabilityIntervalHours(Cuts),
        AvailableHours = AvailabilityIntervalHours(Segments), FullShift = FullShift, SegmentsValid = SegmentsValid, Disjoint = Disjoint];

// Query: AvailabilityResidualHours
// Purpose: Apply the approved allowance-minus-absence rule, capped by residual clock availability.
// Notes: A complete shift receives the allowance even when its clock length differs; no break/minimum rule.
shared AvailabilityResidualHours = (remainingHours as number, absenceHours as number,
    fullShift as logical, fullDayLeave as logical, allowance as number) as number =>
    if fullDayLeave or remainingHours <= 0 then 0
    else if fullShift then allowance
    else List.Max({0, List.Min({remainingHours, allowance - absenceHours})});

// Query: AvailabilityRecordKind
// Purpose: Separate AVAIL, UNAVAIL and recognised leave without interpreting unknown reasons as leave.
shared AvailabilityRecordKind = let
    LeaveNames = List.Buffer(List.Transform(AvailabilityLeaveReasons[Reason], each Text.Upper(Text.Trim(_))))
in (reason as any) as text =>
    let Clean = AvailabilityText(reason)
    // Preserve case-sensitive source AVAIL/UNAVAIL rules, with UNAVAIL taking precedence.
    in if Clean = null then "Unknown"
        else if Text.Contains(Clean, "UNAVAIL") then "UNAVAIL"
        else if Text.Contains(Clean, "AVAIL") then "AVAIL"
        else if List.Contains(LeaveNames, Text.Upper(Clean)) then "Leave"
        else "Unknown";

// Query: AvailabilityDailyWindows
// Purpose: Reset the availability baseline on each calendar day before applying exclusions.
// Inputs: Valid AVAIL intervals and complete calendar boundaries, with an exclusive midnight end.
// Output: Unioned baseline intervals; a day without AVAIL is fully available.
// Notes: An interval spanning midnight affects both dates only for its recorded duration.
shared AvailabilityDailyWindows = (intervals as list, calendarStart as datetime, calendarEnd as datetime) as list =>
    let
        // Reuse the worker's AVAIL list while checking each date in the roster horizon.
        Available = List.Buffer(intervals),
        Days = List.Generate(() => calendarStart, each _ < calendarEnd,
            each _ + #duration(1, 0, 0, 0)),
        DailyWindows = List.Transform(Days, (dayStart) =>
            let
                DayEnd = dayStart + #duration(1, 0, 0, 0),
                Overlapping = List.Select(Available,
                    each [Start] < DayEnd and [End] > dayStart),
                // Only AVAIL on this date restricts this date. Tomorrow is evaluated afresh.
                Baseline = if List.IsEmpty(Overlapping) then {[Start = dayStart, End = DayEnd]}
                    else List.Transform(Overlapping, each [
                        Start = List.Max({[Start], dayStart}), End = List.Min({[End], DayEnd})])
            in Baseline),
        // Merge touching windows across midnight so a fully available night remains one segment.
        Result = AvailabilityUnion(List.Combine(DailyWindows))
    in List.Buffer(Result);

// Query: AvailabilityLeaveDayTotals
// Purpose: Split and union recognised leave by calendar date, never including UNAVAIL.
// Inputs: One worker/facility's recognised leave intervals only.
// Output: Date, raw record-hour sum, distinct leave hours and source-piece count.
shared AvailabilityLeaveDayTotals = (intervals as list) as table =>
    let
        Pieces = List.Combine(List.Transform(intervals, (interval) =>
            List.Generate(() => DateTime.From(Date.From(interval[Start])), (d) => d < interval[End],
                (d) => d + #duration(1, 0, 0, 0),
                (d) => [Date = Date.From(d), Start = List.Max({interval[Start], d}),
                    End = List.Min({interval[End], d + #duration(1, 0, 0, 0)})]))),
        Rows = Table.FromRecords(Pieces, type table [Date = date, Start = datetime, End = datetime]),
        Days = Table.Group(Rows, {"Date"}, {
            {"SumOfLeaveRecordHours", each List.Sum(List.Transform(Table.ToRecords(_), (r) => Duration.TotalHours(r[End] - r[Start]))), type number},
            {"DistinctLeaveHours", each AvailabilityIntervalHours(AvailabilityUnion(Table.ToRecords(Table.SelectColumns(_, {"Start", "End"})))), type number},
            {"LeaveRecordCount", each Table.RowCount(_), Int64.Type}})
    in Days;

// Query: AvailabilityCheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared AvailabilityCheckResult = (name as text, evaluate as function) as record =>
    let Result = try evaluate()
    in [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
        Failures = if Result[HasError] then null else Result[Value],
        Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

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
// Purpose: Parse source intervals and retain separate record kinds and traceable validation issues.
// Notes: Available remains a compatibility flag; only RecordKind = Leave contributes to daily leave hours.
shared AvailabilityRecords = let
    Parsed = Table.AddColumn(#"IMPORT Combined Availabilities", "Parsed", each
        let
            Start = try DateTime.From([Date From]) otherwise null,
            End = try DateTime.From([Date To]) otherwise null,
            Reason = AvailabilityText([Availability or Leave Reason]),
            Kind = AvailabilityRecordKind(Reason),
            Issue = if [Payroll Code] = null then "Missing or invalid payroll code"
                else if Kind = "Unknown" then "Unrecognised or missing reason: " & (if Reason = null then "(blank)" else Reason)
                else if Start = null or End = null then "Missing or invalid datetime"
                else if End <= Start then "End must be later than start"
                else null
        in [Start = Start, End = End, RecordKind = Kind, OriginalReason = Reason,
            Available = if Kind = "Unknown" then null else Kind = "AVAIL", Issue = Issue]),
    Expanded = Table.ExpandRecordColumn(Parsed, "Parsed", {"Start", "End", "RecordKind", "OriginalReason", "Available", "Issue"})
// Buffer only scalar fields reused by identity-independent diagnostics and eligible interval staging.
in Table.Buffer(Table.SelectColumns(Expanded,
    {"Payroll Code", "Facility-Abbrev", "SourceRecord", "OriginalReason", "Start", "End", "RecordKind", "Available", "Issue"}));

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

// Query: AvailabilityEligibleRecords
// Purpose: Keep valid eligible-worker records on complete calendar dates covered by the roster.
// Notes: Unique worker/facility keys prevent multiple roles multiplying the source records.
shared AvailabilityEligibleRecords = let
    WorkerKeys = Table.Distinct(Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"})),
    CalendarStart = DateTime.From(Date.From(List.Min(AvailabilityShiftWindows[ShiftStart]))),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    CalendarEnd = if Time.From(HorizonEnd) = #time(0, 0, 0) then HorizonEnd
        else DateTime.From(Date.From(HorizonEnd)) + #duration(1, 0, 0, 0),
    Valid = Table.SelectRows(AvailabilityRecords, each [Issue] = null and [Start] < CalendarEnd and [End] > CalendarStart),
    Joined = Table.NestedJoin(Valid, {"Payroll Code", "Facility-Abbrev"}, WorkerKeys,
        {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.Inner),
    Narrow = Table.RemoveColumns(Joined, {"Worker", "Available"}),
    Clipped = Table.TransformColumns(Narrow,
        {{"Start", each List.Max({_, CalendarStart}), type datetime}, {"End", each List.Min({_, CalendarEnd}), type datetime}})
in Table.Buffer(Clipped);

// Query: WorkerLeaveDays
// Purpose: Decide full-day leave once per worker/facility/date, independently of role and UNAVAIL.
// Notes: A full-day decision blocks every configured shift STARTING on that date, including NIGHT.
shared WorkerLeaveDays = let
    Leave = Table.SelectRows(AvailabilityEligibleRecords, each [RecordKind] = "Leave"),
    Workers = Table.Group(Leave, {"Payroll Code", "Facility-Abbrev"},
        {{"Days", each AvailabilityLeaveDayTotals(Table.ToRecords(Table.SelectColumns(_, {"Start", "End"}))), type table}}),
    Expanded = Table.ExpandTableColumn(Workers, "Days", {"Date", "SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount"}),
    Identified = Table.RenameColumns(Expanded, {{"Payroll Code", "EmployeeID"}}),
    Threshold = #"EXTRACT EffectiveShiftHrs",
    Blocked = Table.AddColumn(Identified, "FullDayLeaveBlocked", each [DistinctLeaveHours] >= Threshold, type logical)
// Scalar daily results are reused by each role/shift and its diagnostic rows.
in Table.Buffer(Blocked);

// Query: WorkerAvailabilityRules
// Purpose: Build date-specific availability windows and separate raw leave and UNAVAIL intervals.
// Notes: UNAVAIL is never promoted to a full-day exclusion; WorkerLeaveDays owns that leave-only decision.
shared WorkerAvailabilityRules = let
    HorizonStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    CalendarStart = DateTime.From(Date.From(HorizonStart)),
    CalendarEnd = if Time.From(HorizonEnd) = #time(0, 0, 0) then HorizonEnd
        else DateTime.From(Date.From(HorizonEnd)) + #duration(1, 0, 0, 0),
    Intervals = (records as table, kind as text) as list => List.Buffer(AvailabilityUnion(
        Table.ToRecords(Table.SelectColumns(Table.SelectRows(records, each [RecordKind] = kind), {"Start", "End"})))),
    Grouped = Table.Group(AvailabilityEligibleRecords, {"Payroll Code", "Facility-Abbrev"}, {
        {"AvailableIntervals", each AvailabilityDailyWindows(Intervals(_, "AVAIL"), CalendarStart, CalendarEnd), type list},
        {"UnavailIntervals", each Intervals(_, "UNAVAIL"), type list},
        {"LeaveIntervals", each Intervals(_, "Leave"), type list}}),
    Joined = Table.NestedJoin(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"},
        Grouped, {"Payroll Code", "Facility-Abbrev"}, "Rules", JoinKind.LeftOuter),
    Defaults = Table.AddColumn(Joined, "Rule", each if Table.IsEmpty([Rules]) then
        [AvailableIntervals = {[Start = HorizonStart, End = HorizonEnd]}, UnavailIntervals = {}, LeaveIntervals = {}]
        else Record.SelectFields([Rules]{0}, {"AvailableIntervals", "UnavailIntervals", "LeaveIntervals"})),
    Expanded = Table.ExpandRecordColumn(Table.RemoveColumns(Defaults, {"Rules"}), "Rule",
        {"AvailableIntervals", "UnavailIntervals", "LeaveIntervals"})
// Nested interval lists are explicitly buffered in their constructors; the outer table is small.
in Table.Buffer(Expanded);

// Query: WorkerShiftSegments
// Purpose: Measure baseline, leave, UNAVAIL and residual clock hours for every eligible worker/shift.
// Output: Scalar measurements and checks, including zero-hour shifts; no historical lists on expanded rows.
shared WorkerShiftSegments = let
    Windows = Table.Group(AvailabilityShiftWindows, {"Role"}, {{"Shifts", each _, type table}}),
    Joined = Table.NestedJoin(WorkerAvailabilityRules, {"Role"}, Windows, {"Role"}, "Windows", JoinKind.LeftOuter),
    ValidRoles = if Table.RowCount(Table.SelectRows(Joined, each Table.IsEmpty([Windows]))) = 0 then Joined
        else error "An eligible worker role has no configured roster shifts.",
    ShiftLists = Table.AddColumn(ValidRoles, "Shifts", each [Windows]{0}[Shifts], type table),
    Expanded = Table.ExpandTableColumn(Table.RemoveColumns(ShiftLists, {"Windows"}), "Shifts",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"}),
    JoinedDays = Table.NestedJoin(Expanded, {"EmployeeID", "Facility-Abbrev", "Date"},
        WorkerLeaveDays, {"EmployeeID", "Facility-Abbrev", "Date"}, "LeaveDay", JoinKind.LeftOuter),
    DayValues = Table.AddColumn(JoinedDays, "DayAssessment", each if Table.IsEmpty([LeaveDay]) then
        [SumOfLeaveRecordHours = 0, DistinctLeaveHours = 0, LeaveRecordCount = 0, FullDayLeaveBlocked = false]
        else Record.SelectFields([LeaveDay]{0}, {"SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked"})),
    DayColumns = Table.ExpandRecordColumn(Table.RemoveColumns(DayValues, {"LeaveDay"}), "DayAssessment",
        {"SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked"}),
    Measured = Table.AddColumn(DayColumns, "Measurement", each AvailabilityShiftMeasurement(
        [AvailableIntervals], [UnavailIntervals], [LeaveIntervals], [ShiftStart], [ShiftEnd], [FullDayLeaveBlocked])),
    ScalarRows = Table.ExpandRecordColumn(Table.SelectColumns(Measured,
        {"EmployeeID", "Facility-Abbrev", "Role", "Name", "Date", "Shift", "Week", "Day", "ShiftStart", "ShiftEnd",
         "SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked", "Measurement"}),
        "Measurement", {"BaselineHours", "LeaveHoursInShift", "UnavailHoursInShift", "CombinedAbsenceHoursInShift",
            "AvailableHours", "FullShift", "SegmentsValid", "Disjoint"})
in Table.Buffer(ScalarRows);

// Query: ResDayShift_Calculated
// Purpose: Deduct distinct shift absences from ShiftDuration and respect the remaining AVAIL windows.
// Notes: Full-day leave blocks all shifts on the date; a zero-hour shift alone never blocks other shifts.
shared ResDayShift_Calculated = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    WithAllowance = Table.AddColumn(WorkerShiftSegments, "ShiftDuration", each Allowance, type number),
    Effective = Table.AddColumn(WithAllowance, "EffectiveShiftHrs", each AvailabilityResidualHours(
        [AvailableHours], [CombinedAbsenceHoursInShift], [FullShift], [FullDayLeaveBlocked], Allowance), type number),
    Outcome = Table.AddColumn(Effective, "AvailabilityDecision", each
        if [FullDayLeaveBlocked] then "All shifts blocked: daily recognised leave reaches ShiftDuration"
        else if [BaselineHours] = 0 then "Unavailable: outside this date's AVAIL windows"
        else if [AvailableHours] = 0 then "Unavailable: recorded absences cover available time in this shift"
        else if [CombinedAbsenceHoursInShift] >= Allowance then "Unavailable: shift absences exhaust ShiftDuration"
        else if [FullShift] then "Full shift available"
        else "Residual availability", type text)
in Outcome;

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

// Query: Availability-StaffList
// Purpose: Preserve the staff identity interface, including eligible workers with no available shifts.
shared #"Availability-StaffList" = let
    Failures = Table.SelectRows(ReconciledWorkers_CHECK, each [Status] <> "Pass"),
    Workers = if Table.IsEmpty(Failures) then ReconciledWorkers_Eligible
        else error Error.Record("Availability staff validation", "Required worker identity checks failed.", Failures),
    Output = Table.Distinct(Table.SelectColumns(Workers, {"Name", "Role"}))
in Table.Sort(Output, {{"Name", Order.Ascending}, {"Role", Order.Ascending}});

// Query: MultiRoles
// Purpose: Preserve the table imported by MutliRoleCheck; independent of daily leave decisions.
shared MultiRoles = let
    Source = #"IMPORT MultiRolesResRolesProportion",
    #"Removed Columns1" = Table.RemoveColumns(Source,{"Total"}),
    #"Unpivoted Other Columns1" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Name", "AINC4HrsAvilPref"}, "MultiRole", "Value"),
    #"Removed Columns" = Table.RemoveColumns(#"Unpivoted Other Columns1",{"Value"})
in
    #"Removed Columns";

// Query: AvailabilityReasons_DIAGNOSTICS
// Purpose: Surface every unknown reason for mapping review, including records outside the eligible population.
shared AvailabilityReasons_DIAGNOSTICS = let
    Unknown = Table.SelectRows(AvailabilityRecords, each [RecordKind] = "Unknown")
in Table.SelectColumns(Unknown, {"SourceRecord", "Payroll Code", "Facility-Abbrev", "OriginalReason", "Start", "End", "Issue"});

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

// Query: AvailabilityRules_TEST
// Purpose: Exercise production helpers against fixed expected answers without opening source workbooks.
// Notes: Synthetic datetimes, hours and reason labels only; this is also a required publication check.
shared AvailabilityRules_TEST = let
    Origin = #datetime(2026, 7, 22, 0, 0, 0),
    At = (hour as number) as datetime => Origin + #duration(0, 0, 0, hour * 3600),
    Windows = (pairs as list) as list => AvailabilityUnion(List.Transform(pairs, each [Start = At(_{0}), End = At(_{1})])),
    Cases = #table(type table [Case = text, StartHour = number, EndHour = number, Base = list,
        Unavail = list, Leave = list, Allowance = number, ExpectedDayHours = number, ExpectedEffective = number], {
        {"No records", 6, 14, {{0, 48}}, {}, {}, 7.6, 0, 7.6},
        {"Short UNAVAIL preserves AM residual", 6, 14, {{0, 48}}, {{13, 15}}, {}, 7.6, 0, 6.6},
        {"Same UNAVAIL preserves PM residual", 14, 22, {{0, 48}}, {{13, 15}}, {}, 7.6, 0, 6.6},
        {"UNAVAIL alone never blocks other shifts", 14, 22, {{0, 48}}, {{6, 14}}, {}, 7.6, 0, 7.6},
        {"UNAVAIL exhausts its own shift", 6, 14, {{0, 48}}, {{6, 14}}, {}, 7.6, 0, 0},
        {"UNAVAIL excluded from daily leave tally", 14, 22, {{0, 48}}, {{6, 10}}, {{10, 14}}, 7.6, 4, 7.6},
        {"10-15 leave AM", 6, 14, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 3.6},
        {"10-15 leave PM", 14, 22, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 6.6},
        {"10-15 leave NIGHT", 22, 30, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 7.6},
        {"Full-day leave blocks AM", 6, 14, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Full-day leave blocks PM", 14, 22, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Full-day leave blocks entire dated NIGHT", 22, 30, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Duplicate four-hour leave counts once", 6, 14, {{0, 48}}, {}, {{6, 10}, {6, 10}}, 7.6, 4, 3.6},
        {"Separate leave periods reach threshold", 14, 22, {{0, 48}}, {}, {{6, 10}, {12, 16}}, 7.6, 8, 0},
        {"Cross-category overlap deducted once", 6, 14, {{0, 48}}, {{8, 12}}, {{10, 14}}, 7.6, 4, 1.6},
        {"Exact midnight end does not affect next day", 30, 38, {{0, 48}}, {}, {{0, 24}}, 7.6, 0, 7.6},
        {"AVAIL window caps residual hours", 6, 14, {{9, 13}}, {}, {{10, 11}}, 7.6, 1, 3},
        {"AVAIL does not cover this shift", 14, 22, {{9, 13}}, {}, {}, 7.6, 0, 0},
        {"Full 7.5-clock-hour shift gets allowance", 22.5, 30, {{0, 48}}, {}, {}, 7.6, 0, 7.6},
        {"Changed Settings allowance", 6, 14, {{0, 48}}, {}, {{10, 15}}, 8, 5, 4},
        {"Exactly ShiftDuration blocks the day", 22, 30, {{0, 48}}, {}, {{6, 13.6}}, 7.6, 7.6, 0},
        {"Midnight-to-midnight UNAVAIL does not block NIGHT after midnight", 22, 30, {{0, 48}}, {{0, 24}}, {}, 7.6, 0, 5.6}
    }),
    RuleChecks = List.Transform(Table.ToRecords(Cases), (c) => AvailabilityCheckResult(c[Case], () =>
        let
            // Raw leave rows feed daily totals so the duplicate-handling test is meaningful.
            RawLeave = List.Transform(c[Leave], each [Start = At(_{0}), End = At(_{1})]),
            DayTotals = AvailabilityLeaveDayTotals(RawLeave),
            Day = Table.SelectRows(DayTotals, each [Date] = Date.From(At(c[StartHour]))),
            DayHours = if Table.IsEmpty(Day) then 0 else Day{0}[DistinctLeaveHours],
            Blocked = DayHours >= c[Allowance],
            M = AvailabilityShiftMeasurement(Windows(c[Base]), Windows(c[Unavail]), Windows(c[Leave]),
                At(c[StartHour]), At(c[EndHour]), Blocked),
            Effective = AvailabilityResidualHours(M[AvailableHours], M[CombinedAbsenceHoursInShift], M[FullShift], Blocked, c[Allowance])
        in if Number.Abs(DayHours - c[ExpectedDayHours]) < 0.00000001
            and Number.Abs(Effective - c[ExpectedEffective]) < 0.00000001 and M[SegmentsValid] and M[Disjoint] then 0 else 1)),
    OtherChecks = {
        AvailabilityCheckResult("Classification separates leave and UNAVAIL", () =>
            if AvailabilityRecordKind("UNAVAIL - Unpaid Leave") = "UNAVAIL"
                and AvailabilityRecordKind("AVAIL") = "AVAIL"
                and AvailabilityRecordKind("Maternity Lve Unpaid") = "Leave"
                and AvailabilityRecordKind("Unmapped reason") = "Unknown"
                and AvailabilityRecordKind(null) = "Unknown" then 0 else 1),
        AvailabilityCheckResult("AVAIL resets independently on next date", () =>
            let Baseline = AvailabilityDailyWindows(Windows({{9, 13}}), At(0), At(48))
            in if AvailabilityIntervalHours(AvailabilityClipIntervals(Baseline, At(0), At(24))) = 4
                and AvailabilityIntervalHours(AvailabilityClipIntervals(Baseline, At(24), At(48))) = 24 then 0 else 1)
    }
in Table.Buffer(Table.FromRecords(RuleChecks & OtherChecks,
    type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]));

// Query: ResDayShift_CHECK
// Purpose: Gate shift publication using input checks and materialized scalar measurement checks.
shared ResDayShift_CHECK = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    // Scalar columns only; reuse these rows for every shift check in this evaluation.
    Calculated = Table.Buffer(ResDayShift_Calculated),
    Checks = {
        AvailabilityCheckResult("Known-answer availability rules", () => Table.RowCount(Table.SelectRows(AvailabilityRules_TEST, each [Status] <> "Pass"))),
        AvailabilityCheckResult("Unique worker/date leave totals", () => Table.RowCount(WorkerLeaveDays) -
            Table.RowCount(Table.Distinct(WorkerLeaveDays, {"EmployeeID", "Facility-Abbrev", "Date"}))),
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
                or Number.Abs([EffectiveShiftHrs] - (if [FullDayLeaveBlocked] or [AvailableHours] <= 0 then 0
                    else if [FullShift] then Allowance
                    else List.Max({0, List.Min({[AvailableHours], Allowance - [CombinedAbsenceHoursInShift]})}))) > 0.00000001
                or ([FullDayLeaveBlocked] <> ([DistinctLeaveHours] >= Allowance))
                or ([FullDayLeaveBlocked] and ([EffectiveShiftHrs] <> 0 or [AvailableHours] <> 0))
                or [CombinedAbsenceHoursInShift] > [LeaveHoursInShift] + [UnavailHoursInShift] + 0.00000001
                or [EffectiveShiftHrs] < 0 or [EffectiveShiftHrs] > Allowance))),
        AvailabilityCheckResult("Published facility and roles", () => Table.RowCount(Table.SelectRows(Calculated,
            each [#"Facility-Abbrev"] <> "BD" or not List.Contains({"RN", "AIN", "AINC4"}, [Role]))))
    },
    Result = Table.Combine({ReconciledWorkers_CHECK,
        Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])})
// Buffer the small check output; buffers are scoped to an evaluation, not shared across refreshes.
in Table.Buffer(Result);