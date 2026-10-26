section CSUnitDemandNative;
shared CSUnitDemandNativeTests = () as table => let
Evaluate = (Inputs as record) as record => let
Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    WorkbookFolder = #"Filtered Rows"{0}[Value],
    UnitFolder =
        if Text.EndsWith(WorkbookFolder, "\1. Input", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\1. Input"))
        else if Text.EndsWith(WorkbookFolder, "\2. Calculations", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\2. Calculations"))
        else
            error "Root Path did not end in an expected unit workbook folder: " & WorkbookFolder
in
    UnitFolder,
UnitL1PathTABLE = let
    FilePathUrl =
    let
        Source = Inputs[PathInput],
        // Excel exposes the existing single-cell FilePathUrl name as Column1.
        // Keep the canonical FilePath column when the standard input table is installed.
        PathInput = if Table.HasColumns(Source, {"FilePath"}) then Source
            else if Table.ColumnNames(Source) = {"Column1"} then
                Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must be a one-row FilePath table or a single named cell.",
        SelectedColumns = Table.SelectColumns(PathInput, {"FilePath"}),
        ChangedType = Table.TransformColumnTypes(SelectedColumns, {{"FilePath", type text}}),
        ReplacedValue = Table.TransformColumns(ChangedType, {{"FilePath", each if _ = null then null else Text.Replace(_, "/", "\"), type text}}),
        ValidatedTable = if Table.RowCount(ReplacedValue) = 1 then ReplacedValue else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedTable)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if RawFilePathValue = null or Text.Trim(RawFilePathValue) = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else Text.Trim(RawFilePathValue),
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet.
    // Remove only the workbook brackets and sheet suffix before resolving the folder.
    WorkbookPath = if Text.Contains(RawFilePath, "[") then
        Text.BeforeDelimiter(RawFilePath, "[") & Text.BetweenDelimiters(RawFilePath, "[", "]")
        else RawFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    ValidatedWorkbookPath = if Comparer.OrdinalIgnoreCase(InputFileName, "2-DemandExtract.xlsx") = 0 then
        WorkbookPath
        else error "FilePathUrl identifies another workbook. In the named cell use =CELL(""filename"",A1), then save and recalculate 2-DemandExtract.xlsx.",
    CentriSyncPaths_Source = #"IMPORT CentriSyncPaths",
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(ValidatedWorkbookPath),
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
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")
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
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase),
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
            {"FilePathUrl", FilePath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER,
#"IMPORT CentriSyncPaths" = Inputs[PathNavigation],
#"Demand Facility" = let
    UnitRows = Table.SelectRows(UnitL1PathTABLE, each [Variable Name] = "Unit"),
    UnitName = if Table.RowCount(UnitRows) = 1 then UnitRows{0}[Value] else error "UnitL1PathTABLE must identify one unit folder."
in
    if UnitName = null or Text.Trim(UnitName) = "" then error "The workbook path did not identify a unit folder." else UnitName,
#"Demand Roles" = let
    Enabled = Table.SelectRows(#"Demand Role Analysis Prepare", each [Effort Management Analysis] = true),
    Groups = List.Sort(List.Distinct(Enabled[Roles])),
    ValidatedGroups = if List.IsEmpty(Groups) then error "No role groups are enabled for Effort Management Analysis."
        else Groups
in
    List.Buffer(ValidatedGroups),
#"Demand Role Analysis Prepare" = let
    Selected = Table.SelectColumns(#"LINK RoleAnalysis", {"Roles", "Effort Management Analysis"}),
    Normalised = Table.TransformColumns(Selected, {{"Roles", each if _ = null then null else Text.Trim(_), type nullable text}}),
    InvalidRows = Table.SelectRows(Normalised, each [Roles] = null or [Roles] = ""
        or ([Effort Management Analysis] <> null and not Value.Is([Effort Management Analysis], type logical))),
    Counts = Table.Group(Normalised, {"Roles"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    Duplicates = Table.SelectRows(Counts, each [Count] <> 1),
    Result = if not Table.IsEmpty(InvalidRows) then
            error Error.Record("Invalid role analysis", "Roles must be non-blank and effort flags logical or null.", InvalidRows)
        else if not Table.IsEmpty(Duplicates) then
            error Error.Record("Duplicate role analysis", "Each Roles value must occur once.", Duplicates)
        else Normalised,
    BufferedAnalysis = Table.Buffer(Result)
in
    BufferedAnalysis,
#"LINK RoleAnalysis" = Inputs[Analysis],
#"Demand Calendar CHECK" = let
    Periods = #"Demand Planning Periods",
    ExpectedPeriodIDs = List.Sort(Periods[Period]),
    ExpectedPeriods = List.Count(ExpectedPeriodIDs),
    SettingsCalendar = #"IMPORT PermutationDimensions",
    PeriodText = (IDs as list) as text => Text.Combine(List.Transform(IDs, each Text.From(_, "en-AU")), ", "),
    RoleChecks = List.Transform(#"Demand Roles", (RoleName) =>
        let
            Rows = Table.SelectRows(SettingsCalendar, each [RolesList] = RoleName),
            Counts = Table.Group(Rows, {"Period"}, {{"Rows", each Table.RowCount(_), Int64.Type}}),
            ActualRows = Table.RowCount(Rows),
            DistinctPeriods = Table.RowCount(Counts),
            MissingIDs = List.Difference(ExpectedPeriodIDs, Counts[Period]),
            DuplicateIDs = List.Sort(Table.SelectRows(Counts, each [Rows] > 1)[Period]),
            MissingPeriods = List.Count(MissingIDs),
            DuplicateRows = ActualRows - DistinctPeriods,
            Passed = MissingPeriods = 0 and DuplicateRows = 0,
            Message = if Passed then "Complete Settings calendar coverage."
                else if ActualRows = 0 then "Enabled role is absent from Settings PermutationDimensions. Check the Settings Roles query and saved calendar."
                else "Settings calendar has missing or duplicate role/period rows. Inspect the listed period IDs."
        in [Severity=if Passed then "Pass" else "Error", Role=RoleName,
            ExpectedPeriods=ExpectedPeriods, ActualRows=ActualRows, DistinctPeriods=DistinctPeriods,
            MissingPeriods=MissingPeriods, DuplicateRows=DuplicateRows,
            MissingPeriodIDs=PeriodText(MissingIDs), DuplicatePeriodIDs=PeriodText(DuplicateIDs), Message=Message]),
    Diagnostics = Table.FromRecords(RoleChecks),
    BufferedDiagnostics = Table.Buffer(Diagnostics)
in
    BufferedDiagnostics,
#"IMPORT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type",
#"IMPORT Settings Data" = Inputs[Settings],
#"Demand Planning Periods" = let
    Selected = Table.SelectColumns(#"IMPORT PermutationDimensions", {"Date", "Day", "Shifts", "Period"}),
    Periods = Table.Distinct(Table.RenameColumns(Selected, {{"Shifts", "Shift"}})),
    Dates = List.Sort(List.Distinct(Periods[Date])),
    DatesValid = not List.IsEmpty(Dates) and not List.Contains(Dates, null),
    Start = if DatesValid then Dates{0} else error "Settings must provide non-null planning dates.",
    DayCount = Duration.Days(List.Last(Dates) - Start) + 1,
    ExpectedDates = List.Dates(Start, DayCount, #duration(1, 0, 0, 0)),
    // The source pattern contains AM, PM and NS (mapped to NIGHT). The planning horizon
    // may have any positive number of days, including a final partial fortnight.
    PatternShifts = {"AM", "PM", "NIGHT"},
    ExpectedPeriodCount = DayCount * List.Count(PatternShifts),
    InvalidRows = Table.SelectRows(Periods, each [Day] <> Duration.Days([Date] - Start) + 1
        or not List.Contains(PatternShifts, [Shift]) or [Period] = null),
    Validated = if Dates <> ExpectedDates or Date.DayOfWeek(Start, Day.Monday) <> 0 then
            error "The Settings planning calendar must be consecutive and start on the source Week-1 Monday."
        else if Table.RowCount(Periods) <> ExpectedPeriodCount or
            Table.RowCount(Table.Distinct(Table.SelectColumns(Periods, {"Date", "Shift"}))) <> ExpectedPeriodCount then
            error "Settings must provide one shared period for each planning date and source shift."
        else if List.Sort(List.Distinct(Periods[Period])) <> {1..ExpectedPeriodCount} then
            error "Settings shift periods must be unique and sequential across the configured date range."
        else if not Table.IsEmpty(InvalidRows) then error "Invalid sequential Day, Shift or Period in settings."
        else Table.Sort(Periods, {{"Period", Order.Ascending}}),
    BufferedPeriods = Table.Buffer(Validated)
in
    BufferedPeriods,
#"Distributed FTE Exclusions" = let
    WithReason = Table.AddColumn(#"Distributed FTE Demand Prepare", "ExclusionReason", each
        if [Effort Management Analysis] <> true then "Effort Management Analysis is not enabled"
        else null, type nullable text),
    Excluded = Table.SelectRows(WithReason, each [ExclusionReason] <> null),
    Result = Table.Group(Excluded, {"Facility", "SourceRole", "ExclusionReason"}, {
        {"FortnightRosterHours", each List.Sum([DemandHRS]), type number},
        {"FortnightProductiveHours", each if List.NonNullCount([ProductiveHRS]) = 0 then null
            else List.Sum([ProductiveHRS]), type nullable number}
    })
in
    Result,
#"Distributed FTE Demand Prepare" = let
    SourceDemand = #"Distributed Demand FTE Scope",
    PreparedDemand = #"Distributed FTE Rows Prepare"(SourceDemand, false)
in
    PreparedDemand,
#"Distributed FTE Rows Prepare" = (Source as table, IsProfile as logical) as table =>
let
    RequiredColumns = {"Facility", "Role", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex",
        "Shift", "DemandHRS", "ProductiveHRS", "DemandBasis"},
    Selected = if #"Distributed FTE Source Validate" = "RoleDemand" then
            Table.SelectColumns(Source, RequiredColumns) else error "Invalid role-demand handover.",
    Renamed = Table.RenameColumns(Selected, {{"Role", "SourceRole"}, {"Shift", "SourceShift"}}),
    WithRoleKey = Table.AddColumn(Renamed, "SourceRoleKey",
        each if Value.Is([SourceRole], type text) then #"Demand Role Key"([SourceRole]) else null, type nullable text),
    AnalysisKeys = Table.AddColumn(#"Demand Role Analysis Prepare", "SourceRoleKey",
        each #"Demand Role Key"([Roles]), type nullable text),
    JoinedRoles = Table.NestedJoin(WithRoleKey, {"SourceRoleKey"}, AnalysisKeys, {"SourceRoleKey"}, "RoleControl", JoinKind.LeftOuter),
    MissingMappings = Table.SelectRows(JoinedRoles, each Table.RowCount([RoleControl]) <> 1),
    ValidatedMappings = if not Table.IsEmpty(MissingMappings) then
            error Error.Record("Missing demand role control", "Every published Role Group needs one current Role Analysis setting.",
                Table.Distinct(Table.SelectColumns(MissingMappings, {"Facility", "SourceRole"})))
        else JoinedRoles,
    WithRole = Table.ExpandTableColumn(ValidatedMappings, "RoleControl",
        {"Roles", "Effort Management Analysis"}, {"Role", "Effort Management Analysis"}),
    WithShift = Table.AddColumn(WithRole, "Shift", each [SourceShift], type text),
    WithUnit = Table.AddColumn(WithShift, "Facility-Abbrev", each #"Demand Facility", type text),
    InvalidRows = Table.SelectRows(WithUnit, each
        not Value.Is([Facility], type text) or [Facility] = "" or [SourceRoleKey] = null or [SourceRoleKey] = ""
        or not #"Demand Finite Number"([Week No]) or Number.RoundDown([Week No]) <> [Week No]
        or not List.Contains({1, 2}, [FortnightWeek]) or not List.Contains({1..7}, [DayOfWeek])
        or not List.Contains({1..14}, [FortnightDayIndex])
        or [FortnightDayIndex] <> ([FortnightWeek] - 1) * 7 + [DayOfWeek]
        or not List.Contains({"AM", "PM", "NIGHT"}, [Shift])
        or not #"Demand Finite Number"([DemandHRS]) or [DemandHRS] < 0
        or (if [Role] = "CS" then [DemandBasis] <> "MasterRoster" or [ProductiveHRS] <> null
            else [DemandBasis] <> "CareMinutes" or not #"Demand Finite Number"([ProductiveHRS]) or [ProductiveHRS] < 0)),
    Keys = {"Facility", "Role", "FortnightDayIndex", "Shift"},
    DuplicateKeys = Table.RowCount(WithUnit) <> Table.RowCount(Table.Distinct(Table.SelectColumns(WithUnit, Keys))),
    Validated = if not Table.IsEmpty(InvalidRows) then
            error Error.Record("Invalid combined demand cells", "Invalid source keys, net hours or care/CS demand basis.", InvalidRows)
        else if DuplicateKeys then error "Duplicate facility/role/fortnight-day/shift keys in RoleDemand_TABLE."
        else WithUnit
in
    Table.Buffer(Validated),
#"Demand Finite Number" = (Value as any) as logical =>
    if not Value.Is(Value, type number) then false
    else not Number.IsNaN(Value) and Number.Abs(Value) <> #infinity,
#"Demand Role Key" = (RoleName as nullable text) as nullable text =>
    if RoleName = null then null else Text.Upper(Text.Trim(RoleName)),
#"Distributed FTE Source Validate" = let
    Demand = #"Distributed Demand FTE Scope",
    RequiredColumns = {"Facility", "Role", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex",
        "Shift", "DemandHRS", "ProductiveHRS", "DemandBasis"},
    RequiredColumnsPresent = Table.HasColumns(Demand, RequiredColumns),
    Checks = #"Distributed Demand FTE Checks Scope",
    InvalidChecks = Table.SelectRows(Checks, each not List.Contains({"Pass", "Warning"}, [Severity])),
    Result = if not RequiredColumnsPresent then
            error "RoleDemand_TABLE is missing required combined-handover columns. Refresh and save the organisation Manual Read handover."
        else if Table.IsEmpty(Checks) or not Table.IsEmpty(InvalidChecks) then
            error Error.Record("Upstream validation failed", "Inspect the saved RoleDemand_CHECK table.", InvalidChecks)
        else "RoleDemand"
in
    Result,
#"Distributed Demand FTE Checks Scope" = let
    Source = #"EXTRACT Distributed Demand FTE Checks",
    CheckColumns = Table.SelectColumns(Source, {"Severity", "Check", "Facility", "Role", "Message"}),
    ScopedChecks = Table.SelectRows(CheckColumns, each [Facility] = null
        or (Value.Is([Facility], type text)
            and Comparer.OrdinalIgnoreCase(Text.Trim([Facility]), #"Demand Facility") = 0)),
    BufferedChecks = Table.Buffer(ScopedChecks)
in
    BufferedChecks,
#"EXTRACT Distributed Demand FTE Checks" = let
    WorkbookNavigation = #"IMPORT Distributed Demand",
    SourceChecks = WorkbookNavigation{[Item="RoleDemand_CHECK", Kind="Table"]}[Data]
in
    SourceChecks,
#"IMPORT Distributed Demand" = Inputs[DemandNavigation],
#"Distributed Demand FTE Scope" = let
    Source = #"EXTRACT Distributed Demand FTE",
    ScopedDemand = if Table.HasColumns(Source, {"Facility"}) then
            Table.SelectRows(Source, each Value.Is([Facility], type text)
                and Comparer.OrdinalIgnoreCase(Text.Trim([Facility]), #"Demand Facility") = 0)
        else Source,
    BufferedDemand = Table.Buffer(ScopedDemand)
in
    BufferedDemand,
#"EXTRACT Distributed Demand FTE" = let
    WorkbookNavigation = #"IMPORT Distributed Demand",
    DemandTable = WorkbookNavigation{[Item="RoleDemand_TABLE", Kind="Table"]}[Data]
in
    DemandTable,
#"Distributed FTE Profile Prepare" = #"Distributed FTE Demand Prepare",
#"Distributed FTE Source Cells" = let
    Prepared = #"Distributed FTE Demand Prepare",
    EnabledUnit = not Table.IsEmpty(FacilityAnalysisTABLE),
    Included = Table.SelectRows(Prepared, each EnabledUnit and [Effort Management Analysis] = true),
    Coverage = Table.Group(Included, {"Role"}, {{"CellCount", each Table.RowCount(_), Int64.Type}}),
    MissingGroups = List.Difference(#"Demand Roles", List.Distinct(Included[Role])),
    UnexpectedGroups = List.Difference(List.Distinct(Included[Role]), #"Demand Roles"),
    InvalidCoverage = Table.SelectRows(Coverage, each [CellCount] <> 42),
    PatternWeeks = Table.Distinct(Table.SelectColumns(Included, {"FortnightWeek", "Week No"})),
    ValidWeeks = Table.RowCount(PatternWeeks) = 2
        and List.Sort(PatternWeeks[FortnightWeek]) = {1, 2}
        and List.Count(List.Distinct(PatternWeeks[Week No])) = 2
        and (if Table.RowCount(PatternWeeks) = 2 then
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){0}[Week No] <
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){1}[Week No] else false),
    Validated = if not EnabledUnit then error "Facility Analysis does not enable this Unit."
        else if not List.IsEmpty(MissingGroups) or not List.IsEmpty(UnexpectedGroups) or not Table.IsEmpty(InvalidCoverage) then
            error Error.Record("Incomplete combined role demand", "Every enabled Role Group requires 42 source cells. Missing groups: "
                & Text.Combine(MissingGroups, ", ") & ". Refresh and save RoleDemand_TABLE before this unit extract.", Coverage)
        else if not ValidWeeks then error "Role demand must retain one consistent pair of ordered source weeks."
        else Included
in
    Table.Buffer(Validated),
FacilityAnalysisTABLE = let
    Source = #"LINK Facility Analysis",
    UnitRows = Table.SelectRows(Source, each [Title] <> null and Comparer.OrdinalIgnoreCase(Text.Trim([Title]), #"Demand Facility") = 0),
    ValidatedUnit = if Table.RowCount(UnitRows) = 1 then UnitRows else error "Facility Analysis must identify exactly one row for the current Unit.",
    EnabledUnit = Table.SelectRows(ValidatedUnit, each [EffortManagementAnalysis] = true),
    Result = Table.RemoveColumns(EnabledUnit, {"LeaveBalanceAnalysis"})
in
    Result,
#"LINK Facility Analysis" = Inputs[Facilities],
#"Distributed FTE Prepare" = let
    Source = #"Distributed FTE Source Cells",
    WithSourceRoles = Table.AddColumn(Source, "SourceRoles", each [SourceRole], type text),
    Pattern = Table.SelectColumns(WithSourceRoles,
        {"Facility", "Role", "FortnightDayIndex", "Shift", "Week No", "SourceRoles", "DemandHRS", "ProductiveHRS", "DemandBasis"}),
    ExpectedCells = List.Count(#"Demand Roles") * 42,
    Validated = if Table.RowCount(Pattern) <> ExpectedCells then error "Incomplete enabled role-group source pattern." else Pattern
in
    Table.Buffer(Validated),
DemandExtraction_RECONCILIATION = let
    Targets = #"Demand Reconciliation Targets",
    Actual = Table.Group(#"Demand Extraction Prepare", {"Facility", "Role", "PlanningFortnight"}, {
        {"CellCount", each Table.RowCount(_), Int64.Type},
        {"RosterHours", each List.Sum([DemandHRS]), type number},
        {"IntegratedHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[DemandHRS] * (Duration.TotalHours(R[EndTime] - R[StartTime]) / R[DurationOfShifts]))), type number},
        {"ProductiveHours", each if List.NonNullCount([ProductiveHRS]) = 0 then null
            else List.Sum([ProductiveHRS]), type nullable number}
    }),
    Keys = {"Facility", "Role", "PlanningFortnight"},
    Joined = Table.NestedJoin(Targets, Keys, Actual, Keys, "Actual", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Actual", {"CellCount", "RosterHours", "IntegratedHours", "ProductiveHours"}),
    WithResidual = Table.AddColumn(Expanded, "RosterHoursResidual",
        each if #"Demand Finite Number"([RosterHours]) then [RosterHours] - [ExpectedRosterHours] else null, type nullable number),
    WithProductiveMatch = Table.AddColumn(WithResidual, "ProductiveHoursMatch", each
        if [ExpectedProductiveHours] = null then [ProductiveHours] = null
        else #"Demand Finite Number"([ProductiveHours])
            and Number.Abs([ProductiveHours] - [ExpectedProductiveHours]) <= 0.0000001, type logical),
    WithStatus = Table.AddColumn(WithProductiveMatch, "Status", each
        if [CellCount] = [ExpectedCellCount] and #"Demand Finite Number"([RosterHoursResidual])
            and Number.Abs([RosterHoursResidual]) <= 0.0000001 and #"Demand Finite Number"([IntegratedHours])
            and Number.Abs([IntegratedHours] - [ExpectedRosterHours]) <= 0.0000001 and [ProductiveHoursMatch]
        then "Pass" else "Error", type text),
    Result = Table.RemoveColumns(WithStatus, {"ProductiveHoursMatch"})
in
    Table.Buffer(Result),
#"Demand Extraction Prepare" = let
    Joined = Table.NestedJoin(#"Permutation DateTimeRoleShift", {"Role", "FortnightDayIndex", "Shift"},
        #"Distributed FTE Prepare", {"Role", "FortnightDayIndex", "Shift"}, "Pattern", JoinKind.LeftOuter),
    InvalidMatches = Table.SelectRows(Joined, each Table.RowCount([Pattern]) <> 1),
    Validated = if not Table.IsEmpty(InvalidMatches) then
        error "Every calendar cell must match exactly one validated fortnight demand cell." else Joined,
    Expanded = Table.ExpandTableColumn(Validated, "Pattern",
        {"Facility", "SourceRoles", "DemandHRS", "ProductiveHRS", "DemandBasis", "Week No"},
        {"Facility", "SourceRoles", "DemandHRS", "ProductiveHRS", "DemandBasis", "Week No"}),
    WithUnit = Table.AddColumn(Expanded, "Unit", each [Facility], type text),
    // Standard-shift spreading preserves the supplied net hours. Do not apply a care percentage or deduct breaks here.
    WithAttendance = Table.AddColumn(WithUnit, "DemandFTE", each [DemandHRS] / [DurationOfShifts], type number),
    InvalidValues = Table.SelectRows(WithAttendance, each not #"Demand Finite Number"([DemandFTE])
        or [DemandFTE] < 0 or not #"Demand Finite Number"([DemandHRS]) or [DemandHRS] < 0),
    ValidatedAttendance = if not Table.IsEmpty(InvalidValues) then error "Demand hours/attendance conversion failed."
        else WithAttendance
in
    Table.Buffer(ValidatedAttendance),
#"Permutation DateTimeRoleShift" = let
    Periods = Table.SelectRows(#"IMPORT ShiftPeriod", each List.Contains(#"Demand Roles", [Role])),
    Joined = Table.NestedJoin(#"Demand Calendar Prepare", {"Role", "Shift"}, Periods,
        {"Role", "ShiftPeriod"}, "Timing", JoinKind.LeftOuter),
    InvalidMatches = Table.SelectRows(Joined, each Table.RowCount([Timing]) <> 1),
    ValidatedMatches = if not Table.IsEmpty(InvalidMatches) then
        error "Each enabled Role Group calendar cell requires exactly one settings shift definition." else Joined,
    Expanded = Table.ExpandTableColumn(ValidatedMatches, "Timing",
        {"StartDay", "EndDay", "DurationOfShifts"}, {"Start", "End", "DurationOfShifts"}),
    TypedTimes = Table.TransformColumnTypes(Expanded, {{"Start", type time}, {"End", type time}}),
    WithStart = Table.AddColumn(TypedTimes, "StartTime", each [Date] & [Start], type datetime),
    WithEnd = Table.AddColumn(WithStart, "EndTime", each
        (if [Shift] = "NIGHT" then Date.AddDays([Date], 1) else [Date]) & [End], type datetime),
    InvalidDurations = Table.SelectRows(WithEnd, each [StartTime] = null or [EndTime] = null
        or not #"Demand Finite Number"([DurationOfShifts]) or [DurationOfShifts] <= 0
        or Number.Abs(Duration.TotalHours([EndTime] - [StartTime]) - [DurationOfShifts]) > 0.0000001),
    Result = if not Table.IsEmpty(InvalidDurations) then
        error Error.Record("Invalid shift span", "Settings duration must equal EndTime minus StartTime.", InvalidDurations)
        else Table.RemoveColumns(WithEnd, {"Start", "End"}),
    BufferedTiming = Table.Buffer(Result)
in
    BufferedTiming,
#"Demand Calendar Prepare" = let
    Periods = #"Demand Planning Periods",
    Start = List.Min(Periods[Date]),
    SelectedRoles = Table.SelectRows(#"IMPORT PermutationDimensions",
        each List.Contains(#"Demand Roles", [RolesList])),
    Calendar = Table.RenameColumns(SelectedRoles, {{"RolesList", "Role"}, {"Shifts", "Shift"}}),
    ExpectedCells = List.Count(#"Demand Roles") * Table.RowCount(Periods),
    Coverage = #"Demand Calendar CHECK",
    MissingOrDuplicate = Table.SelectRows(Coverage, each [Severity] = "Error"),
    SettingsRoleLabels = List.Sort(List.Distinct(#"IMPORT PermutationDimensions"[RolesList])),
    CoverageSummary = Text.Combine(List.Transform(Table.ToRecords(MissingOrDuplicate), each
        [Role] & ": expected " & Text.From([ExpectedPeriods], "en-AU") & " periods, found "
        & Text.From([DistinctPeriods], "en-AU") & " (missing " & Text.From([MissingPeriods], "en-AU")
        & "; duplicate rows " & Text.From([DuplicateRows], "en-AU") & ")"), "; "),
    // Keep failed coverage visible: never remove an enabled role or invent its Settings rows.
    Validated = if Table.RowCount(Calendar) <> ExpectedCells or not Table.IsEmpty(MissingOrDuplicate) then
            error Error.Record("Incomplete demand calendar", CoverageSummary
                & ". Settings role labels: " & Text.Combine(List.Transform(SettingsRoleLabels,
                    each if _ = null then "<null>" else "[" & _ & "]"), ", ")
                & ". Inspect Demand Calendar CHECK for the affected roles and period IDs.", MissingOrDuplicate)
        else Calendar,
    // Day remains the Settings sequential index. Only the source-pattern key repeats 1..14.
    WithFortnightDay = Table.AddColumn(Validated, "FortnightDayIndex",
        each Number.Mod(Duration.Days([Date] - Start), 14) + 1, Int64.Type),
    WithCycle = Table.AddColumn(WithFortnightDay, "PlanningFortnight",
        each Number.IntegerDivide(Duration.Days([Date] - Start), 14) + 1, Int64.Type),
    BufferedCalendar = Table.Buffer(WithCycle)
in
    BufferedCalendar,
#"IMPORT ShiftPeriod" = let
    Source = #"IMPORT Settings Data",
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}, {"DurationOfShifts", type number}})
in
    #"Changed Type",
#"Demand Reconciliation Targets" = let
    Joined = Table.NestedJoin(#"Demand Calendar Prepare", {"Role", "FortnightDayIndex", "Shift"},
        #"Distributed FTE Source Cells", {"Role", "FortnightDayIndex", "Shift"}, "SourceCells", JoinKind.LeftOuter),
    MissingCells = Table.SelectRows(Joined, each Table.RowCount([SourceCells]) <> 1),
    Validated = if not Table.IsEmpty(MissingCells) then error "A planning period has no unique source-hour reconciliation cell." else Joined,
    WithFacility = Table.AddColumn(Validated, "Facility", each #"Demand Facility", type text),
    WithRosterHours = Table.AddColumn(WithFacility, "DemandHRS", each List.Sum([SourceCells][DemandHRS]), type number),
    WithProductiveHours = Table.AddColumn(WithRosterHours, "ProductiveHRS",
        each if List.NonNullCount([SourceCells][ProductiveHRS]) = 0 then null
            else List.Sum([SourceCells][ProductiveHRS]), type nullable number),
    Targets = Table.Group(WithProductiveHours, {"Facility", "Role", "PlanningFortnight"}, {
        {"ExpectedCellCount", each Table.RowCount(_), Int64.Type},
        {"ExpectedRosterHours", each List.Sum([DemandHRS]), type number},
        {"ExpectedProductiveHours", each if List.NonNullCount([ProductiveHRS]) = 0 then null
            else List.Sum([ProductiveHRS]), type nullable number}
    })
in
    Table.Buffer(Targets),
DemandExtraction_CHECK = let
    CheckStage = (Name as text, Work as function) as record =>
        let Attempt = try Work()
        in [Severity = if Attempt[HasError] then "Error" else "Pass", Check = Name,
            Message = if Attempt[HasError] then Attempt[Error][Message] else "Validated",
            Actual = if Attempt[HasError] then null else Attempt[Value]],
    StageChecks = Table.FromRecords({
        CheckStage("Configured facility and enabled unit", () =>
            let
                Facility = #"Demand Facility",
                EnabledUnit = FacilityAnalysisTABLE
            in
                if Table.IsEmpty(EnabledUnit) then error "Facility Analysis does not enable this Unit."
                else Facility),
        CheckStage("Role-group effort-analysis controls", () => Table.RowCount(#"Demand Role Analysis Prepare")),
        CheckStage("Enabled effort-management role groups", () => List.Count(#"Demand Roles")),
        CheckStage("Source schema and saved publication checks", () => #"Distributed FTE Source Validate"),
        CheckStage("Complete combined role-group source pattern", () => Table.RowCount(#"Distributed FTE Source Cells")),
        CheckStage("Complete enabled role-group fortnight patterns", () => Table.RowCount(#"Distributed FTE Prepare")),
        CheckStage("Complete role coverage across Settings planning periods", () => Table.RowCount(#"Demand Calendar Prepare")),
        CheckStage("Unique timing and exact shift spans", () => Table.RowCount(#"Permutation DateTimeRoleShift")),
        CheckStage("Complete demand and fractional attendance conversion", () => Table.RowCount(#"Demand Extraction Prepare")),
        CheckStage("Retained hours across all full and partial planning fortnights", () =>
            let Reconciliation = DemandExtraction_RECONCILIATION,
                FortnightCount = List.Count(List.Distinct(#"Demand Calendar Prepare"[PlanningFortnight]))
            in if Table.RowCount(Reconciliation) <> List.Count(#"Demand Roles") * FortnightCount or
                not Table.IsEmpty(Table.SelectRows(Reconciliation, each [Status] <> "Pass")) then
                error "Fortnight roster/productive/integrated hours do not reconcile."
            else Table.RowCount(Reconciliation))
    }),
    UpstreamAttempt = try Table.Buffer(#"Distributed Demand FTE Checks Scope"),
    Warnings = if UpstreamAttempt[HasError] then #table({"Severity", "Check", "Message", "Actual"}, {}) else
        Table.AddColumn(Table.SelectColumns(
            Table.SelectRows(UpstreamAttempt[Value], each [Severity] = "Warning"),
            {"Severity", "Check", "Message"}), "Actual", each null),
    CombinedChecks = Table.Combine({StageChecks, Warnings})
in
    CombinedChecks,
ShiftUnitDemandHRS = let
    Checks = Table.Buffer(DemandExtraction_CHECK),
    Failures = Table.SelectRows(Checks, each [Severity] = "Error"),
    Validated = if not Table.IsEmpty(Failures) then
        error Error.Record("Demand Extraction validation failed",
            Text.Combine(List.Transform(Table.ToRecords(Failures), each [Check] & ": " & [Message]), "#(lf)"),
            Failures)
        else #"Demand Extraction Prepare",
    PublishedColumns = Table.SelectColumns(Validated,
        {"Date", "Day", "Shift", "Period", "Role", "StartTime", "EndTime", "Unit",
            "Facility", "DemandFTE", "DemandHRS", "DurationOfShifts"}),
    Result = Table.Sort(PublishedColumns, {{"Period", Order.Ascending}, {"Role", Order.Ascending}})
in
    Result
in [Demand=ShiftUnitDemandHRS,Checks=DemandExtraction_CHECK,Reconciliation=DemandExtraction_RECONCILIATION,
    Pattern=#"Distributed FTE Prepare",Cells=#"Distributed FTE Source Cells",Profile=#"Distributed FTE Profile Prepare",
    Exclusions=#"Distributed FTE Exclusions",CalendarCheck=#"Demand Calendar CHECK",
    Roles=#"Demand Roles",Facility=#"Demand Facility",UnitPath=Unit1Path],
Fixtures = (Evaluate as function) as table => let
    Near=(A as number,B as number) as logical => Number.Abs(A-B)<=0.0000001,
    Change=(Base as record,Fields as record) as record => Record.Combine({Base,Fields}),
    Set=(Row as record,Field as text,Value as any) as record => Record.TransformFields(Row,{{Field,each Value}}),
    UpdateFirst=(T as table,Field as text,Value as any) as table => Table.FromRecords({Set(T{0},Field,Value)} & List.Skip(Table.ToRecords(T))),
    Check=(Name as text,Work as function) as record => let R=try Work(),P=if R[HasError] then false else R[Value]=true
        in [Case=Name,Passed=P,Details=if R[HasError] then R[Error][Message] else if P then null else "Unexpected result"],
    Groups={"RN","AIN","EN","CS"},
    BaseHours=(R as text) as number => if R="RN" then 2 else if R="AIN" then 3 else if R="EN" then 4 else 5,
    Percentage=(R as text) as nullable number => if R="CS" then null else if R="EN" then 0.5 else 1,
    Amount=(R as text,D as number,S as number) as number =>
        if R="CS" and D=3 and S=2 then 0 else BaseHours(R)*(if D<=7 then 1 else 2)+Number.Mod(D-1,7)/10+S/100,
    Handover=Table.FromRecords(List.Combine(List.Transform(Groups,(R)=>List.Combine(List.Transform({1..14},(D)=>
        List.Transform({1..3},(S)=>let H=Amount(R,D,S),W=1+Number.IntegerDivide(D-1,7)
            in [Facility="BD",Role=R,#"Week No"={17,41}{W-1},FortnightWeek=W,DayOfWeek=Number.Mod(D-1,7)+1,
                FortnightDayIndex=D,Shift={"AM","PM","NIGHT"}{S-1},DemandHRS=H,
                ProductiveHRS=if R="CS" then null else H*Percentage(R),DemandBasis=if R="CS" then "MasterRoster" else "CareMinutes"]
    )))))),
    HandoverChecks=#table({"Check","Severity","Facility","MinuteCategory","Role","Actual","Expected","Message"},
        {{"Synthetic generic handover validated","Pass","BD",null,null,168,168,"Synthetic pass"}}),
    Navigation=(D as table,C as table) as table => #table({"Item","Kind","Data"},{{"RoleDemand_TABLE","Table",D},{"RoleDemand_CHECK","Table",C}}),
    Calendar=(Days as number) as table => Table.FromRecords(List.Combine(List.Transform(List.Numbers(0,Days),(D)=>
        List.Combine(List.Transform({1..3},(S)=>List.Transform(Groups,(R)=>
            [Date=Date.AddDays(#date(2026,9,28),D),Day=D+1,Shifts={"AM","PM","NIGHT"}{S-1},Period=D*3+S,RolesList=R])))))),
    Timing=Table.FromRecords(List.Combine(List.Transform(Groups,(R)=>List.Transform({1..3},(S)=>
        [Role=R,ShiftPeriod={"AM","PM","NIGHT"}{S-1},StartTime={0.25,14/24,21/24}{S-1},
            StartDay={0.25,14/24,21/24}{S-1},EndDay={14.25/24,21.25/24,0.25}{S-1},DurationOfShifts={8.25,7.25,9}{S-1}])))),
    Settings=(Days as number,Cal as nullable table,Times as nullable table) as table => #table({"Item","Kind","Data"},{
        {"ShiftDuration","Table",#table({"ShiftDuration"},{{7.6}})},
        {"ShiftStart","Table",#table({"StartAM","StartPM","StartNIGHT"},{{0.25,14/24,21/24}})},
        {"PermutationDimensions","Table",if Cal=null then Calendar(Days) else Cal},
        {"ShiftPeriod","Table",if Times=null then Timing else Times}}),
    Roles=#table({"Roster Roles","DC Role","Role Group","DC Category","Direct Care %"},{
        {"Registered Nurse","RN","RN","RN",1},{"Assistant in Nursing","AIN","AIN","OTHER",1},
        {"Enrolled Nurse","EN","EN","OTHER",0.5},{"Care Services Employee",null,"CS","NA",null}}),
    Analysis=#table({"Roles","Effort Management Analysis","Leave Balance Analysis"},List.Transform(Groups,each {_,true,true})),
    Facilities=#table({"Title","EffortManagementAnalysis","LeaveBalanceAnalysis"},{{"BD",true,false},{"TE",true,false}}),
    Inputs=[DemandNavigation=Navigation(Handover,HandoverChecks),Settings=Settings(28,null,null),Roles=Roles,Analysis=Analysis,Facilities=Facilities,
        PathNavigation=#table({"Item","Kind","Data"},{{"CentriSyncPaths","Table",#table({"SharepointRootUrl","SyncedFolderRootPath"},{{"https://fixture.invalid/ResidentialCare","Z:\ResidentialCare"}})}}),
        PathInput=#table({"FilePath"},{{"Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\1. Input\[2-DemandExtract.xlsx]Path"}})],
    WithHandover=(D as table,C as table) as record => Change(Inputs,[DemandNavigation=Navigation(D,C)]),
    WithSettings=(Cal as nullable table,Times as nullable table) as record => Change(Inputs,[Settings=Settings(28,Cal,Times)]),
    Result=Evaluate(Inputs),Demand=Table.Buffer(Result[Demand]),Checks=Table.Buffer(Result[Checks]),
    Pattern=Table.Buffer(Result[Pattern]),Cells=Table.Buffer(Result[Cells]),Reconciliation=Table.Buffer(Result[Reconciliation]),
    ExpectedColumns={"Date","Day","Shift","Period","Role","StartTime","EndTime","Unit","Facility","DemandFTE","DemandHRS","DurationOfShifts"},
    Sorted=(T as table) as table => Table.Sort(Table.SelectColumns(T,ExpectedColumns),{{"Period",Order.Ascending},{"Role",Order.Ascending}}),
    Reject=(I as record) as logical => let R=Evaluate(I),C=try Table.Buffer(R[Checks]),D=try Table.RowCount(Table.Buffer(R[Demand]))
        in not C[HasError] and Table.RowCount(Table.SelectRows(C[Value],each [Severity]="Error"))>0 and D[HasError],
    ValidHours=(T as table) as logical => List.AllTrue(List.Transform(Table.ToRecords(T),each
        Near([DemandHRS],Amount([Role],Number.Mod([Day]-1,14)+1,List.PositionOf({"AM","PM","NIGHT"},[Shift])+1)))),
    ExpectedTotal=(R as text,Days as number) as number => List.Sum(List.Combine(List.Transform(List.Numbers(0,Days),(D)=>List.Transform({1..3},(S)=>Amount(R,Number.Mod(D,14)+1,S))))),
    AddCheck=(Facility as nullable text,Severity as text) as table => Table.Combine({HandoverChecks,
        #table(Table.ColumnNames(HandoverChecks),{{"Synthetic handover issue",Severity,Facility,null,"CS",1,0,"Deliberate issue"}})}),
    CSRows=Table.SelectRows(Handover,each [Role]="CS"),CareRows=Table.SelectRows(Handover,each [Role]<>"CS"),
    Rows={
        Check("Twelve-column downstream publication is unchanged",()=>Table.ColumnNames(Demand)=ExpectedColumns),
        Check("All four roles have one row per Settings period",()=>Table.RowCount(Demand)=336 and Table.RowCount(Table.Distinct(Demand,{"Role","Period"}))=336 and List.Sort(List.Distinct(Demand[Role]))={"AIN","CS","EN","RN"}),
        Check("Publication checks reconciliation and legacy helper aliases are forced without errors",()=>List.AllTrue(List.Transform(
            {Demand,Checks,Pattern,Cells,Reconciliation,Result[Profile],Result[Exclusions],Result[CalendarCheck]},
            each Table.RowCount(Table.SelectRowsWithErrors(Table.Buffer(_)))=0))
            and Table.RowCount(Table.SelectRows(Checks,each [Severity]="Error"))=0),
        Check("All care and CS hour cells retain the exact unequal source pattern",()=>ValidHours(Demand)),
        Check("Source week numbers are preserved rather than renumbered",()=>List.Sort(List.Distinct(Cells[Week No]))={17,41} and Table.RowCount(Table.SelectRows(Cells,each ([FortnightWeek]=1 and [Week No]<>17) or ([FortnightWeek]=2 and [Week No]<>41)))=0),
        Check("CS productive hours remain null in source and grouped pattern",()=>List.NonNullCount(Table.SelectRows(Cells,each [Role]="CS")[ProductiveHRS])=0 and List.NonNullCount(Table.SelectRows(Pattern,each [Role]="CS")[ProductiveHRS])=0),
        Check("Care productive hours remain numeric and unchanged",()=>List.AllTrue(List.Transform(Table.ToRecords(Table.SelectRows(Cells,each [Role]<>"CS")),each Value.Is([ProductiveHRS],type number) and Near([ProductiveHRS],[DemandHRS]*Percentage([Role]))))),
        Check("CS roster hours are conserved across two planning fortnights",()=>Near(List.Sum(Table.SelectRows(Demand,each [Role]="CS")[DemandHRS]),ExpectedTotal("CS",28))),
        Check("FTE uses each actual shift duration and round-trips to hours",()=>List.AllTrue(List.Transform(Table.ToRecords(Demand),each Near([DemandFTE],[DemandHRS]/[DurationOfShifts]) and Near([DemandFTE]*[DurationOfShifts],[DemandHRS])))),
        Check("Overnight end timestamps advance exactly one date",()=>List.AllTrue(List.Transform(Table.ToRecords(Demand),each Near(Duration.TotalHours([EndTime]-[StartTime]),[DurationOfShifts]) and (if [Shift]="NIGHT" then Date.From([EndTime])=Date.AddDays([Date],1) else Date.From([EndTime])=[Date])))),
        Check("Every full-fortnight reconciliation passes",()=>Table.RowCount(Reconciliation)=8 and List.Distinct(Reconciliation[Status])={"Pass"}),
        Check("One through thirty-five day horizons retain exact source day totals",()=>List.AllTrue(List.Transform({1,7,14,15,21,28,35},(Days)=>let R=Evaluate(Change(Inputs,[Settings=Settings(Days,null,null)])),D=Table.Buffer(R[Demand]),C=Table.Buffer(R[Reconciliation]) in Table.RowCount(D)=Days*3*4 and ValidHours(D) and List.Distinct(C[Status])={"Pass"} and List.AllTrue(List.Transform(Groups,(Role)=>Near(List.Sum(Table.SelectRows(D,each [Role]=Role)[DemandHRS]),ExpectedTotal(Role,Days))))))),
        Check("Final partial fortnight uses actual source days without proration",()=>let R=Evaluate(Change(Inputs,[Settings=Settings(15,null,null)])),C=Table.SelectRows(R[Reconciliation],each [PlanningFortnight]=2) in Table.RowCount(C)=4 and List.AllTrue(List.Transform(Table.ToRecords(C),each [ExpectedCellCount]=3 and Near([ExpectedRosterHours],ExpectedTotal([Role],1))))),
        Check("Folder-derived full BD unit key is retained",()=>Result[Facility]="BD" and Result[UnitPath]="Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD" and List.Distinct(Demand[Facility])={"BD"} and List.Distinct(Demand[Unit])={"BD"}),
        Check("Other facility and BD-prefix alternatives are excluded",()=>let Other=Table.Combine({Handover,Table.TransformColumns(Handover,{{"Facility",each "TE"}}),Table.TransformColumns(Handover,{{"Facility",each "BD-OTHER"}})}) in Sorted(Evaluate(WithHandover(Other,HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Global upstream handover errors block publication",()=>Reject(WithHandover(Handover,AddCheck(null,"Error")))),
        Check("BD upstream handover errors block publication",()=>Reject(WithHandover(Handover,AddCheck("BD","Error")))),
        Check("Other-facility upstream issues cannot block BD",()=>Sorted(Evaluate(WithHandover(Handover,AddCheck("TE","Error")))[Demand])=Sorted(Demand)),
        Check("An empty handover check publication is not a runtime pass",()=>Reject(WithHandover(Handover,Table.FirstN(HandoverChecks,0)))),
        Check("Non-null CS productive hours fail visibly",()=>Reject(WithHandover(Table.Combine({UpdateFirst(CSRows,"ProductiveHRS",1),CareRows}),HandoverChecks))),
        Check("Null care productive hours fail visibly",()=>Reject(WithHandover(Table.Combine({CSRows,UpdateFirst(CareRows,"ProductiveHRS",null)}),HandoverChecks))),
        Check("An unsupported demand basis fails visibly",()=>Reject(WithHandover(UpdateFirst(Handover,"DemandBasis","Unknown"),HandoverChecks))),
        Check("Duplicate handover cells block publication",()=>Reject(WithHandover(Table.Combine({Handover,Table.FirstN(Handover,1)}),HandoverChecks))),
        Check("Missing CS source pattern cells cannot become invented zeroes",()=>Reject(WithHandover(Table.Combine({Table.Skip(CSRows,1),CareRows}),HandoverChecks))),
        Check("Stale Settings without CS fails role coverage",()=>Reject(WithSettings(Table.SelectRows(Calendar(28),each [RolesList]<>"CS"),null))),
        Check("One missing CS Settings period remains a coverage failure",()=>Reject(WithSettings(Table.SelectRows(Calendar(28),each not ([RolesList]="CS" and [Period]=2)),null))),
        Check("Missing CS shift definitions block publication",()=>Reject(WithSettings(null,Table.SelectRows(Timing,each not ([Role]="CS" and [ShiftPeriod]="NIGHT"))))),
        Check("Mismatched duration and timestamp spans block publication",()=>Reject(WithSettings(null,UpdateFirst(Timing,"DurationOfShifts",999)))),
        Check("Missing authoritative DemandHRS cannot fall back to FTE",()=>Reject(WithHandover(Table.AddColumn(Table.RemoveColumns(Handover,{"DemandHRS"}),"FTE",each 999),HandoverChecks))),
        Check("Invalid handover hours fail without defaults",()=>List.AllTrue(List.Transform({null,"7.6",-1,#nan,#infinity},(H)=>Reject(WithHandover(UpdateFirst(Handover,"DemandHRS",H),HandoverChecks))))),
        Check("Additional display FTE cannot change authoritative hours",()=>Sorted(Evaluate(WithHandover(Table.AddColumn(Handover,"FTE",each 999),HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Legacy-only navigation is not a generic handover substitute",()=>Reject(Change(Inputs,[DemandNavigation=#table({"Item","Kind","Data"},{{"MinuteWorkersFTE_TABLE","Table",Handover},{"MinuteWorkersFTE_CHECK","Table",HandoverChecks}})]))),
        Check("Source row order cannot average or renumber source weeks",()=>Sorted(Evaluate(WithHandover(Table.ReverseRows(Handover),HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Upstream warnings remain visible without altering demand",()=>let R=Evaluate(WithHandover(Handover,AddCheck("BD","Warning"))) in Sorted(R[Demand])=Sorted(Demand) and Table.RowCount(Table.SelectRows(R[Checks],each [Severity]="Warning" and [Check]="Synthetic handover issue"))=1)
    }
in Table.FromRecords(Rows)
in Fixtures(Evaluate);
