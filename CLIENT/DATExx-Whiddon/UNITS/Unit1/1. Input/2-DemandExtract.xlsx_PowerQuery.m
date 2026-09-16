// Power Query from: 2-DemandExtract.xlsx
// Pathname: CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\2-DemandExtract.xlsx
// Extracted: 2026-08-26T10:04:05.610Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook's path through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl, either the standard one-row FilePath table or this workbook's single named cell.
// Output: Variable Name / Value rows for the existing UnitL1PathTABLE worksheet table.
// Notes: The worksheet output is not a path input; reading it here would reuse stale query results.
shared UnitL1PathTABLE =
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
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
            {"FilePathUrl", FilePath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

// Query: Unit1Path
// Purpose: Derive the unit root from the resolved workbook folder for relative imports.
shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    WorkbookFolder = #"Filtered Rows"{0}[Value],
    UnitFolder =
        if Text.EndsWith(WorkbookFolder, "\1. Input", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\1. Input"))
        else if Text.EndsWith(WorkbookFolder, "\2. Calculations", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\2. Calculations"))
        else
            error "Root Path did not end in an expected Unit1 workbook folder: " & WorkbookFolder
in
    UnitFolder;

shared LocalUnitFolder = Unit1Path;

shared StartAM = let
    Source = ShiftStart,
    StartAM1 = Source{0}[StartAM]
in
    StartAM1;

shared StartPM = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartPM]
in
    StartPM1;

shared StartNIGHT = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartNIGHT]
in
    StartPM1;

shared ShiftOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t, Column2 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "Shift"}, {"Column2", "Order"}})
in
    #"Renamed Columns";

// Query: Demand Facility
// Purpose: Select the population supplied by the existing Beaudesert allocation input.
// Notes: One facility prevents duplicate period demand in the unchanged A.1 models.
shared #"Demand Facility" = "BD";

// Query: Demand Roles
// Purpose: Retain RN and AIN and supply the existing AINC4 branch with enrolled-nurse demand.
shared #"Demand Roles" = {"RN", "AIN", "AINC4"};

// Query: Demand Role Code
// Purpose: Read the known published role labels, using AINC4 as the enrolled-nurse role value.
shared #"Demand Role Code" = (RoleName as nullable text) as nullable text =>
    let Key = if RoleName = null then "" else Text.Upper(Text.Trim(RoleName))
    in if List.Contains({"RN", "REGISTERED NURSE"}, Key) then "RN"
       else if List.Contains({"AIN", "ASSISTANT IN NURSING"}, Key) then "AIN"
       // Use the existing third-role value directly; RN and AIN demand values are unchanged.
       else if List.Contains({"EN", "ENROLLED NURSE", "AINC4"}, Key) then "AINC4"
       else null;

// Query: Demand Finite Number
// Purpose: Reject nulls, text, NaN and infinities before demand arithmetic.
shared #"Demand Finite Number" = (Value as any) as logical =>
    if not Value.Is(Value, type number) then false
    else not Number.IsNaN(Value) and Number.Abs(Value) <> #infinity;

// Query: IMPORT Distributed FTE Workbook
// Purpose: Share one saved-workbook snapshot across allocation, profile and validation imports.
// Notes: Buffer the binary as well as the navigation table; do not refresh or modify the source.
shared #"IMPORT Distributed FTE Workbook" =
    Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
        Unit1Path & "\\1. Input\\Demand-MasterRoster Manual Read.xlsx")), null, true));

// Query: IMPORT Distributed FTE Allocation
// Purpose: Read published roster FTE, measured in configured standard-FTE equivalents.
shared #"IMPORT Distributed FTE Allocation" =
    Table.Buffer(#"IMPORT Distributed FTE Workbook"{[Item="MinuteWorkersFTE_TABLE", Kind="Table"]}[Data]);

// Query: IMPORT Distributed FTE Profile
// Purpose: Read the complete fortnight grid, including explicitly validated zero cells.
shared #"IMPORT Distributed FTE Profile" =
    Table.Buffer(#"IMPORT Distributed FTE Workbook"{[Item="MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE", Kind="Table"]}[Data]);

// Query: IMPORT Distributed FTE Checks
// Purpose: Read saved upstream validation results; any error blocks demand publication.
shared #"IMPORT Distributed FTE Checks" =
    Table.Buffer(Table.SelectColumns(
        #"IMPORT Distributed FTE Workbook"{[Item="MinuteWorkersFTE_CHECK", Kind="Table"]}[Data],
        {"Severity", "Check", "Facility", "Role", "Message"}));

// Query: IMPORT Settings Data
// Purpose: Share the unit Settings workbook across standard duration, timing and calendar queries.
shared #"IMPORT Settings Data" =
    Table.Buffer(
        Excel.Workbook(Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\Settings Data.xlsx")), null, true)
    );

shared ShiftStart = let
    Source = #"IMPORT Settings Data",
    ShiftStart_Table = Source{[Item="ShiftStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftStart_Table,{{"StartAM", type number}, {"StartPM", type number}, {"StartNIGHT", type number}})
in
    #"Changed Type";

// Query: ShiftDuration
// Purpose: Read and validate the standard-FTE duration in hours from Settings Data.
// Inputs: IMPORT Settings Data, named table ShiftDuration, column ShiftDuration.
// Output: One positive finite duration in hours; missing or invalid settings stop publication.
shared ShiftDuration =
let
    Matches = Table.SelectRows(#"IMPORT Settings Data",
        each [Item] = "ShiftDuration" and [Kind] = "Table"),
    Data = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error "Settings Data must contain exactly one ShiftDuration table.",
    RequiredColumn = if Table.HasColumns(Data, {"ShiftDuration"}) then Data
        else error "The ShiftDuration table must contain the ShiftDuration column.",
    Hours = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn{0}[ShiftDuration]
        else error "ShiftDuration must contain exactly one data row.",
    ValidatedHours = if not Value.Is(Hours, type number) then
            error "ShiftDuration must be a numeric hours value."
        else if Number.IsNaN(Hours) or Hours <= 0 or Number.Abs(Hours) = #infinity then
            error "ShiftDuration must be a positive finite number of hours."
        else Hours
in
    ValidatedHours;

shared #"IMPORT ShiftPeriod" = let
    Source = #"IMPORT Settings Data",
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}, {"DurationOfShifts", type number}})
in
    #"Changed Type";

shared #"IMPORT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

// Query: Distributed FTE Source Validate
// Purpose: Validate the paired publication schema and upstream checks before preparing demand.
// Notes: Supports the known QFR publication and the revised explicit DC-role publication;
// both paired tables must use the same schema. This does not assert their refresh currency.
shared #"Distributed FTE Source Validate" =
let
    Allocation = #"IMPORT Distributed FTE Allocation",
    Profile = #"IMPORT Distributed FTE Profile",
    Common = {"Facility", "Role", "MinuteCategory", "Direct Care %", "Week No",
        "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift"},
    Schema = (T as table) as text =>
        if Table.HasColumns(T, {"DC Role", "DC Category"}) then "DC"
        else if Table.HasColumns(T, {"QFR Category"}) then "QFR"
        else error "Unrecognised distributed FTE publication schema.",
    RequiredColumnsPresent = Table.HasColumns(Allocation, Common & {"FTE"})
        and Table.HasColumns(Profile, Common & {"HistoricalRosterFTE", "RedistributedRosterFTE",
            "RedistributionMatchCount", "HistoricalCoverageStatus", "ProfileAlignmentStatus"}),
    Checks = #"IMPORT Distributed FTE Checks",
    InvalidChecks = Table.SelectRows(Checks, each not List.Contains({"Pass", "Warning"}, [Severity])),
    Result = if not RequiredColumnsPresent then error "Required allocation/profile columns are missing."
        else if Schema(Allocation) <> Schema(Profile) then error "Allocation and profile publication schemas differ."
        else if Table.IsEmpty(Checks) or not Table.IsEmpty(InvalidChecks) then
            error Error.Record("Upstream validation failed", "Inspect MinuteWorkersFTE_CHECK.", InvalidChecks)
        else Schema(Allocation)
in
    Result;

// Query: Distributed FTE Rows Prepare
// Purpose: Preserve source keys and audit attributes while assigning RN, AIN or AINC4.
// Notes: Strictly validate the saved numeric fields, avoiding text-to-number or integer rounding.
shared #"Distributed FTE Rows Prepare" = (Source as table, IsProfile as logical) as table =>
let
    Schema = #"Distributed FTE Source Validate",
    Measures = if IsProfile then {"HistoricalRosterFTE", "RedistributedRosterFTE",
        "RedistributionMatchCount", "HistoricalCoverageStatus", "ProfileAlignmentStatus"} else {"FTE"},
    Common = {"Facility", "Role", "MinuteCategory", "Direct Care %", "Week No",
        "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift"},
    Selected = Table.SelectColumns(Source, Common & Measures &
        (if Schema = "DC" then {"DC Role", "DC Category"} else {"QFR Category"})),
    Renamed = Table.RenameColumns(Selected, {{"Role", "SourceRole"}, {"Shift", "SourceShift"}}),
    WithRole = Table.AddColumn(Renamed, "Role", each #"Demand Role Code"([SourceRole]), type nullable text),
    WithShift = Table.AddColumn(WithRole, "Shift",
        each if [SourceShift] = "NS" then "NIGHT" else [SourceShift], type text),
    InvalidKeys = Table.SelectRows(WithShift, each
        not Value.Is([Facility], type text) or not Value.Is([SourceRole], type text)
        or not #"Demand Finite Number"([Week No])
        or not List.Contains({1, 2}, [FortnightWeek])
        or not List.Contains({1..7}, [DayOfWeek])
        or not List.Contains({1..14}, [FortnightDayIndex])
        or [FortnightDayIndex] <> ([FortnightWeek] - 1) * 7 + [DayOfWeek]
        or not List.Contains({"AM", "PM", "NS"}, [SourceShift])
        or not #"Demand Finite Number"([#"Direct Care %"])
        or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1
        or (Schema = "DC" and Text.Upper(Text.Trim([SourceRole])) <> Text.Upper(Text.Trim([DC Role])))),
    Keys = {"Facility", "SourceRole", "FortnightDayIndex", "Shift"},
    DuplicateKeys = Table.RowCount(WithShift) <> Table.RowCount(Table.Distinct(Table.SelectColumns(WithShift, Keys))),
    InvalidAllocation = if IsProfile then #table({}, {}) else
        Table.SelectRows(WithShift, each not #"Demand Finite Number"([FTE]) or [FTE] < 0),
    Result = if not Table.IsEmpty(InvalidKeys) then
            error Error.Record("Invalid source keys", "Invalid role, fortnight, shift or direct-care attributes.", InvalidKeys)
        else if DuplicateKeys then error "Duplicate facility/role/fortnight-day/shift keys in source publication."
        else if not Table.IsEmpty(InvalidAllocation) then error "Allocation FTE must be finite and non-negative."
        else Table.Buffer(WithShift)
in
    Result;

// Query: Distributed FTE Allocation Prepare
// Purpose: Prepare saved allocation rows at their original facility/role/day/shift grain.
shared #"Distributed FTE Allocation Prepare" =
    #"Distributed FTE Rows Prepare"(#"IMPORT Distributed FTE Allocation", false);

// Query: Distributed FTE Profile Prepare
// Purpose: Prepare complete coverage evidence at the same grain as the sparse allocation.
shared #"Distributed FTE Profile Prepare" =
    #"Distributed FTE Rows Prepare"(#"IMPORT Distributed FTE Profile", true);

// Query: Distributed FTE Exclusions
// Purpose: Report excluded allocations without redistributing their hours to RN or AIN.
// Output: Fortnight roster and productive hours by facility, original role and exclusion reason.
shared #"Distributed FTE Exclusions" =
let
    WithReason = Table.AddColumn(#"Distributed FTE Allocation Prepare", "ExclusionReason", each
        if [Facility] <> #"Demand Facility" then "Outside modeled facility"
        else if not List.Contains(#"Demand Roles", [Role]) then "Role temporarily excluded"
        else null, type nullable text),
    Excluded = Table.SelectRows(WithReason, each [ExclusionReason] <> null),
    Result = Table.Group(Excluded, {"Facility", "SourceRole", "ExclusionReason"}, {
        {"FortnightRosterHours", each List.Sum([FTE]) * ShiftDuration, type number},
        {"FortnightProductiveHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[FTE] * ShiftDuration * R[#"Direct Care %"])), type number}
    })
in
    Result;

// Query: Distributed FTE Prepare
// Purpose: Produce the complete validated 14-day RN/AIN/AINC4 demand pattern for the modeled facility.
// Output: 126 rows: three roles x fourteen distinct days x three shifts, including proven zeros.
// Notes: Join on original source role and week as well as fortnight day. A missing sparse row
// is zero only when the paired profile explicitly proves PASS ZERO with no redistribution match.
shared #"Distributed FTE Prepare" =
let
    Allocation = #"Distributed FTE Allocation Prepare",
    Profile = #"Distributed FTE Profile Prepare",
    ExpectedCells = List.Count(#"Demand Roles") * 14 * 3,
    InScope = (T as table) as table => Table.SelectRows(T,
        each [Facility] = #"Demand Facility" and List.Contains(#"Demand Roles", [Role])),
    IncludedAllocation = InScope(Allocation),
    IncludedProfile = InScope(Profile),
    Keys = {"Facility", "SourceRole", "Week No", "FortnightDayIndex", "Shift"},
    Orphans = Table.NestedJoin(IncludedAllocation, Keys, IncludedProfile, Keys, "Profile", JoinKind.LeftAnti),
    Joined = Table.NestedJoin(IncludedProfile, Keys, IncludedAllocation, Keys, "Allocation", JoinKind.LeftOuter),
    WithCount = Table.AddColumn(Joined, "AllocationCount", each Table.RowCount([Allocation]), Int64.Type),
    InvalidCells = Table.SelectRows(WithCount, each
        [HistoricalCoverageStatus] <> "PASS"
        or not List.Contains({"PASS", "PASS ZERO"}, [ProfileAlignmentStatus])
        or not #"Demand Finite Number"([RedistributedRosterFTE]) or [RedistributedRosterFTE] < 0
        or not #"Demand Finite Number"([HistoricalRosterFTE]) or [HistoricalRosterFTE] < 0
        or [RedistributionMatchCount] <> [AllocationCount]
        or (if [Role] = "RN" then [MinuteCategory] <> "RN" else [MinuteCategory] <> "OTHERS")
        or (if [AllocationCount] = 1 then
            Number.Abs([RedistributedRosterFTE] - [Allocation]{0}[FTE]) > 0.0000001
            or [#"Direct Care %"] <> [Allocation]{0}[#"Direct Care %"]
            or [MinuteCategory] <> [Allocation]{0}[MinuteCategory]
            else [AllocationCount] <> 0 or [ProfileAlignmentStatus] <> "PASS ZERO"
                or [HistoricalRosterFTE] <> 0 or [RedistributedRosterFTE] <> 0)
        or ([ProfileAlignmentStatus] = "PASS ZERO" and
            ([HistoricalRosterFTE] <> 0 or [RedistributedRosterFTE] <> 0))
        or ([ProfileAlignmentStatus] = "PASS" and [HistoricalRosterFTE] <= 0)),
    DestinationKeys = {"Role", "FortnightDayIndex", "Shift"},
    PatternWeeks = Table.Distinct(Table.SelectColumns(IncludedProfile, {"FortnightWeek", "Week No"})),
    ValidWeeks = Table.RowCount(PatternWeeks) = 2
        and List.Count(List.Distinct(PatternWeeks[Week No])) = 2
        and (if Table.RowCount(PatternWeeks) = 2 then
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){0}[Week No] <
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){1}[Week No] else false),
    Validated = if not Table.IsEmpty(Orphans) then error "Allocated source cells are missing from the paired profile."
        else if Table.RowCount(IncludedProfile) <> ExpectedCells or
            Table.RowCount(Table.Distinct(Table.SelectColumns(IncludedProfile, DestinationKeys))) <> ExpectedCells then
            error "RN/AIN/AINC4 require exactly 126 unique fortnight day/shift cells for the modeled facility."
        else if not ValidWeeks then error "The pattern must contain two distinct ordered historical weeks."
        else if not Table.IsEmpty(InvalidCells) then
            error Error.Record("Distributed FTE cells failed", "Allocation/profile mismatch or unproven zero.", Table.RemoveColumns(InvalidCells, {"Allocation"}))
        else Table.RemoveColumns(WithCount, {"Allocation", "AllocationCount"}),
    WithFTE = Table.AddColumn(Validated, "SourceFTE", each [RedistributedRosterFTE], type number),
    // Source FTE already represents roster time. Do not apply Direct Care % again.
    WithHours = Table.AddColumn(WithFTE, "DemandHRS", each [SourceFTE] * ShiftDuration, type number)
in
    Table.Buffer(WithHours);

// Query: Demand Calendar Prepare
// Purpose: Retain the settings period IDs while mapping 28 sequential dates onto two fortnight cycles.
// Notes: Retain RN/AIN/AINC4. No dates, periods or timing are invented for missing settings.
shared #"Demand Calendar Prepare" =
let
    SelectedRoles = Table.SelectRows(#"IMPORT PermutationDimensions",
        each List.Contains(#"Demand Roles", [RolesList])),
    Calendar = Table.RenameColumns(SelectedRoles, {{"RolesList", "Role"}, {"Shifts", "Shift"}}),
    ExpectedCells = List.Count(#"Demand Roles") * 28 * 3,
    Dates = List.Sort(List.Distinct(Calendar[Date])),
    DatesValid = List.Count(Dates) = 28 and not List.Contains(Dates, null),
    Start = if DatesValid then Dates{0} else error "Settings must provide exactly 28 non-null planning dates.",
    ExpectedDates = List.Dates(Start, 28, #duration(1, 0, 0, 0)),
    Periods = Table.Distinct(Table.SelectColumns(Calendar, {"Date", "Shift", "Period"})),
    InvalidRows = Table.SelectRows(Calendar, each [Day] <> Duration.Days([Date] - Start) + 1
        or not List.Contains({"AM", "PM", "NIGHT"}, [Shift]) or [Period] = null),
    Validated = if Dates <> ExpectedDates or Date.DayOfWeek(Start, Day.Monday) <> 0 then
            error "The 28-day planning calendar must be consecutive and start on the source Week-1 Monday."
        else if Table.RowCount(Calendar) <> ExpectedCells or
            Table.RowCount(Table.Distinct(Table.SelectColumns(Calendar, {"Date", "Role", "Shift"}))) <> ExpectedCells then
            error "Expected exactly 252 calendar cells: 28 days x RN/AIN/AINC4 x three shifts."
        else if Table.RowCount(Periods) <> 84 or List.Sort(List.Distinct(Periods[Period])) <> {1..84} then
            error "Settings must provide 84 unique shift periods shared by RN, AIN and AINC4."
        else if not Table.IsEmpty(InvalidRows) then error "Invalid sequential Day, Shift or Period in settings."
        else Calendar,
    // Day remains the original sequential 1..28 index. This separate key repeats 1..14 twice.
    WithFortnightDay = Table.AddColumn(Validated, "FortnightDayIndex",
        each Number.Mod(Duration.Days([Date] - Start), 14) + 1, Int64.Type),
    WithCycle = Table.AddColumn(WithFortnightDay, "PlanningFortnight",
        each Number.IntegerDivide(Duration.Days([Date] - Start), 14) + 1, Int64.Type)
in
    Table.Buffer(WithCycle);

// Query: Permutation DateTimeRoleShift
// Purpose: Attach one validated settings shift span to each RN/AIN/AINC4 calendar cell.
// Notes: Preserve existing overnight handling and require DurationOfShifts to equal the timestamp span.
shared #"Permutation DateTimeRoleShift" =
let
    Periods = Table.SelectRows(#"IMPORT ShiftPeriod", each List.Contains(#"Demand Roles", [Role])),
    Joined = Table.NestedJoin(#"Demand Calendar Prepare", {"Role", "Shift"}, Periods,
        {"Role", "ShiftPeriod"}, "Timing", JoinKind.LeftOuter),
    InvalidMatches = Table.SelectRows(Joined, each Table.RowCount([Timing]) <> 1),
    ValidatedMatches = if not Table.IsEmpty(InvalidMatches) then
        error "Each RN/AIN/AINC4 calendar cell requires exactly one settings shift definition." else Joined,
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
        else Table.RemoveColumns(WithEnd, {"Start", "End"})
in
    Table.Buffer(Result);

// Query: Demand Extraction Prepare
// Purpose: Repeat each source cell once per fortnight and calculate numeric shift-average attendance.
// Output: One RN/AIN/AINC4 demand row per period, with source attributes retained for reconciliation.
shared #"Demand Extraction Prepare" =
let
    Joined = Table.NestedJoin(#"Permutation DateTimeRoleShift", {"Role", "FortnightDayIndex", "Shift"},
        #"Distributed FTE Prepare", {"Role", "FortnightDayIndex", "Shift"}, "Pattern", JoinKind.LeftOuter),
    InvalidMatches = Table.SelectRows(Joined, each Table.RowCount([Pattern]) <> 1),
    Validated = if not Table.IsEmpty(InvalidMatches) then
        error "Every calendar cell must match exactly one validated fortnight allocation." else Joined,
    Expanded = Table.ExpandTableColumn(Validated, "Pattern",
        {"Facility", "SourceRole", "SourceFTE", "DemandHRS", "Direct Care %", "Week No"},
        {"Facility", "SourceRole", "SourceFTE", "DemandHRS", "Direct Care %", "Week No"}),
    WithUnit = Table.AddColumn(Expanded, "Unit", each [Facility], type text),
    // A source standard-FTE equivalent is converted to attendance across the actual shift span.
    // Leave fractional attendance unrounded so interval integration recovers the roster hours.
    WithAttendance = Table.AddColumn(WithUnit, "DemandFTE", each [DemandHRS] / [DurationOfShifts], type number),
    InvalidValues = Table.SelectRows(WithAttendance, each not #"Demand Finite Number"([DemandFTE])
        or [DemandFTE] < 0 or not #"Demand Finite Number"([DemandHRS])
        or Number.Abs([DemandFTE] * [DurationOfShifts] - [DemandHRS]) > 0.0000001)
in
    if not Table.IsEmpty(InvalidValues) then error "Demand hours/attendance conversion failed."
    else Table.Buffer(WithAttendance);

// Query: DemandExtraction_RECONCILIATION
// Purpose: Reconcile each retained role in each fortnight to one complete published source pattern.
// Notes: The 28-day total is twice the fortnight target. Excluded roles do not inflate retained demand.
shared DemandExtraction_RECONCILIATION =
let
    Pattern = Table.Group(#"Distributed FTE Prepare", {"Facility", "Role"}, {
        {"ExpectedRosterHours", each List.Sum([DemandHRS]), type number},
        {"ExpectedProductiveHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[DemandHRS] * R[#"Direct Care %"])), type number}
    }),
    Actual = Table.Group(#"Demand Extraction Prepare", {"Facility", "Role", "PlanningFortnight"}, {
        {"CellCount", each Table.RowCount(_), Int64.Type},
        {"RosterHours", each List.Sum([DemandHRS]), type number},
        {"IntegratedHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[DemandFTE] * R[DurationOfShifts])), type number},
        {"ProductiveHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[DemandHRS] * R[#"Direct Care %"])), type number}
    }),
    Joined = Table.NestedJoin(Actual, {"Facility", "Role"}, Pattern, {"Facility", "Role"}, "Expected", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Expected", {"ExpectedRosterHours", "ExpectedProductiveHours"}),
    WithResidual = Table.AddColumn(Expanded, "RosterHoursResidual", each [RosterHours] - [ExpectedRosterHours], type number),
    Result = Table.AddColumn(WithResidual, "Status", each
        if [CellCount] = 42 and Number.Abs([RosterHoursResidual]) <= 0.0000001
            and Number.Abs([IntegratedHours] - [ExpectedRosterHours]) <= 0.0000001
            and Number.Abs([ProductiveHours] - [ExpectedProductiveHours]) <= 0.0000001
        then "Pass" else "Error", type text)
in
    Table.Buffer(Result);

// Query: DemandExtraction_CHECK
// Purpose: Expose stage failures as readable diagnostics and gate the published demand interface.
// Notes: Each try captures an error for reporting only; failed stages never become empty demand.
shared DemandExtraction_CHECK =
let
    CheckStage = (Name as text, Work as function) as record =>
        let Attempt = try Work()
        in [Severity = if Attempt[HasError] then "Error" else "Pass", Check = Name,
            Message = if Attempt[HasError] then Attempt[Error][Message] else "Validated",
            Actual = if Attempt[HasError] then null else Attempt[Value]],
    StageChecks = Table.FromRecords({
        CheckStage("Source schema and saved publication checks", () => #"Distributed FTE Source Validate"),
        CheckStage("Complete RN/AIN/AINC4 fortnight allocation and profile agreement", () => Table.RowCount(#"Distributed FTE Prepare")),
        CheckStage("28 dates and 84 shared shift periods", () => Table.RowCount(#"Demand Calendar Prepare")),
        CheckStage("Unique timing and exact shift spans", () => Table.RowCount(#"Permutation DateTimeRoleShift")),
        CheckStage("Complete demand and fractional attendance conversion", () => Table.RowCount(#"Demand Extraction Prepare")),
        CheckStage("Retained hours in both fortnights", () =>
            let Reconciliation = DemandExtraction_RECONCILIATION
            in if Table.RowCount(Reconciliation) <> List.Count(#"Demand Roles") * 2 or
                not Table.IsEmpty(Table.SelectRows(Reconciliation, each [Status] <> "Pass")) then
                error "Fortnight roster/productive/integrated hours do not reconcile."
            else Table.RowCount(Reconciliation))
    }),
    UpstreamAttempt = try Table.Buffer(#"IMPORT Distributed FTE Checks"),
    Warnings = if UpstreamAttempt[HasError] then #table({"Severity", "Check", "Message", "Actual"}, {}) else
        Table.AddColumn(Table.SelectColumns(
            Table.SelectRows(UpstreamAttempt[Value], each [Severity] = "Warning"),
            {"Severity", "Check", "Message"}), "Actual", each null)
in
    Table.Combine({StageChecks, Warnings});

// Query: ShiftUnitDemandHRS
// Purpose: Publish validated RN/AIN/AINC4 demand through the unchanged twelve-column downstream interface.
// Output: 252 rows over 28 days, preserving Date/Day/Period and shift timestamps.
// Notes: Retain this query/table/worksheet name and headers at row 1 when explicitly synchronized.
shared ShiftUnitDemandHRS =
let
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
    Result;
