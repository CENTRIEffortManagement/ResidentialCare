// Power Query from: 2-DemandExtract.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\1. Input\2-DemandExtract.xlsx
// Extracted: 2026-09-21T06:32:36.799Z

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
            error "Root Path did not end in an expected unit workbook folder: " & WorkbookFolder
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

// Query: Facility Source Code Mapping
// Purpose: Match the upstream two-character Facility code to the List's canonical abbreviation.
// Notes: Demand-MasterRoster derives its source Facility code from the first two Location characters.
shared #"Facility Source Code Mapping" =
let
    Source = #"LINK Facilities",
    WithCode = Table.AddColumn(Source, "SourceFacilityCode", each
        if [Title] = null then null else Text.Upper(Text.Start(Text.Trim([Title]), 2)), type nullable text),
    Mappings = Table.Distinct(Table.SelectColumns(
        Table.SelectRows(WithCode, each [SourceFacilityCode] <> null and [SourceFacilityCode] <> ""),
        {"SourceFacilityCode", "Facility-Abbrev"})),
    Counts = Table.Group(Mappings, {"SourceFacilityCode"}, {{"MappingCount", each Table.RowCount(_), Int64.Type}}),
    Ambiguous = Table.SelectRows(Counts, each [MappingCount] <> 1),
    Result = if Table.IsEmpty(Ambiguous) then Mappings
        else error Error.Record("Ambiguous facility code", "A source Facility code maps to more than one Facility-Abbrev.", Ambiguous)
in
    Result;

// Query: FacilityAnalysisTABLE
// Purpose: Identify whether the folder-derived Unit is enabled for Effort Management Analysis.
shared FacilityAnalysisTABLE =
let
    Source = #"LINK Facility Analysis",
    UnitRows = Table.SelectRows(Source, each [Title] <> null and Comparer.OrdinalIgnoreCase(Text.Trim([Title]), #"Demand Facility") = 0),
    ValidatedUnit = if Table.RowCount(UnitRows) = 1 then UnitRows else error "Facility Analysis must identify exactly one row for the current Unit.",
    EnabledUnit = Table.SelectRows(ValidatedUnit, each [EffortManagementAnalysis] = true),
    Result = Table.RemoveColumns(EnabledUnit, {"LeaveBalanceAnalysis"})
in
    Result;

shared ShiftOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t, Column2 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "Shift"}, {"Column2", "Order"}})
in
    #"Renamed Columns";

// Query: Demand Finite Number
// Purpose: Reject nulls, text, NaN and infinities before demand arithmetic.
shared #"Demand Finite Number" = (Value as any) as logical =>
    if not Value.Is(Value, type number) then false
    else not Number.IsNaN(Value) and Number.Abs(Value) <> #infinity;

// Query: IMPORT Distributed FTE Workbook
// Purpose: Share one saved-workbook snapshot across allocation, profile and validation imports.
// Notes: Buffer the binary as well as the navigation table; do not refresh or modify the source.
shared #"IMPORT Distributed FTE Workbook" =
let
    SourcePath = Unit1Path & "\1. Input\Demand-MasterRoster Manual Read.xlsx",
    SourceBinary = Binary.Buffer(File.Contents(SourcePath)),
    WorkbookNavigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(WorkbookNavigation)
in
    BufferedNavigation;

// Query: IMPORT Distributed FTE Allocation
// Purpose: Read published roster FTE, measured in configured standard-FTE equivalents.
shared #"IMPORT Distributed FTE Allocation" =
let
    WorkbookNavigation = #"IMPORT Distributed FTE Workbook",
    AllocationTable = WorkbookNavigation{[Item="MinuteWorkersFTE_TABLE", Kind="Table"]}[Data],
    BufferedAllocation = Table.Buffer(AllocationTable)
in
    BufferedAllocation;

// Query: IMPORT Distributed FTE Profile
// Purpose: Read the complete fortnight grid, including explicitly validated zero cells.
shared #"IMPORT Distributed FTE Profile" =
let
    WorkbookNavigation = #"IMPORT Distributed FTE Workbook",
    ProfileTable = WorkbookNavigation{[Item="MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE", Kind="Table"]}[Data],
    BufferedProfile = Table.Buffer(ProfileTable)
in
    BufferedProfile;

// Query: IMPORT Distributed FTE Checks
// Purpose: Read saved upstream validation results; any error blocks demand publication.
shared #"IMPORT Distributed FTE Checks" =
let
    WorkbookNavigation = #"IMPORT Distributed FTE Workbook",
    SourceChecks = WorkbookNavigation{[Item="MinuteWorkersFTE_CHECK", Kind="Table"]}[Data],
    CheckColumns = Table.SelectColumns(SourceChecks, {"Severity", "Check", "Facility", "Role", "Message"}),
    BufferedChecks = Table.Buffer(CheckColumns)
in
    BufferedChecks;

// Query: IMPORT Settings Data
// Purpose: Share the unit Settings workbook across standard duration, timing and calendar queries.
shared #"IMPORT Settings Data" =
let
    SourcePath = Unit1Path & "\2. Calculations\Settings Data.xlsx",
    SourceBinary = Binary.Buffer(File.Contents(SourcePath)),
    WorkbookNavigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(WorkbookNavigation)
in
    BufferedNavigation;

// Query: LINK Roles
// Purpose: Extract the roster/DC-role definitions and the authoritative output Role Group.
shared #"LINK Roles" =
let
    ListNavigation = #"IMPORT Role Lists",
    RoleDefinitions = ListNavigation{[Id="6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items],
    RoleColumns = Table.SelectColumns(RoleDefinitions,
        {"Roster Roles", "DC Category", "DC Role", "Direct Care %", "Role Group"}),
    BufferedRoles = Table.Buffer(RoleColumns)
in
    BufferedRoles;

// Query: LINK RoleAnalysis
// Purpose: Extract role-group analysis controls; only the effort flag governs demand inclusion.
shared #"LINK RoleAnalysis" =
let
    ListNavigation = #"IMPORT Role Lists",
    AnalysisControls = ListNavigation{[Id="bad3beb8-7064-4e11-9edd-d62dac5d702d"]}[Items],
    AnalysisColumns = Table.SelectColumns(AnalysisControls,
        {"Leave Balance Analysis", "Effort Management Analysis", "Roles"}),
    BufferedAnalysis = Table.Buffer(AnalysisColumns)
in
    BufferedAnalysis;

// Query: LINK Facilities
// Purpose: Map source facility titles to their canonical Facility-Abbrev.
shared #"LINK Facilities" =
let
    FacilityItems = #"IMPORT Whiddon Facility Lists"{[Id="1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items],
    SelectedColumns = Table.SelectColumns(FacilityItems, {"Title", "field_1"}),
    NamedColumns = Table.RenameColumns(SelectedColumns, {{"field_1", "Facility-Abbrev"}})
in
    NamedColumns;

// Query: LINK Facility Analysis
// Purpose: Select the same Facility Analysis fields used by AllocationExtracted.
shared #"LINK Facility Analysis" =
let
    AnalysisItems = #"IMPORT Whiddon Facility Lists"{[Id="17ed8e30-5707-4af2-8e21-d68eb8184287"]}[Items],
    SelectedColumns = Table.SelectColumns(AnalysisItems, {"Title", "EffortManagementAnalysis", "LeaveBalanceAnalysis"})
in
    SelectedColumns;

// Query: IMPORT Role Lists
// Purpose: Share the existing Whiddon SharePoint list navigation between the two role links.
// Notes: Retain the approved site and list IDs used by AllocationExtraction.
shared #"IMPORT Role Lists" =
let
    ListNavigation = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI",
        [Implementation="2.0", ViewMode="All"]),
    BufferedLists = Table.Buffer(ListNavigation)
in
    BufferedLists;

// Query: Demand Role Key
// Purpose: Match published DC roles without case or surrounding-space differences.
// Notes: Output Role Group labels retain their configured spelling and case.
shared #"Demand Role Key" = (RoleName as nullable text) as nullable text =>
    if RoleName = null then null else Text.Upper(Text.Trim(RoleName));

// Query: Demand Role Analysis Prepare
// Purpose: Validate one effort-analysis setting per role-group label before any joins.
// Notes: Logical false and null are not enabled. Text or numeric flags are invalid.
shared #"Demand Role Analysis Prepare" =
let
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
    BufferedAnalysis;

// Query: Demand Role Mapping
// Purpose: Apply effort controls before validating the output assignment of each source DC role.
// Output: One row per normalised DC role; wholly excluded roles have no output group when their groups differ.
// Notes: Demand already contains DC-role aggregates, not individual Roster Roles.
// Notes: Preview Conflicts for ambiguous enabled mappings, or CollapsedExcluded for wholly excluded DC roles.
shared #"Demand Role Mapping" =
let
    Source = #"LINK Roles",
    Selected = Table.SelectColumns(Source, {"Roster Roles", "DC Role", "Role Group", "DC Category", "Direct Care %"}),
    Normalised = Table.TransformColumns(Selected, {
        {"DC Role", each #"Demand Role Key"(_), type nullable text},
        {"Role Group", each if _ = null then null else Text.Trim(_), type nullable text},
        {"DC Category", each #"Demand Role Key"(_), type nullable text}
    }),
    // Upstream publications contain RN/OTHER direct-care roles only. Non-care definitions
    // may have no DC role; they do not create demand or require a fabricated mapping.
    CareDefinitions = Table.SelectRows(Normalised, each List.Contains({"RN", "OTHER"}, [DC Category])),
    InvalidRows = Table.SelectRows(CareDefinitions, each [DC Role] = null or [DC Role] = ""
        or [Role Group] = null or [Role Group] = ""
        or not #"Demand Finite Number"([#"Direct Care %"])
        or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1),
    ValidatedRows = if not Table.IsEmpty(InvalidRows) then
            error Error.Record("Invalid demand role mapping", "Direct-care roles need a DC Role, Role Group and valid Direct Care %. Inspect Demand Role Mapping > InvalidRows.", InvalidRows)
        else CareDefinitions,
    // Determine inclusion for every contributing group before deciding whether a shared
    // DC role needs an output assignment. A missing analysis row is not proof of exclusion.
    Joined = Table.NestedJoin(ValidatedRows, {"Role Group"}, #"Demand Role Analysis Prepare", {"Roles"}, "Analysis", JoinKind.LeftOuter),
    MissingAnalysis = Table.SelectRows(Joined, each Table.RowCount([Analysis]) <> 1),
    ValidatedAnalysis = if not Table.IsEmpty(MissingAnalysis) then
            error Error.Record("Missing role analysis", "Every direct-care Role Group needs a LINK RoleAnalysis entry.",
                Table.SelectColumns(MissingAnalysis, {"DC Role", "Role Group"}))
        else Joined,
    Expanded = Table.ExpandTableColumn(ValidatedAnalysis, "Analysis", {"Effort Management Analysis"}),
    // Several roster labels can define the same DC role. Collapse identical attributes
    // before joining facts. If any group is enabled, retain all definitions in the conflict
    // check: filtering out its disabled siblings would assign their hours to an enabled group.
    DefinitionAttributes = Table.SelectColumns(Expanded, {"DC Role", "Role Group", "DC Category", "Direct Care %", "Effort Management Analysis"}),
    DistinctDefinitions = Table.Distinct(DefinitionAttributes),
    Counts = Table.Group(DistinctDefinitions, {"DC Role"}, {
        {"DefinitionCount", each Table.RowCount(_), Int64.Type},
        {"HasEnabledGroup", each List.Contains([Effort Management Analysis], true), type logical},
        {"ConflictingFields", each Text.Combine(List.Select({"Role Group", "DC Category", "Direct Care %", "Effort Management Analysis"},
            (Column) => List.Count(List.Distinct(Table.Column(_, Column))) > 1), ", "), type text}
    }),
    ConflictingRoleKeys = Table.SelectRows(Counts, each [HasEnabledGroup] and [DefinitionCount] <> 1),
    // Keep the original roster labels beside each conflicting definition. Expand the
    // counts to scalar columns so the Applied Step is readable without opening nested tables.
    ConflictSourceRows = Table.NestedJoin(Expanded, {"DC Role"}, ConflictingRoleKeys,
        {"DC Role"}, "ConflictDetails", JoinKind.Inner),
    ExpandedConflicts = Table.ExpandTableColumn(ConflictSourceRows, "ConflictDetails", {"DefinitionCount", "ConflictingFields"}),
    Conflicts = Table.Sort(ExpandedConflicts, {{"DC Role", Order.Ascending}, {"Roster Roles", Order.Ascending}}),
    Validated = if not Table.IsEmpty(Conflicts) then
            error Error.Record("Ambiguous DC role mapping", "A DC role with any enabled group must have one Role Group, category, percentage and inclusion setting. Its combined hours cannot be split between groups. Inspect Demand Role Mapping > Conflicts. Affected DC roles: "
                & Text.Combine(List.Transform(ConflictingRoleKeys[DC Role], each if _ = null then "<blank>" else _), ", "), Conflicts)
        else DistinctDefinitions,
    EnabledDefinitions = Table.SelectRows(Validated, each [Effort Management Analysis] = true),
    ExcludedDefinitions = Table.SelectRows(Validated, each [Effort Management Analysis] <> true),
    CommonValue = (Rows as table, Column as text) as any =>
        let Values = List.Distinct(Table.Column(Rows, Column))
        in if List.Count(Values) = 1 then Values{0} else null,
    // Entirely excluded DC roles need one match for the source/exclusions audit, not a
    // fabricated output group. Preserve common attributes only; source rows own their hours.
    CollapsedExcluded = Table.Group(ExcludedDefinitions, {"DC Role"}, {
        {"Role Group", each CommonValue(_, "Role Group"), type nullable text},
        {"DC Category", each CommonValue(_, "DC Category"), type nullable text},
        {"Direct Care %", each CommonValue(_, "Direct Care %"), type nullable number},
        {"Effort Management Analysis", each false, type logical}
    }),
    ResolvedDefinitions = Table.Combine({EnabledDefinitions, CollapsedExcluded}),
    NamedAttributes = Table.RenameColumns(ResolvedDefinitions, {{"DC Role", "SourceRoleKey"}, {"Role Group", "Role"},
        {"DC Category", "MappedDC Category"}, {"Direct Care %", "MappedDirectCare"}}),
    WithCategory = Table.AddColumn(NamedAttributes, "MappedMinuteCategory",
        each if [MappedDC Category] = null then null else if [MappedDC Category] = "RN" then "RN" else "OTHERS", type nullable text),
    BufferedRoleMapping = Table.Buffer(WithCategory)
in
    BufferedRoleMapping;

// Query: Demand Roles
// Purpose: Supply the distinct role groups enabled for effort management across all demand stages.
shared #"Demand Roles" =
let
    Enabled = Table.SelectRows(#"Demand Role Mapping", each [Effort Management Analysis] = true),
    Groups = List.Sort(List.Distinct(Enabled[Role])),
    ValidatedGroups = if List.IsEmpty(Groups) then error "No direct-care role groups are enabled for Effort Management Analysis."
        else Groups,
    BufferedGroups = List.Buffer(ValidatedGroups)
in
    BufferedGroups;

// Query: Demand Role Code
// Purpose: Preserve the role-code function interface using the linked DC-role mapping.
// Notes: No role-name guesses or forced enrolled-nurse aliases are applied.
// Notes: Returns null when a wholly excluded DC role has several distinct output groups.
shared #"Demand Role Code" = (RoleName as nullable text) as nullable text =>
let
    Key = #"Demand Role Key"(RoleName),
    Matches = Table.SelectRows(#"Demand Role Mapping", each [SourceRoleKey] = Key),
    MappedRole = if Table.RowCount(Matches) = 1 then Matches{0}[Role]
        else error "Published DC role has no linked Role Group: " & (if RoleName = null then "<null>" else RoleName)
in
    MappedRole;

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

// Query: IMPORT CentriSyncPaths
// Purpose: Read the public-machine path mapping used by the single workbook-path resolver.
// Notes: This bootstrap is the deliberate fixed-path exception.
shared #"IMPORT CentriSyncPaths" =
let
    PathMappingBinary = Binary.Buffer(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx")),
    WorkbookNavigation = Excel.Workbook(PathMappingBinary, null, true),
    BufferedNavigation = Table.Buffer(WorkbookNavigation)
in
    BufferedNavigation;

// Query: IMPORT Whiddon Facility Lists
// Purpose: Read the Whiddon SharePoint navigation for facility and analysis references.
shared #"IMPORT Whiddon Facility Lists" =
    SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation=null, ApiVersion=15]);

// Query: Demand Facility
// Purpose: Select the facility matching the unit folder identified by FilePathUrl.
// Notes: The upstream Master Roster extraction applies the Facilities List lookup to source locations.
shared #"Demand Facility" =
let
    UnitRows = Table.SelectRows(UnitL1PathTABLE, each [Variable Name] = "Unit"),
    UnitName = if Table.RowCount(UnitRows) = 1 then UnitRows{0}[Value] else error "UnitL1PathTABLE must identify one unit folder."
in
    if UnitName = null or Text.Trim(UnitName) = "" then error "The workbook path did not identify a unit folder." else UnitName;

// Query: Distributed FTE Rows Prepare
// Purpose: Preserve source DC-role keys and attach linked output groups and effort-analysis controls.
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
    JoinedFacilities = Table.NestedJoin(Selected, {"Facility"}, #"Facility Source Code Mapping",
        {"SourceFacilityCode"}, "FacilityLookup", JoinKind.LeftOuter),
    WithAbbrev = Table.ExpandTableColumn(JoinedFacilities, "FacilityLookup",
        {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    UnmappedFacilities = Table.SelectRows(WithAbbrev, each [#"Facility-Abbrev"] = null),
    ValidatedFacilities = if Table.IsEmpty(UnmappedFacilities) then WithAbbrev
        else error Error.Record("Missing facility mapping", "Every distributed FTE Facility code must map through LINK Facilities.",
            Table.Distinct(Table.SelectColumns(UnmappedFacilities, {"Facility"}))),
    Renamed = Table.RenameColumns(ValidatedFacilities, {{"Role", "SourceRole"}, {"Shift", "SourceShift"}}),
    WithRoleKey = Table.AddColumn(Renamed, "SourceRoleKey", each #"Demand Role Key"([SourceRole]), type nullable text),
    JoinedRoles = Table.NestedJoin(WithRoleKey, {"SourceRoleKey"}, #"Demand Role Mapping", {"SourceRoleKey"}, "RoleMapping", JoinKind.LeftOuter),
    MissingMappings = Table.SelectRows(JoinedRoles, each [#"Facility-Abbrev"] = #"Demand Facility" and Table.RowCount([RoleMapping]) <> 1),
    ValidatedMappings = if not Table.IsEmpty(MissingMappings) then
            error Error.Record("Missing demand role mapping", "Every published DC role for the modeled facility needs one linked definition.",
                Table.Distinct(Table.SelectColumns(MissingMappings, {"Facility", "SourceRole"})))
        else JoinedRoles,
    WithRole = Table.ExpandTableColumn(ValidatedMappings, "RoleMapping",
        {"Role", "Effort Management Analysis", "MappedDC Category", "MappedDirectCare", "MappedMinuteCategory"}),
    WithShift = Table.AddColumn(WithRole, "Shift",
        each if [SourceShift] = "NS" then "NIGHT" else [SourceShift], type text),
    InvalidKeys = Table.SelectRows(WithShift, each
        not Value.Is([Facility], type text) or [SourceRoleKey] = null or [SourceRoleKey] = ""
        or not #"Demand Finite Number"([Week No])
        or not List.Contains({1, 2}, [FortnightWeek])
        or not List.Contains({1..7}, [DayOfWeek])
        or not List.Contains({1..14}, [FortnightDayIndex])
        or [FortnightDayIndex] <> ([FortnightWeek] - 1) * 7 + [DayOfWeek]
        or not List.Contains({"AM", "PM", "NS"}, [SourceShift])
        or not #"Demand Finite Number"([#"Direct Care %"])
        or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1
        or (Schema = "DC" and [SourceRoleKey] <> #"Demand Role Key"([DC Role]))),
    InvalidRoleAttributes = Table.SelectRows(WithShift, each [#"Facility-Abbrev"] = #"Demand Facility"
        and [Effort Management Analysis] = true and (
            [MinuteCategory] <> [MappedMinuteCategory] or [#"Direct Care %"] <> [MappedDirectCare]
            or (Schema = "DC" and #"Demand Role Key"([DC Category]) <> [MappedDC Category]))),
    Keys = {"Facility", "SourceRoleKey", "FortnightDayIndex", "Shift"},
    DuplicateKeys = Table.RowCount(WithShift) <> Table.RowCount(Table.Distinct(Table.SelectColumns(WithShift, Keys))),
    InvalidAllocation = if IsProfile then #table({}, {}) else
        Table.SelectRows(WithShift, each not #"Demand Finite Number"([FTE]) or [FTE] < 0),
    Result = if not Table.IsEmpty(InvalidKeys) then
            error Error.Record("Invalid source keys", "Invalid role, fortnight, shift or direct-care attributes.", InvalidKeys)
        else if DuplicateKeys then error "Duplicate facility/role/fortnight-day/shift keys in source publication."
        else if not Table.IsEmpty(InvalidRoleAttributes) then
            error Error.Record("Published role attributes differ", "The retained source category and Direct Care % must agree with LINK Roles.", InvalidRoleAttributes)
        else if not Table.IsEmpty(InvalidAllocation) then error "Allocation FTE must be finite and non-negative."
        else Table.Buffer(WithShift)
in
    Result;

// Query: Distributed FTE Allocation Prepare
// Purpose: Prepare saved allocation rows at their original facility/role/day/shift grain.
shared #"Distributed FTE Allocation Prepare" =
let
    SourceAllocation = #"IMPORT Distributed FTE Allocation",
    PreparedAllocation = #"Distributed FTE Rows Prepare"(SourceAllocation, false)
in
    PreparedAllocation;

// Query: Distributed FTE Profile Prepare
// Purpose: Prepare complete coverage evidence at the same grain as the sparse allocation.
shared #"Distributed FTE Profile Prepare" =
let
    SourceProfile = #"IMPORT Distributed FTE Profile",
    PreparedProfile = #"Distributed FTE Rows Prepare"(SourceProfile, true)
in
    PreparedProfile;

// Query: Distributed FTE Exclusions
// Purpose: Report excluded allocations without redistributing their hours to retained role groups.
// Output: Fortnight roster and productive hours by facility, original role and exclusion reason.
shared #"Distributed FTE Exclusions" =
let
    WithReason = Table.AddColumn(#"Distributed FTE Allocation Prepare", "ExclusionReason", each
        if [#"Facility-Abbrev"] <> #"Demand Facility" then "Outside modeled facility"
        else if [Effort Management Analysis] <> true then "Effort Management Analysis is not enabled"
        else null, type nullable text),
    Excluded = Table.SelectRows(WithReason, each [ExclusionReason] <> null),
    Result = Table.Group(Excluded, {"Facility", "SourceRole", "ExclusionReason"}, {
        {"FortnightRosterHours", each List.Sum([FTE]) * ShiftDuration, type number},
        {"FortnightProductiveHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[FTE] * ShiftDuration * R[#"Direct Care %"])), type number}
    })
in
    Result;

// Query: Distributed FTE Source Cells
// Purpose: Validate every retained DC-role/day/shift cell before combining output role groups.
// Output: Forty-two cells per retained source DC role, including proven zeros.
// Notes: Join on original source role and week as well as fortnight day. A missing sparse row
// is zero only when the paired profile explicitly proves PASS ZERO with no redistribution match.
shared #"Distributed FTE Source Cells" =
let
    Allocation = #"Distributed FTE Allocation Prepare",
    Profile = #"Distributed FTE Profile Prepare",
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    InScope = (T as table) as table => Table.SelectRows(T,
        each [#"Facility-Abbrev"] = #"Demand Facility" and AnalysisEnabled and [Effort Management Analysis] = true),
    IncludedAllocation = InScope(Allocation),
    IncludedProfile = InScope(Profile),
    Keys = {"Facility", "SourceRoleKey", "Week No", "FortnightDayIndex", "Shift"},
    Orphans = Table.NestedJoin(IncludedAllocation, Keys, IncludedProfile, Keys, "Profile", JoinKind.LeftAnti),
    Joined = Table.NestedJoin(IncludedProfile, Keys, IncludedAllocation, Keys, "Allocation", JoinKind.LeftOuter),
    WithCount = Table.AddColumn(Joined, "AllocationCount", each Table.RowCount([Allocation]), Int64.Type),
    InvalidCells = Table.SelectRows(WithCount, each
        [HistoricalCoverageStatus] <> "PASS"
        or not List.Contains({"PASS", "PASS ZERO"}, [ProfileAlignmentStatus])
        or not #"Demand Finite Number"([RedistributedRosterFTE]) or [RedistributedRosterFTE] < 0
        or not #"Demand Finite Number"([HistoricalRosterFTE]) or [HistoricalRosterFTE] < 0
        or [RedistributionMatchCount] <> [AllocationCount]
        or [MinuteCategory] <> [MappedMinuteCategory]
        or (if [AllocationCount] = 1 then
            Number.Abs([RedistributedRosterFTE] - [Allocation]{0}[FTE]) > 0.0000001
            or [#"Direct Care %"] <> [Allocation]{0}[#"Direct Care %"]
            or [MinuteCategory] <> [Allocation]{0}[MinuteCategory]
            else [AllocationCount] <> 0 or [ProfileAlignmentStatus] <> "PASS ZERO"
                or [HistoricalRosterFTE] <> 0 or [RedistributedRosterFTE] <> 0)
        or ([ProfileAlignmentStatus] = "PASS ZERO" and
            ([HistoricalRosterFTE] <> 0 or [RedistributedRosterFTE] <> 0))
        or ([ProfileAlignmentStatus] = "PASS" and [HistoricalRosterFTE] <= 0)),
    // Validate each contributing source role separately: a complete group must not hide
    // a missing day/shift for one of its contributing DC roles.
    Coverage = Table.Group(IncludedProfile, {"Facility", "SourceRoleKey"}, {{"CellCount", each Table.RowCount(_), Int64.Type}}),
    InvalidCoverage = Table.SelectRows(Coverage, each [CellCount] <> 42),
    PatternWeeks = Table.Distinct(Table.SelectColumns(IncludedProfile, {"FortnightWeek", "Week No"})),
    ValidWeeks = Table.RowCount(PatternWeeks) = 2
        and List.Count(List.Distinct(PatternWeeks[Week No])) = 2
        and (if Table.RowCount(PatternWeeks) = 2 then
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){0}[Week No] <
            Table.Sort(PatternWeeks, {{"FortnightWeek", Order.Ascending}}){1}[Week No] else false),
    Validated = if not Table.IsEmpty(Orphans) then error "Allocated source cells are missing from the paired profile."
        else if Table.IsEmpty(Coverage) or not Table.IsEmpty(InvalidCoverage) then
            error Error.Record("Incomplete DC-role coverage", "Every retained source DC role requires 42 unique fortnight day/shift cells.", InvalidCoverage)
        else if not ValidWeeks then error "The pattern must contain two distinct ordered historical weeks."
        else if not Table.IsEmpty(InvalidCells) then
            error Error.Record("Distributed FTE cells failed", "Allocation/profile mismatch or unproven zero.", Table.RemoveColumns(InvalidCells, {"Allocation"}))
        else Table.RemoveColumns(WithCount, {"Allocation", "AllocationCount"}),
    WithFTE = Table.AddColumn(Validated, "SourceFTE", each [RedistributedRosterFTE], type number),
    // Source FTE already represents roster time. Do not apply Direct Care % again.
    WithHours = Table.AddColumn(WithFTE, "DemandHRS", each [SourceFTE] * ShiftDuration, type number),
    // Productive hours are an audit measure. Apply the percentage per source role before
    // grouping; percentages from different roles must never be added or averaged.
    WithProductiveHours = Table.AddColumn(WithHours, "ProductiveHRS", each [DemandHRS] * [#"Direct Care %"], type number),
    BufferedSourceCells = Table.Buffer(WithProductiveHours)
in
    BufferedSourceCells;

// Query: Distributed FTE Prepare
// Purpose: Sum validated DC-role cells into the configured output Role Group at fortnight day/shift grain.
// Output: Forty-two cells per enabled group; roster and productive hours remain separate measures.
shared #"Distributed FTE Prepare" =
let
    Source = #"Distributed FTE Source Cells",
    GroupedByUnit = Table.Group(Source, {"Facility-Abbrev", "Role", "FortnightDayIndex", "Shift", "Week No"}, {
        {"SourceRoles", each Text.Combine(List.Sort(List.Distinct([SourceRole])), ", "), type text},
        {"SourceFTE", each List.Sum([SourceFTE]), type number},
        {"DemandHRS", each List.Sum([DemandHRS]), type number},
        {"ProductiveHRS", each List.Sum([ProductiveHRS]), type number}
    }),
    Grouped = Table.RenameColumns(GroupedByUnit, {{"Facility-Abbrev", "Facility"}}),
    ExpectedCells = List.Count(#"Demand Roles") * 42,
    ValidGroups = List.Sort(List.Distinct(Grouped[Role])) = #"Demand Roles",
    Validated = if not ValidGroups or Table.RowCount(Grouped) <> ExpectedCells
        or Table.RowCount(Table.Distinct(Table.SelectColumns(Grouped, {"Role", "FortnightDayIndex", "Shift"}))) <> ExpectedCells then
            error "Every enabled Role Group must have a complete 14-day, three-shift source pattern for the modeled facility."
        else Grouped,
    BufferedPattern = Table.Buffer(Validated)
in
    BufferedPattern;

// Query: Demand Planning Periods
// Purpose: Read the full planning horizon and shared shift periods from the Settings calendar.
// Inputs: IMPORT PermutationDimensions, generated by Settings DateList and Shifts.
// Output: One row per configured date/shift, retaining the Settings Day and Period values.
// Notes: Read before filtering roles so missing enabled-role dates cannot shorten the expected horizon.
shared #"Demand Planning Periods" =
let
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
    BufferedPeriods;

// Query: Demand Calendar CHECK
// Purpose: Show missing or duplicate Settings periods separately for each enabled Role Group.
// Output: One diagnostic row per enabled role, including missing and duplicated period IDs.
// Notes: Compare exact configured labels with the roles in the saved Settings calendar.
shared #"Demand Calendar CHECK" =
let
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
    BufferedDiagnostics;

// Query: Demand Calendar Prepare
// Purpose: Require one cell per enabled group and Settings period, then map dates onto the source fortnight.
// Notes: Preserve the complete Settings horizon, including a final partial fortnight.
shared #"Demand Calendar Prepare" =
let
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
    BufferedCalendar;

// Query: Permutation DateTimeRoleShift
// Purpose: Attach one validated settings shift span to each enabled role-group calendar cell.
// Notes: Preserve existing overnight handling and require DurationOfShifts to equal the timestamp span.
shared #"Permutation DateTimeRoleShift" =
let
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
    BufferedTiming;

// Query: Demand Extraction Prepare
// Purpose: Apply the source fortnight pattern across the Settings horizon and calculate shift-average attendance.
// Output: One enabled role-group demand row per period, retaining contributing roles and audit measures.
shared #"Demand Extraction Prepare" =
let
    Joined = Table.NestedJoin(#"Permutation DateTimeRoleShift", {"Role", "FortnightDayIndex", "Shift"},
        #"Distributed FTE Prepare", {"Role", "FortnightDayIndex", "Shift"}, "Pattern", JoinKind.LeftOuter),
    InvalidMatches = Table.SelectRows(Joined, each Table.RowCount([Pattern]) <> 1),
    Validated = if not Table.IsEmpty(InvalidMatches) then
        error "Every calendar cell must match exactly one validated fortnight allocation." else Joined,
    Expanded = Table.ExpandTableColumn(Validated, "Pattern",
        {"Facility", "SourceRoles", "SourceFTE", "DemandHRS", "ProductiveHRS", "Week No"},
        {"Facility", "SourceRoles", "SourceFTE", "DemandHRS", "ProductiveHRS", "Week No"}),
    WithUnit = Table.AddColumn(Expanded, "Unit", each [Facility], type text),
    // A source standard-FTE equivalent is converted to attendance across the actual shift span.
    // Leave fractional attendance unrounded so interval integration recovers the roster hours.
    WithAttendance = Table.AddColumn(WithUnit, "DemandFTE", each [DemandHRS] / [DurationOfShifts], type number),
    InvalidValues = Table.SelectRows(WithAttendance, each not #"Demand Finite Number"([DemandFTE])
        or [DemandFTE] < 0 or not #"Demand Finite Number"([DemandHRS])
        or Number.Abs([DemandFTE] * [DurationOfShifts] - [DemandHRS]) > 0.0000001),
    ValidatedAttendance = if not Table.IsEmpty(InvalidValues) then error "Demand hours/attendance conversion failed."
        else WithAttendance,
    BufferedDemand = Table.Buffer(ValidatedAttendance)
in
    BufferedDemand;

// Query: Demand Reconciliation Targets
// Purpose: Sum source-role hours only for the days and shifts covered by each planning fortnight.
// Output: Expected cell count, roster hours and productive hours per facility/group/planning fortnight.
// Notes: A final partial fortnight uses its actual source days, not a prorated full-fortnight average.
shared #"Demand Reconciliation Targets" =
let
    // Read the original DC-role cells independently of the grouped demand calculation.
    Joined = Table.NestedJoin(#"Demand Calendar Prepare", {"Role", "FortnightDayIndex", "Shift"},
        #"Distributed FTE Source Cells", {"Role", "FortnightDayIndex", "Shift"}, "SourceCells", JoinKind.LeftOuter),
    MissingCells = Table.SelectRows(Joined, each Table.IsEmpty([SourceCells])),
    Validated = if not Table.IsEmpty(MissingCells) then error "A planning period has no source-role reconciliation cells."
        else Joined,
    WithFacility = Table.AddColumn(Validated, "Facility", each #"Demand Facility", type text),
    WithRosterHours = Table.AddColumn(WithFacility, "DemandHRS", each List.Sum([SourceCells][DemandHRS]), type number),
    WithProductiveHours = Table.AddColumn(WithRosterHours, "ProductiveHRS", each List.Sum([SourceCells][ProductiveHRS]), type number),
    Targets = Table.Group(WithProductiveHours, {"Facility", "Role", "PlanningFortnight"}, {
        {"ExpectedCellCount", each Table.RowCount(_), Int64.Type},
        {"ExpectedRosterHours", each List.Sum([DemandHRS]), type number},
        {"ExpectedProductiveHours", each List.Sum([ProductiveHRS]), type number}
    }),
    BufferedTargets = Table.Buffer(Targets)
in
    BufferedTargets;

// Query: DemandExtraction_RECONCILIATION
// Purpose: Reconcile every full or partial planning fortnight against its source-role hour targets.
// Notes: Targets follow Settings dates; excluded roles do not inflate retained demand.
shared DemandExtraction_RECONCILIATION =
let
    Targets = #"Demand Reconciliation Targets",
    Actual = Table.Group(#"Demand Extraction Prepare", {"Facility", "Role", "PlanningFortnight"}, {
        {"CellCount", each Table.RowCount(_), Int64.Type},
        {"RosterHours", each List.Sum([DemandHRS]), type number},
        {"IntegratedHours", each List.Sum(List.Transform(Table.ToRecords(_),
            (R) => R[DemandFTE] * R[DurationOfShifts])), type number},
        {"ProductiveHours", each List.Sum([ProductiveHRS]), type number}
    }),
    // Start from the expected groups so a missing output group cannot disappear from the check.
    Keys = {"Facility", "Role", "PlanningFortnight"},
    Joined = Table.NestedJoin(Targets, Keys, Actual, Keys, "Actual", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Actual", {"CellCount", "RosterHours", "IntegratedHours", "ProductiveHours"}),
    WithResidual = Table.AddColumn(Expanded, "RosterHoursResidual", each [RosterHours] - [ExpectedRosterHours], type number),
    Result = Table.AddColumn(WithResidual, "Status", each
        if [CellCount] = [ExpectedCellCount] and Number.Abs([RosterHoursResidual]) <= 0.0000001
            and Number.Abs([IntegratedHours] - [ExpectedRosterHours]) <= 0.0000001
            and Number.Abs([ProductiveHours] - [ExpectedProductiveHours]) <= 0.0000001
        then "Pass" else "Error", type text),
    BufferedReconciliation = Table.Buffer(Result)
in
    BufferedReconciliation;

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
        CheckStage("Facility mapping and enabled unit", () =>
            let
                Mapping = #"Facility Source Code Mapping",
                EnabledUnit = FacilityAnalysisTABLE
            in
                if Table.IsEmpty(EnabledUnit) then error "Facility Analysis does not enable this Unit."
                else Table.RowCount(Mapping)),
        CheckStage("Linked role definitions and effort-analysis controls", () => Table.RowCount(#"Demand Role Mapping")),
        CheckStage("Enabled effort-management role groups", () => List.Count(#"Demand Roles")),
        CheckStage("Source schema and saved publication checks", () => #"Distributed FTE Source Validate"),
        CheckStage("Complete source DC-role allocation and profile agreement", () => Table.RowCount(#"Distributed FTE Source Cells")),
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
    UpstreamAttempt = try Table.Buffer(#"IMPORT Distributed FTE Checks"),
    Warnings = if UpstreamAttempt[HasError] then #table({"Severity", "Check", "Message", "Actual"}, {}) else
        Table.AddColumn(Table.SelectColumns(
            Table.SelectRows(UpstreamAttempt[Value], each [Severity] = "Warning"),
            {"Severity", "Check", "Message"}), "Actual", each null),
    CombinedChecks = Table.Combine({StageChecks, Warnings})
in
    CombinedChecks;

// Query: ShiftUnitDemandHRS
// Purpose: Publish validated effort-management Role Groups through the unchanged twelve-column downstream interface.
// Output: One row per enabled group and Settings period, preserving Date/Day/Period and shift timestamps.
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