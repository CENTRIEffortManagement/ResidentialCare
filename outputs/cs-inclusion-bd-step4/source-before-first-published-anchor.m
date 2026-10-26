// Power Query from: 1-AllocationExtracted.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\1. Input\1-AllocationExtracted.xlsx
// Extracted: 2026-10-25T09:53:39.589Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Preserve the existing path-table interface using the single resolver in PathTABLE.
shared UnitL1PathTABLE = let
    Source = PathTABLE
in
    Source;

// Query: Unit1Path
// Purpose: Expose the resolved unit folder for relative external imports.
// Inputs: The Root Path row of UnitL1PathTABLE, ending in an approved workbook subfolder.
shared Unit1Path = let
    Source = PathTABLE,
    RootPathRow = Table.SelectRows(Source, each [Variable Name] = "Root Path"),
    WorkbookFolder = RootPathRow{0}[Value],
    UnitFolder =
        if Text.EndsWith(WorkbookFolder, "\1. Input", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\1. Input"))
        else if Text.EndsWith(WorkbookFolder, "\2. Calculations", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\2. Calculations"))
        else
            error "Root Path did not end in an expected Unit1 workbook folder: " & WorkbookFolder
in
    UnitFolder;

// Query: Unit
// Purpose: Expose the unit name from the authoritative path metadata.
shared Unit = let
    Source = PathTABLE,
    UnitRow = Table.SelectRows(Source, each [Variable Name] = "Unit"),
    UnitName = UnitRow{0}[Value]
in
    UnitName;

// Query: PathTABLE
// Purpose: Resolve the saved workbook path through the shared ResidentialCare sync mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value interface; Root Path is the workbook's containing folder.
shared PathTABLE = let
    NormalizePath = (value as nullable text) as nullable text =>
        let
            SlashNormalized = if value = null then null else Text.Replace(Text.Trim(value), "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,

    // Accept existing name casing, but require exactly one path input and never read a query-output table.
    WorkbookPathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
    WorkbookPathTable = if Table.RowCount(WorkbookPathInputs) = 1 then WorkbookPathInputs{0}[Content]
        else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
    // Excel exposes a single named cell as Column1; the standard table exposes FilePath.
    WorkbookPathColumn = if Table.HasColumns(WorkbookPathTable, {"FilePath"}) then Table.SelectColumns(WorkbookPathTable, {"FilePath"})
        else if Table.ColumnNames(WorkbookPathTable) = {"Column1"} then Table.RenameColumns(WorkbookPathTable, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or be a single named cell.",
    WorkbookPathRow = if Table.RowCount(WorkbookPathColumn) = 1 then WorkbookPathColumn{0} else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = WorkbookPathRow[FilePath],
    // An unsaved workbook or malformed path table must stop dependent imports.
    NormalizedWorkbookPath =
        if not Value.Is(RawFilePath, type text) then
            error "FilePathUrl[FilePath] must contain the current workbook path as text."
        else if Text.Trim(RawFilePath) = "" then
            error "FilePathUrl[FilePath] must contain the current workbook path."
        else
            NormalizePath(RawFilePath),

    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; isolate its final segment.
    InputWorkbookFolder = Text.BeforeDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputWorkbookSuffix = Text.AfterDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputFileName = if Text.StartsWith(InputWorkbookSuffix, "[") then Text.BetweenDelimiters(InputWorkbookSuffix, "[", "]") else InputWorkbookSuffix,
    // A copied path formula must identify this workbook before any external imports use it.
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "1-AllocationExtracted.xlsx") = 0 then InputWorkbookFolder & "\" & InputFileName
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate 1-AllocationExtracted.xlsx.",

    // Fixed public-machine location (%PUBLIC%); independent of the signed-in user.
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    // Reuse the small mapping table for all three candidate-root forms.
    BufferedMappings = Table.Buffer(CentriSyncPaths_Normalized),
    SharePointCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type text),
    SharePointDocumentsCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each if [SharepointRootUrl] = null or Text.Trim([SharepointRootUrl]) = "" then null else [SharepointRootUrl] & "\Shared Documents", type text),
    LocalCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    // Match whole path segments, ignoring case, so similarly prefixed folders cannot collide.
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
    LongestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    // Equally specific roots must agree on the destination; mapping row order is not a tie-breaker.
    LongestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    MatchedDestinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch =
        if Table.RowCount(SortedMatches) = 0 then
            error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath
        else if List.Count(MatchedDestinations) = 1 then
            LongestMatch
        else
            error "CentriSyncPaths contains conflicting mappings for: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    // The normalized path ends in the workbook filename, without a worksheet suffix.
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase),
    // UserName is optional display metadata; it does not participate in path resolution.
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = InputFileName,
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

// Query: Path
// Purpose: Preserve the existing scalar interface for the resolved unit folder.
shared Path = let
    Source = Unit1Path
in
    Source;

shared FacilityAnalysisTABLE = let
    Source = #"LINK Facility Analysis",
    #"Filtered Rows" = Table.SelectRows(Source, each ([EffortManagementAnalysis] = true)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"LeaveBalanceAnalysis"})
in
    #"Removed Columns";

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"Trimmed SPACES" = Table.TransformColumns(Source,{{"Location", Text.Trim, type text}, {"Role", Text.Trim, type text}, {"Employee Roster Name", Text.Trim, type text}}),
    #"Filter AGENCY" = Table.SelectRows(#"Trimmed SPACES", each ([Employment Type] <> " " and [Employment Type] <> "Agency ")),
    #"Merged Queries" = Table.NestedJoin(#"Filter AGENCY", {"Location"}, #"LINK Facilities", {"Title"}, "LINK Facilities", JoinKind.LeftOuter),
    #"Expanded LINK Facilities" = Table.ExpandTableColumn(#"Merged Queries", "LINK Facilities", {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded LINK Facilities", "AnalyseFacilityIdentify", each if [#"Facility-Abbrev"] = Unit then true else false),
    
    #"Filtered FACILITY" = Table.SelectRows(#"Added Conditional Column", each ([AnalyseFacilityIdentify] = true)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered FACILITY",{"AnalyseFacilityIdentify"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Removed Columns1",{{"Role", "RoleShift"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns5",{{"Employee_Code", "Code"}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Renamed Columns",{{"Employee Roster Name", "Name"}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Renamed Columns1",{{"Start Time", "Start"}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Renamed Columns2",{{"End Time", "End"}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Renamed Columns3",{{"End", "Finish"}, {"Break Length Minutes", "Break"}, {"Shift Net Length", "Hours"}}),
    BUFFER = Table.Buffer(#"Renamed Columns4")
in
    BUFFER;

shared #"Roles-Raw" = let
    Source = #"IMPORT Roster",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Role"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared RosteredDays = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", Int64.Type}}),
    #"Calculated Distinct Count" = List.NonNullCount(List.Distinct(#"Changed Type"[Date]))
in
    #"Calculated Distinct Count";

shared RosterStart = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Calculated Maximum" = List.Min(#"Removed Other Columns"[Date])
in
    #"Calculated Maximum";

// Query: AllocationExtraction
// Purpose: Apply existing role controls to published allocations and the filled Master Roster CS supplement.
shared AllocationExtraction = let
    Source = #"Allocation Combined Prepare",
    #"Merged Queries" = Table.NestedJoin(Source, {"RoleShift"}, #"LINK Roles", {"Roster Roles"}, "LINK Roles", JoinKind.LeftOuter),
    #"Expanded LINK Roles" = Table.ExpandTableColumn(#"Merged Queries", "LINK Roles", {"Role Group"}, {"Role"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded LINK Roles", {"Role"}, #"LINK RoleAnalysis", {"Roles"}, "LINK RoleAnalysis", JoinKind.LeftOuter),
    #"Expanded LINK RoleAnalysis" = Table.ExpandTableColumn(#"Merged Queries1", "LINK RoleAnalysis", {"Effort Management Analysis"}, {"Effort Management Analysis"}),
    #"Filtered ROLES EMA" = Table.SelectRows(#"Expanded LINK RoleAnalysis", each ([Effort Management Analysis] = true)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered ROLES EMA",{"Effort Management Analysis"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns1",{{"Finish", "End"}}),
    #"Added UNITBLANK" = Table.AddColumn(#"Renamed Columns2", "Unit", each null)
in
    #"Added UNITBLANK";

shared AllocationExtracted = let
    Source = AllocationExtraction,

    #"Merged ResDoubleRoles" = Table.NestedJoin(
        Source,
        {"Name"},
        ResDoubleRoles,
        {"Name"},
        "ResDoubleRoles",
        JoinKind.LeftOuter
    ),

    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(
        #"Merged ResDoubleRoles",
        "ResDoubleRoles",
        {"MultiRole"},
        {"MultiRole"}
    ),

    #"Renamed Columns" = Table.RenameColumns(
        #"Expanded ResDoubleRoles",
        {{"Name", "NameX"}}
    ),

    #"Inserted MULTIROLENAMES" = Table.AddColumn(
        #"Renamed Columns",
        "Name",
        each
            if [MultiRole] <> null then
                [NameX] & " (" & [Role] & ")"
            else
                [NameX],
        type text
    ),

    #"Removed Columns" = Table.RemoveColumns(
        #"Inserted MULTIROLENAMES",
        {"NameX", "MultiRole"}
    )

in
    #"Removed Columns";

shared DayAdjustment = let
    Source = Excel.CurrentWorkbook(){[Name="DayAdjustment"]}[Content],
    Column1 = Source{0}[Column1]
in
    Column1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents(Unit1Path & "\1. Input\Allocation\Published Roster.xlsx"), null, true),
    Combined_Sheet = Source{[Item="Roster Data",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(Combined_Sheet, [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type1")
in
    BUFFER;

shared #"LINK Roles" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"]),
    #"6b0b0767-44f4-4bb8-b116-f1e99b3476f0" = Source{[Id="6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"6b0b0767-44f4-4bb8-b116-f1e99b3476f0",{"Roster Roles", "DC Category", "DC Role", "Direct Care %", "Role Group"})
in
    #"Removed Other Columns";

shared AllocationExtractedCheck = let
    Source = AllocationExtraction,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom2" = Table.AddColumn(#"Changed Type1", "Subtraction", each if [End] < [Start]
then ([End]) + (1-[Start])
else  [End]-[Start]),
    #"Filtered Rows3" = Table.SelectRows(#"Added Custom2", each true),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows3", each true),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows2",{{"Subtraction", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Role"}, {{"Days", each List.Sum([Subtraction]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "AllocationTime", each [Days]*24/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "AllocatedEffort", each [AllocationTime]*0.9366)
in
    #"Added Custom1";

shared AllocatedStaffList = let
    Source = AllocationExtracted,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Role", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows";

[ Description = "BY PASSED" ]
// Query: ResDoubleRoles
// Purpose: Retain the existing name-suffix interface for workers allocated to two or more roles.
shared ResDoubleRoles = let
    Source = AllocationExtraction,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Count] >= 2)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "MultiRole", each "MultiRole")
in
    #"Added Custom";

// Query: ResRolesProportion
// Purpose: Preserve the legacy AIN/AINC4 proportion for workers allocated to both of those roles.
// Notes: Other role allocations must not enter this legacy two-role denominator.
shared ResRolesProportion = let
    OriginalRoleRows = Table.SelectRows(AllocationExtraction, each List.Contains({"AIN", "AINC4"}, [Role])),
    OriginalRoleCounts = Table.Group(OriginalRoleRows, {"Name"}, {{"Roles", each List.Count(List.Distinct([Role])), Int64.Type}}),
    OriginalRoleWorkers = Table.SelectColumns(Table.SelectRows(OriginalRoleCounts, each [Roles] = 2), {"Name"}),
    Source = Table.NestedJoin(OriginalRoleRows, {"Name"}, OriginalRoleWorkers, {"Name"}, "ResDoubleRoles", JoinKind.LeftOuter),
    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(Source, "ResDoubleRoles", {"Name"}, {"Name.1"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDoubleRoles", each ([Name.1] <> null)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Name", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Name", "Date", "Start", "Hours", "Role"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"Name", "Role"}, {{"Hours", each List.Sum([Hours]), type nullable number}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", {"AIN", "AINC4"}, "Role", "Hours", List.Sum),
    #"Inserted Addition" = Table.AddColumn(#"Pivoted Column", "Addition", each [AIN] + [AINC4], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Addition",{{"Addition", "Total"}}),
    #"Inserted Division" = Table.AddColumn(#"Renamed Columns", "AINC4HrsAvilPref", each [AINC4] / [Total], type number),
    #"Rounded Off" = Table.TransformColumns(#"Inserted Division",{{"AINC4HrsAvilPref", each Number.Round(_, 1), type number}})
in
    #"Rounded Off";

shared #"LINK RoleAnalysis" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"]),
    #"bad3beb8-7064-4e11-9edd-d62dac5d702d" = Source{[Id="bad3beb8-7064-4e11-9edd-d62dac5d702d"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"bad3beb8-7064-4e11-9edd-d62dac5d702d",{"Leave Balance Analysis", "Effort Management Analysis", "Roles"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Roles", type text}})
in
    #"Changed Type";

[ Description = "Convert Location attributes to Facility using" ]
shared #"LINK Facilities" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation=null, ApiVersion=15]),
    #"1987eadb-ad2e-491a-a927-e5585667d4c5" = Source{[Id="1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"1987eadb-ad2e-491a-a927-e5585667d4c5",{"Title", "field_1"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"field_1", "Facility-Abbrev"}})
in
    #"Renamed Columns";

shared #"LINK Facility Analysis" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation=null, ApiVersion=15]),
    #"17ed8e30-5707-4af2-8e21-d68eb8184287" = Source{[Id="17ed8e30-5707-4af2-8e21-d68eb8184287"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"17ed8e30-5707-4af2-8e21-d68eb8184287",{"Title", "EffortManagementAnalysis", "LeaveBalanceAnalysis"})
in
    #"Removed Other Columns";

// Query: Master Roster Portfolio Path
// Purpose: Derive the organisation folder from the existing authoritative resolved unit path.
// Notes: Existing path resolution and Published Roster imports are unchanged.
shared #"Master Roster Portfolio Path" =
let
    Segments = Text.Split(Text.TrimEnd(Unit1Path, "\"), "\"),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    Portfolio = if UnitsIndex > 0 and List.Count(Segments) = UnitsIndex + 2
        then Text.Combine(List.FirstN(Segments, UnitsIndex), "\")
        else error "The resolved allocation unit path must end in UNITS followed by the complete unit name."
in
    Portfolio;

// Query: IMPORT Master Roster
// Purpose: Land the approved organisation Master Roster workbook navigation unchanged.
shared #"IMPORT Master Roster" =
    Excel.Workbook(File.Contents(#"Master Roster Portfolio Path" & "\1. Input\Master Roster.xlsx"), null, true);

// Query: EXTRACT Master Roster Templates
// Purpose: Land the raw Roster Templates sheet unchanged.
shared #"EXTRACT Master Roster" =
    #"IMPORT Master Roster"{[Item = "Roster Templates", Kind = "Sheet"]}[Data];

// Query: INPUT MasterRosterWeek1Monday
// Purpose: Land the local Week-1 Monday named cell or one-column table unchanged.
shared #"INPUT MasterRosterWeek1Monday" =
    Excel.CurrentWorkbook(){[Name = "MasterRosterWeek1Monday"]}[Content];

// Query: MasterRosterWeek1Monday
// Purpose: Require one local Monday matching the Week-1 Monday used by BD demand.
// Notes: This input prevents a circular import from Settings into Allocation.
shared MasterRosterWeek1Monday =
let
    Source = #"INPUT MasterRosterWeek1Monday",
    Columns = Table.ColumnNames(Source),
    InputColumn = if Columns = {"Column1"} then "Column1"
        else if Columns = {"MasterRosterWeek1Monday"} then "MasterRosterWeek1Monday"
        else error "MasterRosterWeek1Monday must be one named cell or a one-column table with that column name.",
    Value = if Table.RowCount(Source) = 1 then Record.Field(Source{0}, InputColumn)
        else error "MasterRosterWeek1Monday must contain exactly one date.",
    Parsed = try Date.From(Value, "en-AU") otherwise null,
    Result = if Parsed <> null and Date.DayOfWeek(Parsed, Day.Monday) = 0 then Parsed
        else error "Set MasterRosterWeek1Monday to the same Week-1 Monday date used by BD demand."
in
    Result;

// Query: CS Allocation Finite Number
// Purpose: Validate supplied numeric quantities without substituting zero for missing values.
shared #"CS Allocation Finite Number" = (Value as any) as logical =>
    Value.Is(Value, type number) and not Number.IsNaN(Value) and Value <> #infinity and Value <> -#infinity;

// Query: CS Allocation Role Mapping Prepare
// Purpose: Resolve CS roster labels once, retaining the exact labels consumed by existing role joins.
// Notes: CSH is separate. Duplicate relevant definitions would multiply allocations and are rejected.
shared #"CS Allocation Role Mapping Prepare" =
let
    Analysis = Table.SelectRows(#"LINK RoleAnalysis", each [Roles] <> null and Text.Upper(Text.Trim([Roles])) = "CS"),
    Control = if Table.RowCount(Analysis) = 1 and Analysis{0}[Roles] = "CS"
            and Value.Is(Analysis{0}[Effort Management Analysis], type logical) then Analysis{0}[Effort Management Analysis]
        else error "Role Analysis must contain exactly one CS row with a logical Effort Management Analysis setting.",
    Definitions = Table.SelectColumns(#"LINK Roles", {"Roster Roles", "Role Group", "DC Category"}),
    WithKey = Table.AddColumn(Definitions, "RosterRoleKey", each
        if [Roster Roles] = null then null else Text.Upper(Text.Trim([Roster Roles])), type nullable text),
    CSDefinitions = Table.SelectRows(WithKey, each [Role Group] <> null and Text.Upper(Text.Trim([Role Group])) = "CS"),
    Invalid = Table.SelectRows(CSDefinitions, each [RosterRoleKey] = null or [RosterRoleKey] = ""
        or [Role Group] <> "CS" or [DC Category] = null or Text.Upper(Text.Trim([DC Category])) <> "NA"),
    Keys = List.Buffer(List.Distinct(CSDefinitions[RosterRoleKey])),
    Relevant = Table.SelectRows(WithKey, each List.Contains(Keys, [RosterRoleKey])),
    Counts = Table.Group(Relevant, {"RosterRoleKey"}, {{"Definitions", each Table.RowCount(_), Int64.Type}}),
    Validated = if not Control then Table.FirstN(CSDefinitions, 0)
        else if Table.IsEmpty(CSDefinitions) or not Table.IsEmpty(Invalid)
            or not Table.IsEmpty(Table.SelectRows(Counts, each [Definitions] <> 1)) then
            error "Enabled CS requires one unambiguous non-direct-care mapping per roster label."
        else CSDefinitions,
    Result = Table.RenameColumns(Table.SelectColumns(Validated, {"RosterRoleKey", "Roster Roles"}), {{"Roster Roles", "RoleShift"}})
in
    Table.Buffer(Result);

// Query: CS Allocation Facility Mapping Prepare
// Purpose: Resolve source locations to the existing current-unit location and facility labels.
shared #"CS Allocation Facility Mapping Prepare" =
let
    Definitions = Table.SelectColumns(#"LINK Facilities", {"Title", "Facility-Abbrev"}),
    WithKey = Table.AddColumn(Definitions, "SourceLocationKey", each
        if [Title] = null then null else Text.Upper(Text.Trim([Title])), type nullable text),
    UnitRows = Table.SelectRows(WithKey, each [#"Facility-Abbrev"] <> null
        and Text.Upper(Text.Trim([#"Facility-Abbrev"])) = Text.Upper(Unit)),
    Invalid = Table.SelectRows(UnitRows, each [SourceLocationKey] = null or [SourceLocationKey] = "" or [#"Facility-Abbrev"] <> Unit),
    Keys = List.Buffer(List.Distinct(UnitRows[SourceLocationKey])),
    Relevant = Table.SelectRows(WithKey, each List.Contains(Keys, [SourceLocationKey])),
    Counts = Table.Group(Relevant, {"SourceLocationKey"}, {{"Definitions", each Table.RowCount(_), Int64.Type}}),
    Validated = if Table.IsEmpty(UnitRows) or not Table.IsEmpty(Invalid)
            or not Table.IsEmpty(Table.SelectRows(Counts, each [Definitions] <> 1)) then
            error "Master CS locations require unique Facilities mappings to the complete current Unit."
        else UnitRows,
    Result = Table.RenameColumns(Table.SelectColumns(Validated, {"SourceLocationKey", "Title", "Facility-Abbrev"}), {{"Title", "Location"}})
in
    Table.Buffer(Result);

// Query: CS Allocation Source Prepare
// Purpose: Prepare all current-unit Master rows before choosing CS employees or source weeks.
// Notes: Retain all legitimate templates and a source-row key; no position deduplication is applied.
shared #"CS Allocation Source Prepare" =
let
    Headers = Table.PromoteHeaders(#"EXTRACT Master Roster", [PromoteAllScalars = true]),
    Selected = Table.SelectColumns(Headers, {"Master Template", "Template", "Location", "Department", "Area", "Role",
        "Employee Code", "Employee Name", "Week No", "Week Day", "Start Time", "End Time", "Break Length",
        "Shift Net Length", "Shift Type", "Non Attended"}),
    WithSourceRow = Table.AddIndexColumn(Selected, "MasterSourceRow", 1, 1, Int64.Type),
    WithLocationKey = Table.AddColumn(WithSourceRow, "SourceLocationKey", each
        if [Location] = null then null else Text.Upper(Text.Trim(Text.From([Location]))), type nullable text),
    WithoutRawLocation = Table.RemoveColumns(WithLocationKey, {"Location"}),
    JoinedUnit = Table.NestedJoin(WithoutRawLocation, {"SourceLocationKey"}, #"CS Allocation Facility Mapping Prepare",
        {"SourceLocationKey"}, "FacilityMapping", JoinKind.Inner),
    ExpandedUnit = Table.ExpandTableColumn(JoinedUnit, "FacilityMapping", {"Location", "Facility-Abbrev"}),
    WithRoleKey = Table.AddColumn(ExpandedUnit, "RosterRoleKey", each
        if [Role] = null then null else Text.Upper(Text.Trim(Text.From([Role]))), type nullable text)
in
    WithRoleKey;

// Query: CS Allocation Weeks Prepare
// Purpose: Map the complete unit source fortnight to ordered weeks 1 and 2 before employee exclusions.
shared #"CS Allocation Weeks Prepare" =
let
    Selected = Table.SelectColumns(#"CS Allocation Source Prepare", {"Week No"}),
    Parsed = Table.TransformColumns(Selected, {{"Week No", each try Number.From(_) otherwise null, type nullable number}}),
    Invalid = Table.SelectRows(Parsed, each not #"CS Allocation Finite Number"([Week No]) or Number.RoundDown([Week No]) <> [Week No]),
    DistinctWeeks = Table.Sort(Table.Distinct(Parsed), {{"Week No", Order.Ascending}}),
    Validated = if not Table.IsEmpty(Invalid) or Table.RowCount(DistinctWeeks) <> 2 then
            error "The complete current-unit Master Roster must supply exactly two ordered numeric source weeks."
        else DistinctWeeks,
    Result = Table.AddIndexColumn(Validated, "FortnightWeek", 1, 1, Int64.Type)
in
    Table.Buffer(Result);

// Query: CS Allocation Rows Prepare
// Purpose: Retain filled CS source positions and their unchanged net hours, breaks and clock times.
// Notes: Blank employee codes denote unfilled demand and are excluded only from this allocation supplement.
shared #"CS Allocation Rows Prepare" =
let
    JoinedRoles = Table.NestedJoin(#"CS Allocation Source Prepare", {"RosterRoleKey"}, #"CS Allocation Role Mapping Prepare",
        {"RosterRoleKey"}, "CSMapping", JoinKind.Inner),
    ExpandedRole = Table.ExpandTableColumn(JoinedRoles, "CSMapping", {"RoleShift"}),
    Assigned = Table.SelectRows(ExpandedRole, each [Employee Code] <> null and Text.Trim(Text.From([Employee Code], "en-AU")) <> ""),
    WithEmployeeID = Table.AddColumn(Assigned, "MasterEmployeeID", each
        let Parsed = try Number.From([Employee Code]) otherwise null
        in if #"CS Allocation Finite Number"(Parsed) and Number.RoundDown(Parsed) = Parsed then Int64.From(Parsed)
            else error "A filled Master CS position has an invalid non-integer Employee Code.", Int64.Type),
    Typed = Table.TransformColumnTypes(WithEmployeeID, {{"Employee Name", type text}, {"Week No", type number},
        {"Week Day", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length", type number},
        {"Shift Net Length", type number}, {"Department", type text}, {"Area", type text}, {"Shift Type", type text}, {"Non Attended", type logical}}),
    WithDay = Table.AddColumn(Typed, "DayOfWeek", each List.PositionOf(
        {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"},
        if [Week Day] = null then null else Text.Upper(Text.Trim([Week Day]))) + 1, Int64.Type),
    Invalid = Table.SelectRows(WithDay, each not List.Contains({1..7}, [DayOfWeek])
        or not Value.Is([Start Time], type time) or not Value.Is([End Time], type time)
        or not #"CS Allocation Finite Number"([Shift Net Length]) or [Shift Net Length] < 0
        or not #"CS Allocation Finite Number"([Break Length]) or [Break Length] < 0 or Number.RoundDown([Break Length]) <> [Break Length]
        or ([Start Time] = [End Time] and [Shift Net Length] > 0)),
    Validated = if Table.IsEmpty(Invalid) then WithDay
        else error "Filled Master CS positions require valid weekdays, clock times, nonnegative net hours and break minutes.",
    JoinedWeeks = Table.NestedJoin(Validated, {"Week No"}, #"CS Allocation Weeks Prepare", {"Week No"}, "WeekMapping", JoinKind.LeftOuter),
    MissingWeeks = Table.SelectRows(JoinedWeeks, each Table.RowCount([WeekMapping]) <> 1),
    ValidWeeks = if Table.IsEmpty(MissingWeeks) then JoinedWeeks else error "Every filled CS row must match its source fortnight week.",
    ExpandedWeeks = Table.ExpandTableColumn(ValidWeeks, "WeekMapping", {"FortnightWeek"}),
    Narrow = Table.SelectColumns(ExpandedWeeks, {"MasterSourceRow", "Master Template", "Template", "Location", "Facility-Abbrev",
        "Department", "Area", "RoleShift", "MasterEmployeeID", "Employee Name", "FortnightWeek", "DayOfWeek",
        "Start Time", "End Time", "Break Length", "Shift Net Length", "Shift Type", "Non Attended"})
in
    Table.Buffer(Narrow);

// Query: CS Allocation Identity Prepare
// Purpose: Resolve Master CS identities to existing published employee codes and names without changing published rows.
// Notes: Only identities used by the supplement are checked. New workers retain their actual Master code and name.
shared #"CS Allocation Identity Prepare" =
let
    Source = #"CS Allocation Rows Prepare",
    EmployeeIDs = List.Buffer(List.Distinct(Source[MasterEmployeeID])),
    Published = Table.SelectColumns(#"Allocation Prepare", {"Code", "Employee Code", "Name"}),
    WithAliases = Table.AddColumn(Published, "MasterEmployeeID", each List.Distinct(List.RemoveNulls({[Code], [Employee Code]})), type list),
    ExpandedAliases = Table.ExpandListColumn(WithAliases, "MasterEmployeeID"),
    RelevantAliases = Table.SelectRows(ExpandedAliases, each List.Contains(EmployeeIDs, [MasterEmployeeID])),
    DistinctAliases = Table.Distinct(Table.SelectColumns(RelevantAliases, {"MasterEmployeeID", "Code", "Name"})),
    AliasCounts = Table.Group(DistinctAliases, {"MasterEmployeeID"}, {{"Matches", each Table.RowCount(_), Int64.Type}}),
    InvalidAliases = Table.SelectRows(DistinctAliases, each [Code] = null or [Name] = null or Text.Trim([Name]) = ""),
    ValidAliases = if not Table.IsEmpty(InvalidAliases) or not Table.IsEmpty(Table.SelectRows(AliasCounts, each [Matches] <> 1)) then
            error "A Master CS employee code resolves to missing or conflicting published worker identities."
        else DistinctAliases,
    Joined = Table.NestedJoin(Source, {"MasterEmployeeID"}, ValidAliases, {"MasterEmployeeID"}, "PublishedIdentity", JoinKind.LeftOuter),
    WithCode = Table.AddColumn(Joined, "Code", each if Table.IsEmpty([PublishedIdentity]) then [MasterEmployeeID] else [PublishedIdentity]{0}[Code], Int64.Type),
    WithName = Table.AddColumn(WithCode, "Name", each if Table.IsEmpty([PublishedIdentity]) then
        (if [Employee Name] = null then null else Text.Trim([Employee Name])) else [PublishedIdentity]{0}[Name], type nullable text),
    WithNameKey = Table.AddColumn(WithName, "NameKey", each if [Name] = null then null else Text.Upper(Text.Trim([Name])), type nullable text),
    WithMasterNameKey = Table.AddColumn(WithNameKey, "MasterNameKey", each
        if [Employee Name] = null then null else Text.Upper(Text.Trim([Employee Name])), type nullable text),
    NameKeys = List.Buffer(List.Distinct(WithMasterNameKey[NameKey] & WithMasterNameKey[MasterNameKey])),
    PublishedNameKeys = Table.AddColumn(Published, "NameKey", each if [Name] = null then null else Text.Upper(Text.Trim([Name])), type nullable text),
    RelevantNames = Table.SelectRows(PublishedNameKeys, each List.Contains(NameKeys, [NameKey])),
    MasterNames = Table.RenameColumns(Table.SelectColumns(Table.SelectRows(WithMasterNameKey,
        each [MasterNameKey] <> null and [MasterNameKey] <> ""), {"Code", "MasterNameKey"}), {{"MasterNameKey", "NameKey"}}),
    AllNames = Table.Distinct(Table.Combine({Table.SelectColumns(WithMasterNameKey, {"Code", "NameKey"}),
        MasterNames, Table.SelectColumns(RelevantNames, {"Code", "NameKey"})})),
    NameCounts = Table.Group(AllNames, {"NameKey"}, {{"Codes", each List.Count(List.Distinct([Code])), Int64.Type}}),
    CodeNames = Table.Distinct(Table.SelectColumns(WithMasterNameKey, {"Code", "Name"})),
    CodeNameCounts = Table.Group(CodeNames, {"Code"}, {{"Names", each Table.RowCount(_), Int64.Type}}),
    InvalidNames = Table.SelectRows(WithMasterNameKey, each [NameKey] = null or [NameKey] = ""),
    Validated = if not Table.IsEmpty(InvalidNames) or not Table.IsEmpty(Table.SelectRows(NameCounts, each [Codes] <> 1))
            or not Table.IsEmpty(Table.SelectRows(CodeNameCounts, each [Names] <> 1)) then
            error "Filled Master CS allocations require one unambiguous employee code and name; inspect the source and published identity matches."
        else WithMasterNameKey,
    Result = Table.RemoveColumns(Validated, {"PublishedIdentity", "NameKey", "MasterNameKey"})
in
    Table.Buffer(Result);

// Query: CS Allocation Dates Prepare
// Purpose: Repeat the agreed source fortnight within the original Published Roster date boundaries.
shared #"CS Allocation Dates Prepare" =
let
    // Consume the complete source-week prerequisite even when every CS position is unfilled.
    SourceWeeks = Table.RowCount(#"CS Allocation Weeks Prepare"),
    PublishedDates = if SourceWeeks = 2 then List.Sort(List.Distinct(List.RemoveNulls(#"Allocation Prepare"[Date])))
        else error "Master CS allocation requires a valid two-week source pattern.",
    Start = if List.IsEmpty(PublishedDates) then error "Published Roster must supply the allocation date range." else List.First(PublishedDates),
    Finish = List.Last(PublishedDates),
    Dates = Table.FromColumns({List.Dates(Start, Duration.Days(Finish - Start) + 1, #duration(1, 0, 0, 0))}, type table [Date = date]),
    // Normalize modulo twice so dates before the anchor still map to source days 1 through 14.
    WithDayIndex = Table.AddColumn(Dates, "FortnightDayIndex", each
        Number.Mod(Number.Mod(Duration.Days([Date] - MasterRosterWeek1Monday), 14) + 14, 14) + 1, Int64.Type),
    WithWeek = Table.AddColumn(WithDayIndex, "FortnightWeek", each Number.IntegerDivide([FortnightDayIndex] - 1, 7) + 1, Int64.Type),
    WithDay = Table.AddColumn(WithWeek, "DayOfWeek", each Number.Mod([FortnightDayIndex] - 1, 7) + 1, Int64.Type)
in
    Table.Buffer(WithDay);

// Query: CS Allocation Aligned Prepare
// Purpose: Date filled CS positions and match the complete existing published-preparation column contract.
// Notes: Net hours are copied unchanged. Missing publication/employment metadata is typed null.
shared #"CS Allocation Aligned Prepare" =
let
    JoinedDates = Table.NestedJoin(#"CS Allocation Identity Prepare", {"FortnightWeek", "DayOfWeek"},
        #"CS Allocation Dates Prepare", {"FortnightWeek", "DayOfWeek"}, "PlanningDates", JoinKind.Inner),
    Dated = Table.ExpandTableColumn(JoinedDates, "PlanningDates", {"Date"}),
    UniqueKeys = Table.Distinct(Table.SelectColumns(Dated, {"MasterSourceRow", "Date"})),
    Validated = if Table.RowCount(UniqueKeys) = Table.RowCount(Dated) then Dated
        else error "A Master CS source position was expanded more than once for the same date.",
    Records = List.Transform(Table.ToRecords(Validated), each [
        Location = [Location], #"Pay Company" = null, Department = [Department], Area = [Area], Name = [Name],
        #"Employment Type" = null, RoleShift = [RoleShift], #"Employee Code" = [MasterEmployeeID], Date = [Date],
        #"Day Of Week" = Date.DayOfWeekName([Date], "en-AU"), #"Shift Type" = [Shift Type],
        Start = [Start Time], Finish = [End Time], Break = Int64.From([Break Length]),
        #"Shift Length" = Duration.TotalHours([End Time] - [Start Time]) + (if [End Time] < [Start Time] then 24 else 0),
        Hours = [Shift Net Length], Rate = null, Published = null, #"Published At" = null, #"Published By" = null,
        #"Non Attended" = [Non Attended], Status = null, Code = [Code], #"Facility-Abbrev" = [#"Facility-Abbrev"]]),
    PublishedColumns = Table.ColumnNames(#"Allocation Prepare"),
    Aligned = Table.FromRecords(Records, PublishedColumns, MissingField.UseNull),
    Typed = Table.TransformColumnTypes(Aligned, {{"Location", type text}, {"Pay Company", type text}, {"Department", type text},
        {"Area", type text}, {"Name", type text}, {"Employment Type", type text}, {"RoleShift", type text},
        {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text},
        {"Start", type time}, {"Finish", type time}, {"Break", Int64.Type}, {"Shift Length", type number}, {"Hours", type number},
        {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text},
        {"Non Attended", type logical}, {"Status", type text}, {"Code", Int64.Type}}),
    Result = Value.ReplaceType(Typed, Value.Type(#"Allocation Prepare"))
in
    Table.Buffer(Result);

// Query: CS Allocation Reconciliation Prepare
// Purpose: Compare dated row/hour totals with the filled source pattern repeated over the published date range.
shared #"CS Allocation Reconciliation Prepare" =
let
    DateCounts = Table.Group(#"CS Allocation Dates Prepare", {"FortnightWeek", "DayOfWeek"}, {{"Dates", each Table.RowCount(_), Int64.Type}}),
    Joined = Table.NestedJoin(#"CS Allocation Rows Prepare", {"FortnightWeek", "DayOfWeek"}, DateCounts,
        {"FortnightWeek", "DayOfWeek"}, "Occurrences", JoinKind.LeftOuter),
    WithCount = Table.AddColumn(Joined, "ExpectedRows", each if Table.IsEmpty([Occurrences]) then 0 else [Occurrences]{0}[Dates], Int64.Type),
    WithHours = Table.AddColumn(WithCount, "ExpectedHours", each [Shift Net Length] * [ExpectedRows], type number),
    ExpectedRows = List.Sum({0} & WithHours[ExpectedRows]),
    ExpectedHours = List.Sum({0} & WithHours[ExpectedHours]),
    ActualRows = Table.RowCount(#"CS Allocation Aligned Prepare"),
    ActualHours = List.Sum({0} & #"CS Allocation Aligned Prepare"[Hours]),
    Result = #table(type table [Check = text, Actual = number, Expected = number], {
        {"Filled CS dated-row conservation", ActualRows, ExpectedRows},
        {"Filled CS net-hour conservation", ActualHours, ExpectedHours}})
in
    Result;

// Query: Allocation Combined Prepare
// Purpose: Append the reconciled CS supplement before the existing role-analysis and multi-role processing.
// Notes: Published CS triggers an explicit source-cutover decision; never count both sources silently.
shared #"Allocation Combined Prepare" =
let
    Published = #"Allocation Prepare",
    Mapping = #"CS Allocation Role Mapping Prepare",
    Result = if Table.IsEmpty(Mapping) then Published else
        let
            WithRoleKey = Table.AddColumn(Table.SelectColumns(Published, {"RoleShift"}), "RosterRoleKey", each
                if [RoleShift] = null then null else Text.Upper(Text.Trim([RoleShift])), type nullable text),
            ExistingCS = Table.NestedJoin(WithRoleKey, {"RosterRoleKey"}, Mapping, {"RosterRoleKey"}, "CSMapping", JoinKind.Inner),
            Reconciliation = if not Table.IsEmpty(ExistingCS) then
                    error "Published Roster now contains CS. Review the Master-to-Published CS cutover before combining allocation sources."
                else #"CS Allocation Reconciliation Prepare",
            Failures = Table.SelectRows(Reconciliation, each Number.Abs([Actual] - [Expected]) > 0.0000001),
            Combined = if not Table.IsEmpty(Failures) then error "Master CS allocation row/hour reconciliation failed. Inspect CS Allocation_CHECK."
                else Table.Combine({Published, #"CS Allocation Aligned Prepare"}, Value.Type(Published))
        in Combined
in
    Result;

// Query: CS Allocation_CHECK
// Purpose: Expose the source-period filled/unfilled counts and dated allocation conservation for manual BD acceptance.
shared #"CS Allocation_CHECK" =
let
    Mapping = #"CS Allocation Role Mapping Prepare",
    Result = if Table.IsEmpty(Mapping) then #table(type table [Check = text, Severity = text, Actual = number, Expected = nullable number, Message = text],
        {{"CS allocation inclusion", "Pass", 0, 0, "CS is disabled in Role Analysis."}}) else
        let
            ScopedCS = Table.NestedJoin(#"CS Allocation Source Prepare", {"RosterRoleKey"}, Mapping, {"RosterRoleKey"}, "CSMapping", JoinKind.Inner),
            Filled = Table.RowCount(#"CS Allocation Rows Prepare"),
            PublishedKeys = Table.AddColumn(Table.SelectColumns(#"Allocation Prepare", {"RoleShift"}), "RosterRoleKey", each
                if [RoleShift] = null then null else Text.Upper(Text.Trim([RoleShift])), type nullable text),
            PublishedCS = Table.RowCount(Table.NestedJoin(PublishedKeys, {"RosterRoleKey"}, Mapping, {"RosterRoleKey"}, "CSMapping", JoinKind.Inner)),
            SourceRows = #table(type table [Check = text, Severity = text, Actual = number, Expected = nullable number, Message = text], {
                {"Filled CS source positions", "Pass", Filled, null, "Included before date expansion."},
                {"Unfilled CS source positions", "Pass", Table.RowCount(ScopedCS) - Filled, null, "Blank employee codes excluded from allocation; retained in demand."},
                {"Published CS source overlap", if PublishedCS = 0 then "Pass" else "Error", PublishedCS, 0,
                    if PublishedCS = 0 then "Validated" else "Review the Master-to-Published CS cutover before combining allocation sources."}}),
            WithSeverity = Table.AddColumn(#"CS Allocation Reconciliation Prepare", "Severity", each
                if Number.Abs([Actual] - [Expected]) <= 0.0000001 then "Pass" else "Error", type text),
            WithMessage = Table.AddColumn(WithSeverity, "Message", each if [Severity] = "Pass" then "Validated" else "Source/dated allocation totals differ.", type text),
            Checks = Table.Combine({SourceRows, Table.SelectColumns(WithMessage, {"Check", "Severity", "Actual", "Expected", "Message"})})
        in Checks
in
    Result;