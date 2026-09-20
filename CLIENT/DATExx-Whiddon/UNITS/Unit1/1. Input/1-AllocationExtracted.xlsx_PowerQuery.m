// Power Query from: 1-AllocationExtracted.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\1-AllocationExtracted.xlsx
// Extracted: 2026-09-20T05:13:03.399Z

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

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"REMOVE Filtered BD" = Table.SelectRows(Source, each ([Location] = "BD: Beaudesert")),
    #"Filter AGENCY" = Table.SelectRows(#"REMOVE Filtered BD", each ([Employment Type] <> " " and [Employment Type] <> "Agency ")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filter AGENCY",{"Location", "Department", "Area", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Employee_Code"}),
    #"Extracted First Characters" = Table.TransformColumns(#"Removed Other Columns", {{"Location", each Text.Start(_, 2), type text}}),
    #"Renamed Columns5" = Table.RenameColumns(#"Extracted First Characters",{{"Role", "RoleShift"}}),
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
// Purpose: Prepare roster allocations using RN, AIN and the existing AINC4 role value.
shared AllocationExtraction = let
    Source = #"Allocation Prepare",
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
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayAdjustment", Int64.Type}}),
    DayAdjustment1 = #"Changed Type"{0}[DayAdjustment]
in
    DayAdjustment1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents(Unit1Path & "\1. Input\Allocation\Published Roster.xlsx"), null, true),
    Combined_Sheet = Source{[Item="Combined",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(Combined_Sheet, [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", Int64.Type}})
in
    #"Changed Type1";

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
shared ResDoubleRoles = let
    Source = AllocationExtraction,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Count] = 2)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "MultiRole", each "MultiRole")
in
    #"Added Custom";

shared ResRolesProportion = let
    Source = Table.NestedJoin(AllocationExtraction, {"Name"}, ResDoubleRoles, {"Name"}, "ResDoubleRoles", JoinKind.LeftOuter),
    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(Source, "ResDoubleRoles", {"Name"}, {"Name.1"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDoubleRoles", each ([Name.1] <> null)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Name", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Name", "Date", "Start", "Hours", "Role"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"Name", "Role"}, {{"Hours", each List.Sum([Hours]), type nullable number}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", List.Distinct(#"Grouped Rows"[Role]), "Role", "Hours", List.Sum),
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