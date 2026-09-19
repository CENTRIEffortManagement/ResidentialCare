// Power Query from: 1-AllocationExtracted.xlsx
// Pathname: CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\1-AllocationExtracted.xlsx
// Extracted: 2026-08-26T09:58:02.786Z

section Section1;

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"Filter AGENCY" = Table.SelectRows(Source, each ([Employment Type] <> " " and [Employment Type] <> "Agency ")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filter AGENCY",{"Location", "Department", "Area", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Employee_Code"}),
    #"Extracted First Characters" = Table.TransformColumns(#"Removed Other Columns", {{"Location", each Text.Start(_, 2), type text}}),
    #"Renamed Columns5" = Table.RenameColumns(#"Extracted First Characters",{{"Role", "RoleShift"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns5",{{"Employee_Code", "Code"}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Renamed Columns",{{"Employee Roster Name", "Name"}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Renamed Columns1",{{"Start Time", "Start"}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Renamed Columns2",{{"End Time", "End"}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Renamed Columns3",{{"End", "Finish"}, {"Break Length Minutes", "Break"}, {"Shift Net Length", "Hours"}}),
    Custom1 = Table.AddColumn(#"Renamed Columns4", "Role", each if [Department] <> "Care" and  [Department] <> "Catering" 
         then [RoleShift]
         else 
            Text.Start([RoleShift], Text.Length([RoleShift])-3)),
    #"Reordered Columns" = Table.ReorderColumns(Custom1,{"Code", "Name", "Date", "Start", "Finish", "Break", "Hours", "Location", "Department", "Area", "Role"}),
    #"Filtered ROLES" = Table.SelectRows(#"Reordered Columns", each ([Role] = "Asst in Nursing" or [Role] = "Asst in Nursing " or [Role] = "Asst in Nursing Med Comp" or [Role] = "Enrolled Nurse" or [Role] = "REGN" or [Role] = "REGN -  In Charge" or [Role] = "REGN - In Charge")),
    BUFFER = Table.Buffer(#"Filtered ROLES")
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
    #"Renamed Columns2" = Table.RenameColumns(Source,{{"Finish", "End"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns2",{{"Role", "RoleX"}}),
    // Temporarily use AINC4 for enrolled-nurse allocations so the third role branch receives them.
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Role", each if Text.Contains([RoleX], "REGN") then "RN" else if [RoleX] = "Asst in Nursing Med Comp" then "AINC4" else if Text.Contains([RoleX], "Asst") then "AIN" else if [RoleX] = "Enrolled Nurse" then "AINC4" else [RoleX]),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Role] = "AIN" or [Role] = "AINC4" or [Role] = "Asst in Nursing " or [Role] = "Enrolled Nurse" or [Role] = "RN")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Unit", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RoleX"})
in
    #"Removed Columns";

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

shared PathTABLE = UnitL1PathTABLE;

shared Path = Unit1Path;

shared DayAdjustment = let
    Source = Excel.CurrentWorkbook(){[Name="DayAdjustment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayAdjustment", Int64.Type}}),
    DayAdjustment1 = #"Changed Type"{0}[DayAdjustment]
in
    DayAdjustment1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents(Unit1Path & "\1. Input\Allocation\Published Roster.xlsx"), null, true),
    Beaudesert_Sheet = Source{[Item="Beaudesert",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Beaudesert_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", Int64.Type}})
in
    #"Changed Type";

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

shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        SelectedColumns = Table.SelectColumns(Source, {"FilePath"}),
        ChangedType = Table.TransformColumnTypes(SelectedColumns, {{"FilePath", type text}}),
        ReplacedValue = Table.TransformColumns(ChangedType, {{"FilePath", each if _ = null then null else Text.Replace(_, "/", "\"), type text}}),
        ValidatedTable = if Table.RowCount(ReplacedValue) = 1 then ReplacedValue else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedTable)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if RawFilePathValue = null or Text.Trim(RawFilePathValue) = "" then error "FilePathUrl[FilePath] must contain the current workbook path." else RawFilePathValue,
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

shared #"UnitL1PathTABLE (2)" = UnitL1PathTABLE;

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

shared Unit = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;
