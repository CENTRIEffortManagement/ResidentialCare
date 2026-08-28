// Power Query from: Settings Data.xlsx
// Pathname: CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx
// Extracted: 2026-08-27T06:57:45.754Z

section Section1;

shared #"Dates Listed" = let
    Source = DateRoleShiftAllocation,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Allocation"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Date", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared #"Min Date" = let
    Source = #"Dates Listed",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] >= DateFrom)),
    Custom1 = Table.Min(#"Filtered Rows","Date"),
    Date = Custom1[Date]
in
    Date;

shared DateList = let
    Source = #"Dates Listed",
    #"Changed Type4" = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Custom1 = #"Changed Type4"[Date],
    Custom2 = List.Max(#"Changed Type4" [Date]),
    #"Converted to Table" = #table(1, {{Custom2}}),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Max"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Min", each #"Min Date"),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom1",{{"Max", type date}, {"Min", type date}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"Max", Int64.Type}, {"Min", Int64.Type}}),
    #"Added Custom3" = Table.AddColumn(#"Changed Type2", "Custom", each {[Min]..[Max]}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom3",{{"Max", type date}, {"Min", type date}}),
    #"Expanded Custom" = Table.ExpandListColumn(#"Changed Type1", "Custom"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Custom",{"Max", "Min"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Custom", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"Custom", "Date"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared PermutationDimensions = let
    Source = #"DateList",
    #"Added SHIFTS" = Table.AddColumn(Source, "Shifts", each Shifts),
    #"Expanded Shifts" = Table.ExpandListColumn(#"Added SHIFTS", "Shifts"),
    #"Added Index" = Table.AddIndexColumn(#"Expanded Shifts", "Period", 1, 1, Int64.Type),
    #"ADD ROLES" = Table.AddColumn(#"Added Index", "RolesList", each Roles),
    #"Expanded RolesList" = Table.ExpandListColumn(#"ADD ROLES", "RolesList"),
    #"Sorted Rows" = Table.Sort(#"Expanded RolesList",{{"Date", Order.Ascending}, {"Period", Order.Ascending}, {"RolesList", Order.Ascending}})
in
    #"Sorted Rows";

shared DateNameRoleShiftAllocation = let
    Source = AllocationExtracted,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Custom", each (if [End]<[Start]
then (1-[Start])+[End]
else [End]-[Start])*24),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each ([End] <> null)),
    #"Added Custom2" = Table.AddColumn(#"Filtered Rows", "TimeWorked", each if [Custom] >6 then [Custom] - .5
else [Custom]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom2", "Shift", each if [Start] < 0.604 then "AM" else if [Start] > 0.89 then "NIGHT" else "PM"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom1",{ "Start", "End", "Custom"})
in
    #"Removed Columns1";

shared DateRoleShiftAllocation = let
    Source = DateNameRoleShiftAllocation,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Role", "Shift"}, {{"Allocation", each List.Sum([TimeWorked]), type number}})
in
    #"Grouped Rows";

shared Folder = let
    Source = Unit1Path
in
    Source;

shared AllocationExtracted = let
    Source = Excel.Workbook(File.Contents(Folder & "\1. Input\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Date", type date}, {"Start", type time}, {"End", type time}, {"Break", Int64.Type}, {"Hours", type number}, {"Name", type text}, {"Code", Int64.Type}, {"Role", type text}})
in
    #"Changed Type";

shared Roles = let
    Source = Excel.CurrentWorkbook(){[Name="Roles"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ROLES", type text}}),
    ROLES = #"Changed Type"[ROLES]
in
    ROLES;

shared Shifts = let
    Source = Excel.CurrentWorkbook(){[Name="Shifts"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}}),
    #"Added Index1" = Table.AddIndexColumn(#"Changed Type", "Index.1", 1, 1, Int64.Type),
    Shifts1 = #"Added Index1"[Shifts]
in
    Shifts1;

shared DateFrom = let
    Source = Excel.CurrentWorkbook(){[Name="DateFrom"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateFrom", type date}}),
    DateFrom1 = #"Changed Type"{0}[DateFrom]
in
    DateFrom1;

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
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

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

