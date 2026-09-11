// Power Query from: Shift-StaffDistribution.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Shift-StaffDistribution.xlsx
// Extracted: 2026-05-18T06:14:31.704Z

section Section1;

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

shared #"Folder-2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"Source SHifts" = let
    Source = Excel.Workbook(File.Contents(#"Folder-2Calculations" & "\Shifts.xlsx"), null, true),
    NAMES_INTERVALSLIST_Table = Source{[Item="NAMES_INTERVALSLIST",Kind="Table"]}[Data]
in
    NAMES_INTERVALSLIST_Table;

shared ShiftTypeDistribution = let
    Source = #"Source SHifts",
    #"Filtered Rows" = Table.SelectRows(Source, each ([StaffCount] = 1)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Name", "TimeDate", "ShiftType", "Role"}),
    #"Inserted Date" = Table.AddColumn(#"Removed Other Columns", "Date", each DateTime.Date([TimeDate]), type date),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Date",{"TimeDate"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Name", Order.Ascending}, {"Date", Order.Ascending}})
in
    #"Sorted Rows";

shared ShiftTypeEffort = let
    Source = #"Source SHifts",
    #"Filtered Rows1" = Table.SelectRows(Source, each ([StaffCount] = 1)),
    #"Inserted Date" = Table.AddColumn(#"Filtered Rows1", "Date", each DateTime.Date([TimeDate]), type date),
    #"Sorted Rows" = Table.Sort(#"Inserted Date",{{"Name", Order.Ascending}, {"TimeDate", Order.Ascending}}),
    #"Removed Duplicates" = Table.Distinct(#"Sorted Rows", {"Date", "Name"}),
    #"Grouped Rows" = Table.Group(#"Removed Duplicates", {"Name", "ShiftType", "Date", "Role"}, {{"ShiftEffort", each List.Sum([Effective Duration]), type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows", each ([Name] <> null))
in
    #"Filtered Rows";
