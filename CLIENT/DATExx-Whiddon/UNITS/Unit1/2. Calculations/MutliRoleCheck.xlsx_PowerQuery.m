// Power Query from: MutliRoleCheck.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\MutliRoleCheck.xlsx
// Extracted: 2026-05-18T06:14:26.566Z

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

shared FilePath = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT MultiRoles" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Capacity-ShiftAvailability.xlsx"), null, true),
    MultiRoles_Table = Source{[Item="MultiRoles",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(MultiRoles_Table,{{"Name", type text}, {"AINC4HrsAvilPref", type number}, {"MultiRole", type text}})
in
    #"Changed Type";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared #"IMPORT NAMES_INTERVALS" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Shifts.xlsx"), null, true),
    NAMES_INTERVALS_Table = Source{[Item="NAMES_INTERVALS",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(NAMES_INTERVALS_Table,{{"DayDate", type date}, {"Role", type text}, {"IntervalStart", type datetime}, {"TimeType", type text}, {"IntervalEnd", type datetime}, {"Name", type text}, {"IntervalID.1", Int64.Type}, {"AllocatedIntervals", Int64.Type}, {"IntervalDurationTemp", type number}, {"ShiftPeriod", type text}, {"Double Shift", type any}, {"Effective Duration", type number}, {"RealDuration", type number}, {"ShiftType", type text}})
in
    #"Changed Type";

shared #"IMPORT MultiRoles (2)" = let
    Source = #"IMPORT MultiRoles"
in
    Source;

shared MultiRoles = let
    Source = #"IMPORT MultiRoles",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns")
in
    #"Removed Duplicates";

shared MultiRoleMATRIX = let
    Source = #"IMPORT NAMES_INTERVALS",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"DayDate", "Role", "Name", "ShiftPeriod"}),
    BUFFER = Table.Buffer(#"Removed Other Columns"),
    #"Filtered Rows" = Table.SelectRows(BUFFER, each ([Name] <> "x")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "(")),
    #"Merged Queries1" = Table.NestedJoin(#"Filtered Rows1", {"Name"}, MultiRoles, {"Name"}, "MultiRoles", JoinKind.LeftOuter),
    #"Expanded MultiRoles" = Table.ExpandTableColumn(#"Merged Queries1", "MultiRoles", {"Name"}, {"Name.1"}),
    #"Added Custom" = Table.AddColumn(#"Expanded MultiRoles", "Custom", each "X"),
    #"Merged Queries" = Table.NestedJoin(#"Added Custom", {"DayDate", "Role", "ShiftPeriod"}, #"IMPORT PermutationDimensions", {"Date", "RolesList", "Shifts"}, "IMPORT PermutationDimensions", JoinKind.RightOuter),
    #"Expanded IMPORT PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions", {"Period"}, {"Period"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded IMPORT PermutationDimensions",{"DayDate", "ShiftPeriod"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Period", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Custom"),
    #"Sorted Rows1" = Table.Sort(#"Pivoted Column",{{"Name", Order.Ascending}})
in
    #"Sorted Rows1";
