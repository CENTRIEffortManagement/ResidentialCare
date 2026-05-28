// Power Query from: Capacity.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Capacity.xlsx
// Extracted: 2026-05-18T06:14:12.971Z

section Section1;

shared #"IMPORT C#TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C#"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RosteredPeriodStatus"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityCapped", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C##TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C###TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&Role2&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared UPLOADRole = let
    Source = Excel.CurrentWorkbook(){[Name="Table6"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"RoleID", type text}, {"RoleName", type text}})
in
    #"Changed Type";

shared Role1 = let
    Source = UPLOADRole,
    RoleName = Source{0}[RoleName]
in
    RoleName;

shared Role2 = let
    Source = UPLOADRole,
    RoleName = Source{1}[RoleName]
in
    RoleName;

shared Role3 = let
    Source = UPLOADRole,
    RoleName = Source{2}[RoleName]
in
    RoleName;

shared Folder = let
    Source = Unit1Path
in
    Source;

shared Facility = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT Role1C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role1C##TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role1C###TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder &"\" & #"Role1"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared Capacity = let
    Source = #"IMPORT Role1C###TABLE1",
    #"Appended Query" = Table.Combine({Source, #"IMPORT C###TABLE2", #"IMPORT Role3C###TABLE"}),
    #"Filtered Rows" = Table.SelectRows(#"Appended Query", each ([Resource] = 94)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period", "Role"}, PermutationDimensions, {"Period", "RolesList"}, "PermutationDimensions", JoinKind.LeftOuter),
    #"Expanded PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries", "PermutationDimensions", {"Date"}, {"Date"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PermutationDimensions",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows"
   ;

shared AvailabilityDeveloped = let
    Source = Table.Combine({#"IMPORT OriginalResPeriodAvailabilityTABLE1", #"IMPORT Role1C#TABLE", #"IMPORT Role1C##TABLE1", #"IMPORT Role1C###TABLE1", #"IMPORT C#TABLE2", #"IMPORT C##TABLE2", #"IMPORT C###TABLE2", #"IMPORT Role3C#TABLE", #"IMPORT Role3C##TABLE", #"IMPORT Role3C###TABLE", #"IMPORT OriginalResPeriodAvailabilityTABLE2", #"IMPORT OriginalResPeriodAvailabilityTABLE3"}),
    #"Added Custom" = Table.AddColumn(Source, "Facility", each Facility),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",0,null,Replacer.ReplaceValue,{"Availability"})
in
    #"Replaced Value";

shared AvailabilityDevelopedMATRIX = let
    Source = AvailabilityDeveloped,
    #"Removed Columns" = Table.RemoveColumns(Source,{"ResAvailability", "ResMaxAvail"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[AvailabilityType]), "AvailabilityType", "Availability", List.Sum)
in
    #"Pivoted Column";

shared PermutationDimensions = let
    Source = Excel.Workbook(File.Contents(Folder&"\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared #"IMPORT Role3C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role3C##TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role3C###TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE3" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

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

shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;
