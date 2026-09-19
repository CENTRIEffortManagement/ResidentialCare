// Canonical prepared source for: Capacity.xlsx
// Purpose: Preserve the current Capacity interfaces while routing the AIN C# stage to the merged B workbook.
// Notes: Prepared only. Do not synchronize until the AIN2 two-workbook acceptance gates and manual promotion approval pass.

section Section1;

// Query: IMPORT C#TABLE2
// Purpose: Import the C# compatibility table for the semantic role stored in Role2.
shared #"IMPORT C#TABLE2" = let
    SemanticRole = Text.Trim(#"Role2"),
    CSharpWorkbookName = if Comparer.OrdinalIgnoreCase(SemanticRole, "AIN") = 0 then "CapacityDistrib(B)-shifts.xlsx" else "CapacityDistrib(A.2)-shifts.xlsx",
    Source = Excel.Workbook(File.Contents(Folder & "\" & SemanticRole & "\" & CSharpWorkbookName), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C#"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RosteredPeriodStatus"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityCapped", "Availability"}})
in
    #"Renamed Columns";

// Query: IMPORT C##TABLE2
// Purpose: Import the C## table for the semantic role stored in Role2.
shared #"IMPORT C##TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

// Query: IMPORT C###TABLE2
// Purpose: Import the C### table for the semantic role stored in Role2.
shared #"IMPORT C###TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&Role2&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

// Query: IMPORT OriginalResPeriodAvailabilityTABLE2
// Purpose: Import original availability for the semantic role stored in Role2.
shared #"IMPORT OriginalResPeriodAvailabilityTABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

// Query: UPLOADRole
// Purpose: Read the ordered semantic role configuration from the current workbook.
shared UPLOADRole = let
    Source = Excel.CurrentWorkbook(){[Name="Table6"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"RoleID", type text}, {"RoleName", type text}})
in
    #"Changed Type";

// Query: Role1
// Purpose: Return the first configured semantic role name.
shared Role1 = let
    Source = UPLOADRole,
    RoleName = Source{0}[RoleName]
in
    RoleName;

// Query: Role2
// Purpose: Return the second configured semantic role name.
shared Role2 = let
    Source = UPLOADRole,
    RoleName = Source{1}[RoleName]
in
    RoleName;

// Query: Role3
// Purpose: Return the third configured semantic role name.
shared Role3 = let
    Source = UPLOADRole,
    RoleName = Source{2}[RoleName]
in
    RoleName;

// Query: Folder
// Purpose: Return the resolved Unit calculation folder.
shared Folder = let
    Source = Unit1Path
in
    Source;

// Query: Facility
// Purpose: Return the resolved Unit name for output attribution.
shared Facility = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: IMPORT Role1C#TABLE
// Purpose: Import the C# compatibility table for the semantic role stored in Role1.
shared #"IMPORT Role1C#TABLE" = let
    SemanticRole = Text.Trim(#"Role1"),
    CSharpWorkbookName = if Comparer.OrdinalIgnoreCase(SemanticRole, "AIN") = 0 then "CapacityDistrib(B)-shifts.xlsx" else "CapacityDistrib(A.2)-shifts.xlsx",
    Source = Excel.Workbook(File.Contents(Folder & "\" & SemanticRole & "\" & CSharpWorkbookName), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

// Query: IMPORT Role1C##TABLE1
// Purpose: Import the C## table for the semantic role stored in Role1.
shared #"IMPORT Role1C##TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

// Query: IMPORT Role1C###TABLE1
// Purpose: Import the C### table for the semantic role stored in Role1.
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

// Query: IMPORT OriginalResPeriodAvailabilityTABLE1
// Purpose: Import original availability for the semantic role stored in Role1.
shared #"IMPORT OriginalResPeriodAvailabilityTABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder &"\" & #"Role1"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

// Query: Capacity
// Purpose: Preserve the existing filtered C### output interface.
shared Capacity = let
    Source = #"IMPORT Role1C###TABLE1",
    #"Appended Query" = Table.Combine({Source, #"IMPORT C###TABLE2", #"IMPORT Role3C###TABLE"}),
    #"Filtered Rows" = Table.SelectRows(#"Appended Query", each ([Resource] = 94)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period", "Role"}, PermutationDimensions, {"Period", "RolesList"}, "PermutationDimensions", JoinKind.LeftOuter),
    #"Expanded PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries", "PermutationDimensions", {"Date"}, {"Date"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PermutationDimensions",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

// Query: AvailabilityDeveloped
// Purpose: Combine Original, C#, C## and C### availability for all configured roles.
shared AvailabilityDeveloped = let
    Source = Table.Combine({#"IMPORT OriginalResPeriodAvailabilityTABLE1", #"IMPORT Role1C#TABLE", #"IMPORT Role1C##TABLE1", #"IMPORT Role1C###TABLE1", #"IMPORT C#TABLE2", #"IMPORT C##TABLE2", #"IMPORT C###TABLE2", #"IMPORT Role3C#TABLE", #"IMPORT Role3C##TABLE", #"IMPORT Role3C###TABLE", #"IMPORT OriginalResPeriodAvailabilityTABLE2", #"IMPORT OriginalResPeriodAvailabilityTABLE3"}),
    #"Added Custom" = Table.AddColumn(Source, "Facility", each Facility),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",0,null,Replacer.ReplaceValue,{"Availability"})
in
    #"Replaced Value";

// Query: AvailabilityDevelopedMATRIX
// Purpose: Pivot developed availability stages for comparison.
shared AvailabilityDevelopedMATRIX = let
    Source = AvailabilityDeveloped,
    #"Removed Columns" = Table.RemoveColumns(Source,{"ResAvailability", "ResMaxAvail"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[AvailabilityType]), "AvailabilityType", "Availability", List.Sum)
in
    #"Pivoted Column";

// Query: PermutationDimensions
// Purpose: Import Period-to-Date and role dimensional attributes from Settings Data.
shared PermutationDimensions = let
    Source = Excel.Workbook(File.Contents(Folder&"\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

// Query: IMPORT Role3C#TABLE
// Purpose: Import the C# compatibility table for the semantic role stored in Role3.
shared #"IMPORT Role3C#TABLE" = let
    SemanticRole = Text.Trim(#"Role3"),
    CSharpWorkbookName = if Comparer.OrdinalIgnoreCase(SemanticRole, "AIN") = 0 then "CapacityDistrib(B)-shifts.xlsx" else "CapacityDistrib(A.2)-shifts.xlsx",
    Source = Excel.Workbook(File.Contents(Folder & "\" & SemanticRole & "\" & CSharpWorkbookName), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

// Query: IMPORT Role3C##TABLE
// Purpose: Import the C## table for the semantic role stored in Role3.
shared #"IMPORT Role3C##TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

// Query: IMPORT Role3C###TABLE
// Purpose: Import the C### table for the semantic role stored in Role3.
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

// Query: IMPORT OriginalResPeriodAvailabilityTABLE3
// Purpose: Import original availability for the semantic role stored in Role3.
shared #"IMPORT OriginalResPeriodAvailabilityTABLE3" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook through FilePathUrl and the public-machine CentriSyncPaths mapping.
// Output: One row per derived path attribute, including Root Path, Unit and FileName.
shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    FilePathUrl =
    let
        PathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
        Source = if Table.RowCount(PathInputs) = 1 then PathInputs{0}[Content]
            else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
        PathColumn = if Table.HasColumns(Source, {"FilePath"}) then Table.SelectColumns(Source, {"FilePath"})
            else if Table.ColumnNames(Source) = {"Column1"} then Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must contain a FilePath column or be a single named cell.",
        ValidatedRows = if Table.RowCount(PathColumn) = 1 then PathColumn
            else error "FilePathUrl must contain exactly one data row.",
        TypedPath = Table.TransformColumnTypes(ValidatedRows, {{"FilePath", type text}}),
        BufferedTable = Table.Buffer(TypedPath)
    in
        BufferedTable,

    RawFilePath = NormalizePath(FilePathUrl{0}[FilePath]),
    NonBlankFilePath = if RawFilePath = null or RawFilePath = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else RawFilePath,
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; retain only the workbook path.
    WorkbookPath = if Text.Contains(NonBlankFilePath, "[") then
        Text.BeforeDelimiter(NonBlankFilePath, "[") & Text.BetweenDelimiters(NonBlankFilePath, "[", "]")
        else NonBlankFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "Capacity.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate Capacity.xlsx.",
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
    LongestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0}
        else error "FilePathUrl has no usable CentriSyncPaths mapping. Check SharepointRootUrl and SyncedFolderRootPath for: " & FilePath,
    BestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    BestMatch = if Table.IsEmpty(SortedMatches) then LongestMatch
        else if List.Count(List.Distinct(BestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase)) = 1 then LongestMatch
        else error "CentriSyncPaths contains conflicting local folders for: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    FolderSegments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    UnitsIndex = List.PositionOf(FolderSegments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    CalcIndex = List.PositionOf(FolderSegments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    RootPath = if UnitsIndex >= 0
        and CalcIndex = UnitsIndex + 2
        and List.Count(FolderSegments) = CalcIndex + 1
        and Text.Trim(FolderSegments{UnitsIndex + 1}) <> "" then WorkbookFolder
        else error "FilePathUrl must identify Capacity.xlsx under UNITS\<unit>\2. Calculations. Save and recalculate its path input.",
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
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
// Purpose: Return the resolved calculation-folder root used by Capacity imports.
shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;
