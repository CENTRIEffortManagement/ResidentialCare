// Power Query from: Availability.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Availability.xlsx
// Extracted: 2026-09-10T09:03:17.781Z

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

shared #"FilePath-2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT AvailabilityDayShiftMATRIXRaw" = let

    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations" & "\Capacity-ShiftAvailability.xlsx"), null, true),
    #"Availability-StaffList_Sheet" = Source{[Item="Availability-StaffList",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(#"Availability-StaffList_Sheet",{{"Column1", type text}, {"Column2", type text}}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Changed Type", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Name", type text}, {"Role", type text}})
in
    #"Changed Type1";

shared #"Masterlist Join" = let
    Source = AvailableStaffList,
    #"Appended Query" = Table.Combine({Source, AllocationTable_StaffList}),
    #"Removed Columns" = Table.RemoveColumns(#"Appended Query",{"Source"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns", {"Name"})
in
    #"Removed Duplicates";

shared Masterlist = let
    Source = #"Masterlist Join",
    #"Added RESOURCE INDEX" = Table.AddIndexColumn(Source, "Resource", 1, 1, Int64.Type),
    #"Merged Queries" = Table.NestedJoin(#"Added RESOURCE INDEX", {"Name", "Role"}, #"Misaligned Join", {"Misaligned-Name", "Misaligned-Role"}, "Misaligned Join", JoinKind.FullOuter),
    #"Expanded Misaligned Join" = Table.ExpandTableColumn(#"Merged Queries", "Misaligned Join", {"Misalignment"}, {"Misalignment"})
in
    #"Expanded Misaligned Join";

shared AllocationTable_StaffList = let
     Source1 = #"IMPORT Table_AllocatedStaffList",
    #"Sorted Rows" = Table.Sort(Source1,{{"Name", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows", "Source", each "Allocation")
in
    #"Added Custom";

shared AvailableStaffListInitial = let
    Source = #"IMPORT AvailabilityDayShiftMATRIXRaw",
    #"Grouped Rows" = Table.Group(Source, {"Role", "Name"}, {{"AvailableListInitial", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"AvailableListInitial"})
in
    #"Removed Columns";

shared AvailableStaffList = let
     Source1 = AvailableStaffListInitial,
    #"Added Custom" = Table.AddColumn(Source1, "Source", each "Availability"),
    #"Duplicated Column" = Table.DuplicateColumn(#"Added Custom", "Name", "Name.A"),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Duplicated Column", "Name", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Name.1", "Name.2", "Name.3", "Name.4", "Name.5", "Name.6"}),
    #"Removed Columns" = Table.RemoveColumns(#"Split Column by Delimiter",{"Name.4", "Name.5", "Name.6"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns",null,"",Replacer.ReplaceValue,{"Name.2", "Name.3"}),
    #"Replaced NANDEEENI" = Table.ReplaceValue(#"Replaced Value","Muni","",Replacer.ReplaceText,{"Name.2"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced NANDEEENI",{"Name.2", "Name.3"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Merged"),
    #"Trimmed Text" = Table.TransformColumns(#"Merged Columns",{{"Merged", Text.Trim, type text}}),
    #"Merged Columns1" = Table.CombineColumns(#"Trimmed Text",{"Name.1", "Merged"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Name"),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Merged Columns1", "Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Name.1", "Name.2"}),
    #"Renamed Columns" = Table.RenameColumns(#"Split Column by Delimiter1",{{"Name.A", "Name"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Name", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Name.1", "Name.2"})
in
    #"Removed Columns1";

shared #"IMPORT Table_AllocatedStaffList" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-1Input"& "\1-AllocationExtracted.xlsx"), null, true),
    Table_AllocatedStaffList_Table = Source{[Item="Table_AllocatedStaffList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_AllocatedStaffList_Table,{{"Name", type text}, {"Role", type text}})
in
    #"Changed Type";

shared #"Misaligned Join" = let
    Source = Table.NestedJoin(AllocationTable_StaffList, {"Name", "Role"}, AvailableStaffList, {"Name", "Role"}, "Availability-StaffList", JoinKind.FullOuter),
    #"Expanded Availability-StaffList" = Table.ExpandTableColumn(Source, "Availability-StaffList", {"Name", "Role"}, {"Availability-StaffList.Name", "Availability-StaffList.Role"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Availability-StaffList",{{"Name", Order.Ascending}}),
    #"Added Conditional Column" = Table.AddColumn(#"Sorted Rows", "Custom", each if [Name] <> [#"Availability-StaffList.Name"] then "Misaligned" else null),
    #"Sorted Rows1" = Table.Sort(#"Added Conditional Column",{{"Availability-StaffList.Name", Order.Ascending}}),
    #"Filtered Rows1" = Table.SelectRows(#"Sorted Rows1", each ([Custom] = "Misaligned")),
    #"Inserted Merged Column" = Table.AddColumn(#"Filtered Rows1", "Misaligned-Name", each Text.Combine({[Name], [#"Availability-StaffList.Name"]}, ""), type text),
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Misaligned-Role", each Text.Combine({[Role], [#"Availability-StaffList.Role"]}, ""), type text),
    #"Added Conditional Column1" = Table.AddColumn(#"Inserted Merged Column1", "Misalignment", each if [Name] = null then "Available Only" else if [#"Availability-StaffList.Name"] = null then "Allocated Only" else "ERROR")
in
    #"Added Conditional Column1";

shared #"Misaligned Allocation Availability Names" = let
    Source = #"Misaligned Join",
    #"Merged Queries" = Table.NestedJoin(Source, {"Misaligned-Name", "Misaligned-Role"}, Masterlist, {"Name", "Role"}, "Masterlist", JoinKind.LeftOuter),
    #"Expanded Masterlist" = Table.ExpandTableColumn(#"Merged Queries", "Masterlist", {"Resource"}, {"Resource"}),
    #"Sorted Rows2" = Table.Sort(#"Expanded Masterlist",{{"Misaligned-Role", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows2",{"Custom", "Misaligned-Name", "Misaligned-Role", "Misalignment", "Resource"})
in
    #"Removed Other Columns";

shared #"FilePath-1Input" = let
    Source = UnitL1PathTABLE,
    #"Replaced Value" = Table.ReplaceValue(Source,"2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{1}[Value]
in
    Value;