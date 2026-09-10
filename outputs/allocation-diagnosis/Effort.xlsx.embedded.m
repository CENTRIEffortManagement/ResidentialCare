section Section1;

shared ResShiftAllocation = let
    Source = Table.NestedJoin(#"IMPORT ResourceShiftAllocation", {"Name", "Role"}, #"IMPORT Table_Masterlist", {"Name", "Role"}, "IMPORT Table_Masterlist", JoinKind.LeftOuter),
    #"Expanded IMPORT Table_Masterlist" = Table.ExpandTableColumn(Source, "IMPORT Table_Masterlist", {"Name", "Resource"}, {"Name.1", "Resource"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded IMPORT Table_Masterlist",{"ShiftPeriod", "Role", "TimeDate", "ResShiftFTE", "Resource"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Other Columns", {"TimeDate", "ShiftPeriod"}, DateShiftPeriod, {"Date", "Shift"}, "DateShiftPeriod", JoinKind.LeftOuter),
    #"Expanded DateShiftPeriod" = Table.ExpandTableColumn(#"Merged Queries", "DateShiftPeriod", {"Period"}, {"Period"}),
    #"Added Custom" = Table.AddColumn(#"Expanded DateShiftPeriod", "Facility", each Unit),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Type", each "Allocation")
in
    #"Added Custom1";

shared RoleList = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCvJTitWJVnIOdlWKjQUA", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "RoleList"}}),
    RoleList1 = #"Renamed Columns"[RoleList]
in
    RoleList1;

shared #"IMPORT DemandCorrected" = let

    Source = Excel.Workbook(File.Contents(Unit1Path &  "\2. Calculations\Demand.xlsx"), null, true),
    ShiftDemandHCAverageANACC_Table = Source{[Item="ShiftDemandHCAverageANACC",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDemandHCAverageANACC_Table,{{"Facility", type any}})
in
    #"Changed Type";

shared #"IMPORT AvailabilityDeveloped" = let
    Source = Excel.Workbook(File.Contents(Unit1Path&"\2. Calculations\Capacity.xlsx"), null, true),
    AvailabilityDeveloped_Table = Source{[Item="AvailabilityDeveloped",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AvailabilityDeveloped_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"AvailabilityType", type text}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}, {"Facility", type text}})
in
    #"Changed Type";

shared #"IMPORT Allocation" = let
    
    Source = Excel.Workbook(File.Contents(Unit1Path & "\2. Calculations\Allocation.xlsx"), null, true),
    RoleShiftAllocation_Table = Source{[Item="Allocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RoleShiftAllocation_Table,{{"ShiftDate", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"ShiftDate", "Date"}})
in
    #"Renamed Columns";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents(Unit1Path&"\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared IMPORTMaxCapacity = let
    Source = Excel.Workbook(File.Contents(Unit1Path&"\2. Calculations\Settings Data.xlsx"), null, true),
    MaxCapacity_Table = Source{[Item="MaxCapacity",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(MaxCapacity_Table,{{"MaxCapacityFactor", Int64.Type}}),
    MaxCapacityFactor = #"Changed Type"{0}[MaxCapacityFactor]
in
    MaxCapacityFactor;

shared RoleShiftAllocationDay = let
    Source = #"IMPORT Allocation",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ShiftPeriod", type text}, {"Role", type text}})
in
    #"Changed Type";

shared RoleDATEShiftDemand = let
    Source = #"IMPORT DemandCorrected"
in
    Source;

shared RoleShiftCapacity = let
    Source = #"IMPORT AvailabilityDeveloped",
    #"Filtered Rows" = Table.SelectRows(Source, each ([AvailabilityType] = "C###")),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"AvailabilityType", "Capacity"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns1", {"Period"}, DateShiftPeriod, {"Period"}, "IMPORT PermutationDimensions (2)", JoinKind.LeftOuter),
    #"Expanded PERMUTATIONS DATE SHIFT" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions (2)", {"Date", "Shift"}, {"Date", "Shift"}),
    #"Grouped CAPACITY CAPCAITY MAX" = Table.Group(#"Expanded PERMUTATIONS DATE SHIFT", {"Role", "Date", "Shift"}, {{"RoleShiftCapacity", each List.Sum([Availability]), type nullable number}, {"RoleMaxCapacity", each List.Max([ResMaxAvail]), type nullable number}}),
    #"Inserted Day Name" = Table.AddColumn(#"Grouped CAPACITY CAPCAITY MAX", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Inserted First Characters" = Table.AddColumn(#"Inserted Day Name", "First Characters", each Text.Start([Day Name], 3), type text),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted First Characters",{{"First Characters", "Day"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Day Name"}),
    #"Added CAPACITYX" = Table.AddColumn(#"Removed Columns", "CapacityX", each [RoleShiftCapacity] * IMPORTMaxCapacity),
    #"Sorted Rows" = Table.Sort(#"Added CAPACITYX",{{"Date", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared #"EffortAllMatrixAG1-1D" = let
    Source = Table.NestedJoin(RoleShiftCapacity, {"Date", "Role", "Shift"}, RoleDATEShiftDemand, {"Date", "Role", "ShiftPeriod"}, "Demand", JoinKind.FullOuter),
    #"Expanded DEMAND" = Table.ExpandTableColumn(Source, "Demand", {"Date", "Role", "ShiftDemandHCAverage", "ShiftPeriod"}, {"Date.1", "Role.1", "ShiftDemandHCAverage", "ShiftPeriod"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded DEMAND",{{"Date.1", type date}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type",{{"Date", "Date-Capacity"}, {"Role", "Role-Capacity"}, {"Role.1", "Role"}, {"Shift", "Shift-Capacity"}, {"Date.1", "Date"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns2",{"Role-Capacity", "Date-Capacity", "Shift-Capacity"}),
    #"Merged ALLOCATION" = Table.NestedJoin(#"Removed Columns1", {"Role", "Date", "ShiftPeriod"}, #"IMPORT Allocation", {"Role", "Date", "ShiftPeriod"}, "RoleShiftAllocation", JoinKind.LeftOuter),
    #"Expanded RoleShiftAllocation" = Table.ExpandTableColumn(#"Merged ALLOCATION", "RoleShiftAllocation", {"RoleShiftFTE"}, {"RoleShiftFTE"}),
    #"Sorted Rows" = Table.Sort(#"Expanded RoleShiftAllocation",{{"Date", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows",{{"RoleShiftCapacity", "Capacity"}, {"ShiftDemandHCAverage", "Demand"}, {"RoleMaxCapacity", "CapacityMaxHC"}, {"RoleShiftFTE", "Allocation"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Role", "Date", "ShiftPeriod", "Day", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation"}),
    #"Replaced Value" = Table.ReplaceValue(#"Reordered Columns",null,0,Replacer.ReplaceValue,{"Allocation"}),
    #"Removed Columns" = Table.RemoveColumns(#"Replaced Value",{"Day"}),
    #"Inserted Merged Column" = Table.AddColumn(#"Removed Columns", "Merged", each Text.Combine({Text.From([Date], "en-AU"), [ShiftPeriod]}, ""), type text),
    #"Filtered Rows" = Table.SelectRows(#"Inserted Merged Column", each ([Role] <> "Endorsed Enrolled Nurse")),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"Merged", "DATESHIFT"}, {"ShiftPeriod", "Shift"}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Renamed Columns1",null,0,Replacer.ReplaceValue,{"Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation"}),
    #"Added Facility" = Table.AddColumn(#"Replaced Value1", "Facility", each Unit),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Facility",{"Facility", "Role", "Date", "Shift", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation", "DATESHIFT"}),
    #"Sorted Rows1" = Table.Sort(#"Reordered Columns1",{{"Role", Order.Ascending}, {"DATESHIFT", Order.Ascending}})
in
    #"Sorted Rows1";

shared #"EffortDAMatrixAG1-1D" = let
    Source = Table.NestedJoin(RoleShiftAllocationDay, {"Date", "ShiftPeriod", "Role"}, RoleDATEShiftDemand, {"Date", "ShiftPeriod", "Role"}, "RoleDATEShiftDemand", JoinKind.LeftOuter),
    #"Expanded RoleDATEShiftDemand" = Table.ExpandTableColumn(Source, "RoleDATEShiftDemand", {"ShiftDemandHCAverage"}, {"ShiftDemandHCAverage"})
in
    #"Expanded RoleDATEShiftDemand";

shared RoleShiftAvailabilities = let
    Source = #"IMPORT AvailabilityDeveloped",
    #"Renamed Columns1" = Table.RenameColumns(Source,{{"AvailabilityType", "Capacity"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns1", {"Period"}, DateShiftPeriod, {"Period"}, "IMPORT PermutationDimensions (2)", JoinKind.LeftOuter),
    #"Expanded PERMUTATIONS DATE SHIFT" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions (2)", {"Date", "Shift"}, {"Date", "Shift"})
in
    #"Expanded PERMUTATIONS DATE SHIFT";

shared DateShiftPeriod = let
    Source = #"IMPORT PermutationDimensions",
    #"Removed Columns" = Table.RemoveColumns(Source,{"RolesList", "Day"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Duplicates",{{"Shifts", "Shift"}})
in
    #"Renamed Columns";

shared #"IMPORT ResourceShiftAllocation" = let
    
    Source = Excel.Workbook(File.Contents(Unit1Path & "\2. Calculations\Allocation.xlsx"), null, true),
    ResourceShiftAllocation_Table = Source{[Item="ResourceShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResourceShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Name", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Role", Order.Ascending}, {"Name", Order.Ascending}, {"TimeDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared #"IMPORT Table_Masterlist" = let
    Source = Excel.Workbook(File.Contents(Unit1Path&"\2. Calculations\StaffListMaster.xlsx"), null, true),
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_Masterlist_Table,{ {"Name", type text}, {"Role", type text},  {"Resource", Int64.Type}})
in
    #"Changed Type";

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
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","\2. Calculations","",Replacer.ReplaceText,{"Folder"}),
    Folder = #"Replaced Value"{0}[Folder]
in
    Folder;

shared Unit = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared AvailabilityDeveloped2 = let
    Source = Excel.Workbook(File.Contents(Unit1Path&"\2. Calculations\Capacity.xlsx"), null, true),
    AvailabilityDeveloped_Table = Source{[Item="AvailabilityDeveloped",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AvailabilityDeveloped_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"AvailabilityType", type text}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}, {"Facility", type text}})
in
    #"Changed Type";

shared CentriSyncPaths = let
    Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(CentriSyncPaths_Table,{{"User", type text}, {"SharepointRootUrl", type text}, {"Site", type text}, {"SyncedFolderRootPath", type text}})
in
    #"Changed Type";