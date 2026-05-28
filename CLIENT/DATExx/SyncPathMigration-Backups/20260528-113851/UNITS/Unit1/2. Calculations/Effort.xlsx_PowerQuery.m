// Power Query from: Effort.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Effort.xlsx
// Extracted: 2026-05-18T06:14:21.038Z

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

shared UnitL1PathTABLE = // Version 25.00    250107 
// Updated 250107

let

    // PART A - Define FilePathUrl and Replaced Value
    FilePathUrl = 
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content],
        #"Renamed Columns" = Table.RenameColumns(Source, {{"Column1", "FilePath"}}),
        ReplacedValue = Table.ReplaceValue(#"Renamed Columns", "/", "\", Replacer.ReplaceText, {"FilePath"}),
        BufferedTable = Table.Buffer(ReplacedValue) // Buffer the table for better performance
    in 
        BufferedTable,

    // PART B - RootPath
    IMPORTRootPath =
    let
        // Step 1: Import CentriSyncPaths
        CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
        CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
        CentriSyncPaths_ChangedType = Table.TransformColumnTypes(CentriSyncPaths_Table, {{"User", type text}, {"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),

        // Step 2: UrlSite - Use buffered FilePathUrl
        UrlSite_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "sites\"), type text}}),
        UrlSite_ExtractedTextBeforeDelimiter = Table.TransformColumns(UrlSite_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\"), type text}}),
        RenamedColumns1 = Table.RenameColumns(UrlSite_ExtractedTextBeforeDelimiter, {{"FilePath", "Site"}}),

        // Step 3: Prefix
        Prefix_NestedJoin = Table.NestedJoin(RenamedColumns1, {"Site"}, CentriSyncPaths_ChangedType, {"Site"}, "CentriSyncPaths", JoinKind.LeftOuter),
        Prefix_Expanded = Table.ExpandTableColumn(Prefix_NestedJoin, "CentriSyncPaths", {"SyncedFolderRootPath"}, {"Prefix"}),
        Prefix = Prefix_Expanded{0}[Prefix],

        // Step 4: Core
        Core_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "Shared Documents\"), type text}}),
        Core_ExtractedTextBeforeDelimiter = Table.TransformColumns(Core_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
        Core = Core_ExtractedTextBeforeDelimiter{0}[FilePath],

        // Step 5: FilePath
        FilePath = Prefix & "\" & Core,
        ConvertedToTable = #table(1, {{FilePath}}),
        RenamedColumns = Table.RenameColumns(ConvertedToTable, {{"Column1", "FilePath"}})
    in  
        RenamedColumns,

    // PART C - Dimensions
    // Extract UserName directly
    UserName = 
    let
        Source = IMPORTRootPath,
        Extracted = Text.BeforeDelimiter(Text.AfterDelimiter(Source[FilePath]{0}, "\", 1), "\")
    in
        Extracted,

    // Extract Client directly
    Client = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{3, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,

    // Extract Unit directly
    Unit = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{0, RelativePosition.FromEnd}), 
                    "\")
    in
        Extracted,

    // Extract Date directly
    Date = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{2, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,

    // Extract Filename
    FileName = 
    let
        Source = FilePathUrl,
        #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"FilePath", each Text.AfterDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
        #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"FilePath", each Text.BeforeDelimiter(_, "]"), type text}}),
        #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"FilePath"}),
        String = #"Replaced Value"{0}[FilePath]
    in 
        String,

    // PART D - Table
    // Create the table with variable names and their corresponding values
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", IMPORTRootPath{0}[FilePath]},
            {"FilePathUrl", FilePathUrl{0}[FilePath]},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    )
in
    TABLE;

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
    Value = Source{5}[Value]
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