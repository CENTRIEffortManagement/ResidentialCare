// Power Query from: Effort-All.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\E-O-I\Effort-All.xlsx
// Extracted: 2026-05-21T00:47:25.835Z

section Section1;

shared EffectiveAvailability = 7.6/8 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared DateAlignment = let
    Source = Excel.CurrentWorkbook(){[Name="DateAlignment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateAlignment", Int64.Type}}),
    DateAlignment1 = #"Changed Type"{0}[DateAlignment]
in
    DateAlignment1;

shared Folder = let
    Source = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Folder", type text}}),
    Folder1 = #"Changed Type"{0}[Folder]
in
    Folder1;

shared Facilities = let
    Source = Excel.CurrentWorkbook(){[Name="Table3"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Facility1", type text}, {"Facility2", type text}})
in
    #"Changed Type";

shared Facility1 = let
    Source = Facilities,
    Facility1 = Source{0}[Facility1]
in
    Facility1;

shared Facility2 = let
    Source = Facilities,
    Facility1 = Source{0}[Facility2]
in
    Facility1;

shared #"Facility1-EffortAllMatrixAG1_1D" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility1 & "\2. Calculations\Effort.xlsx"), null, true),
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text},  {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

shared #"Facility1 - Availabilities" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility1 & "\2. Calculations\Effort.xlsx"), null, true),
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text}, {"Facility", type text}, {"Date", type date}, {"Shift", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared #"Facility1 - ResAllocation" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility1 & "\2. Calculations\Effort.xlsx"), null, true),
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type1",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}})
in
    #"Changed Type";

shared #"Facility2-EffortAllMatrixAG1_1D" = let
     Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder & Facility2&  "\2. Calculations\Effort.xlsx"), null, true),
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Renamed Columns" = Table.RenameColumns(EffortAllMatrixAG1_1D_Table,{{"Date", "Datetemp"}}),
    #"Added ALIGNWITHBERR" = Table.AddColumn(#"Renamed Columns", "Custom", each Date.AddDays([Datetemp],DateAlignment)),
    #"Renamed Columns1" = Table.RenameColumns(#"Added ALIGNWITHBERR",{{"Custom", "Date"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns1",{"Facility", "Role", "Date", "Datetemp", "Shift", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation", "DATESHIFT"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Datetemp"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "ShiftDurations.ShiftDuration", each 1),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Date", type date}})
in
    #"Changed Type";

shared #"Facility2 - Availabilities ##" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility2 & "\2. Calculations\Effort.xlsx"), null, true),
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text},  {"Date", type date}, {"Shift", type text}})
in
    #"Changed Type1";

shared #"Facility2 - ResAllocation" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility2 & "\2. Calculations\Effort.xlsx"), null, true),
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}})
in
    #"Changed Type1";

shared Table_RevisedAllocated = let

    
    Source = Excel.Workbook(File.Contents(Folder & "\2. Calculations\Change\AllocationChange.xlsx"), null, true),
    Table_RevisedAllocated_Table = Source{[Item="Table_RevisedAllocated",Kind="Table"]}[Data],
    #"Added Custom" = Table.AddColumn(Table_RevisedAllocated_Table, "Version", each "ChangeAllcoation")
in
    #"Added Custom";

shared ResShiftAllocation = let
    Source1 = #"Facility2 - ResAllocation",
    #"Appended Query" = Table.Combine({Source1, #"Facility1 - ResAllocation"}),
    //CombinedTables = Table.Combine({Source1, #"Facility1 - ResAllocation"}),
    #"Renamed Columns" = Table.RenameColumns(#"Appended Query",{{"ShiftPeriod", "Shift"}, {"TimeDate", "Date"}, {"ResShiftFTE", "Effort"}, {"Type", "EffortType"}})
in
    #"Renamed Columns";

shared #"EffortAllMatrixAG1_1D-base !!" = let
   Source = #"Facility2-EffortAllMatrixAG1_1D",
    #"Appended Query" = Table.Combine({Source, #"Facility1-EffortAllMatrixAG1_1D"}),
    #"Multiplied Column" = Table.TransformColumns(#"Appended Query", {{"Capacity", each _ * EffectiveAvailability, type number}}),
    #"Multiplied Column1" = Table.TransformColumns(#"Multiplied Column", {{"CapacityX", each _ * EffectiveAvailability, type number}}),
    //#"Appended Query" = Table.Combine({Source, #"Facility1-EffortAllMatrixAG1_1D"}),
    #"Renamed Columns" = Table.RenameColumns(#"Multiplied Column1",{{"Date", "DateX"}}),  // Add back Facility 2
    #"Added DATE ALIGNMENT" = Table.AddColumn(#"Renamed Columns", "Date", each if [Facility] = "ASHB" then  Date.AddDays([DateX],-DateAlignment)
else [DateX]),
    #"Removed Columns" = Table.RemoveColumns(#"Added DATE ALIGNMENT",{"DateX"})
in
    #"Removed Columns";

shared #"EffortAllMatrixAG1_1D-Revised" = let
    Source = #table({}, {})
    
   /* let
    Source = #"EffortAllMatrixAG1_1D-base !!",
    #"Merged Queries" = Table.NestedJoin(Source, {"Facility", "Role", "Date", "Shift"}, Table_RevisedAllocated, {"Facility", "Role", "Date", "ShiftPeriod"}, "Table_RevisedAllocated", JoinKind.FullOuter),
    #"Expanded Table_RevisedAllocated1" = Table.ExpandTableColumn(#"Merged Queries", "Table_RevisedAllocated", {"RevisedAllocation", "Version"}, {"Table_RevisedAllocated.RevisedAllocation", "Table_RevisedAllocated.Version"}),
    #"Added ALLOCATIONBASE+CHANGE" = Table.AddColumn(#"Expanded Table_RevisedAllocated1", "Allocation.1", each if [Table_RevisedAllocated.RevisedAllocation]= null 
then [Allocation]
else [Table_RevisedAllocated.RevisedAllocation]),
    #"Removed Columns" = Table.RemoveColumns(#"Added ALLOCATIONBASE+CHANGE",{"Allocation", "Table_RevisedAllocated.RevisedAllocation", "Table_RevisedAllocated.Version"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Allocation.1", "Allocation"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns1",{"Version"}),
    #"Added Custom1" = Table.AddColumn(#"Removed Columns1", "Version", each "AllocationChange"),
    #"Filtered Rows1" = Table.SelectRows(#"Added Custom1", each ([Facility] <> null))
in
    #"Filtered Rows1" 
    */

    in
    Source;

shared EffortAllMatrixAG1_1D = let
    Source = #"EffortAllMatrixAG1_1D-base !!",
    // = Table.Combine({#"EffortAllMatrixAG1_1D-base !!", #"EffortAllMatrixAG1_1D-Revised"}),
    #"Sorted Rows1" = Table.Sort(Source,{{"DATESHIFT", Order.Ascending}}),
    #"Sorted Rows" = Table.Sort(#"Sorted Rows1",{{"DATESHIFT", Order.Ascending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each ([Role] <> null)),
    #"Sorted Rows2" = Table.Sort(#"Filtered Rows",{{"Facility", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows2",{{"Date", type date}}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Changed Type", "Date", "Date - Copy"),
    #"Renamed Columns" = Table.RenameColumns(#"Duplicated Column",{{"Date - Copy", "DateX"}})
in
    #"Renamed Columns";

shared #"Availabilities Appended" = let
    Source = #"Facility2 - Availabilities ##",
    #"Appended FAC1" = Table.Combine({Source, #"Facility1 - Availabilities"}),
    //Table.Combine({#"Facility1 - Availabilities", #"Facility2 - Availabilities"}),    //return Facility 1
    #"Renamed Columns" = Table.RenameColumns(#"Appended FAC1",{{"Capacity", "EffortType"}, {"Availability", "Effort"}}),
    #"Multiplied Column" = Table.TransformColumns(#"Renamed Columns", {{"Effort", each _ * EffectiveAvailability, type number}})
in
    #"Multiplied Column";

shared #"Availability Effort" = let
    Source = #"EffortAllMatrixAG1_1D-base !!",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Facility", "Role", "Shift", "Demand", "Date"}),
    #"Added DEMAND" = Table.AddColumn(#"Removed Other Columns", "EffortType", each "Demand"),
    #"Renamed Columns" = Table.RenameColumns(#"Added DEMAND",{{"Demand", "Effort"}}),
    #"Appended AVAILABILITIES" = Table.Combine({#"Renamed Columns", #"Availabilities Appended"}),
    #"Removed Columns" = Table.RemoveColumns(#"Appended AVAILABILITIES",{"ResAvailability", "ResMaxAvail"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text}, {"EffortType", type text}, {"Effort", type number}, {"Resource", Int64.Type}}),
    #"Appended ALLOCATION" = Table.Combine({#"Changed Type", ResShiftAllocation}),
    #"Replaced Value" = Table.ReplaceValue(#"Appended ALLOCATION",0,null,Replacer.ReplaceValue,{"Effort"}),
    #"Sorted Rows" = Table.Sort(#"Replaced Value",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared ResRoleAvailabilityDevelopedMATRIX = let
    Source = #"Availability Effort",
    #"Sorted Rows2" = Table.Sort(Source,{{"Date", Order.Ascending}, {"EffortType", Order.Ascending}}),
    #"Sorted Rows" = Table.Sort(#"Sorted Rows2",{{"Effort", Order.Descending}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Sorted Rows",{{"Effort", type number}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type1",{{"Resource", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"EffortType", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(#"Sorted Rows1", List.Distinct(#"Sorted Rows1"[EffortType]), "EffortType", "Effort")
in
    #"Pivoted Column";

shared #"Capacity SUM" = let
    Source = EffortAllMatrixAG1_1D,
    //#"Filtered Rows" = Table.SelectRows(Source, each ([Version] = "Base")),
    #"Grouped Rows" = Table.Group(#"Source", {"Facility", "Role"}, {{"Capacity", each List.Sum([Capacity]), type number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped Rows",null,"?",Replacer.ReplaceValue,{"Role"}),
    Custom1 = List.Sum(#"Replaced Value"[Capacity]),
    #"Converted to Table" = #table(1, {{Custom1}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Converted to Table",{{"Column1", "TOTAL"}}),
    TOTAL = #"Renamed Columns1"{0}[TOTAL]
in
    TOTAL;

shared ResShiftAllocation1 = let
    Source = Excel.Workbook(File.Contents(Folder & Facility2 & "\2. Calculations\Effort.xlsx"), null, true),
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"TimeDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared RoleAvailabilityDevelopedMATRIXDELTA = let
    Source = ResRoleAvailabilityDevelopedMATRIX,
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role", "Shift", "Date"}, {{"Demand", each List.Sum([Demand]), type nullable number}, {"C###", each List.Sum([#"C###"]), type nullable number}, {"Allocation", each List.Sum([Allocation]), type nullable number}, {"Original Availability", each List.Sum([OriginalAvailability]), type nullable number}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Grouped Rows", "C###-D", each [#"C###"] - [Demand], type number),
    #"Inserted Addition" = Table.AddColumn(#"Inserted Subtraction", "A-D", each [Allocation] - [Demand], type number),
    #"Inserted Subtraction1" = Table.AddColumn(#"Inserted Addition", "A-C###", each [#"C###"] - [Allocation], type number),
    #"Inserted Subtraction2" = Table.AddColumn(#"Inserted Subtraction1", "CO-D", each [Original Availability] - [Demand], type number)
in
    #"Inserted Subtraction2";

shared MaxAvailabilities = let
    Source = #"Availabilities Appended",
    #"Filtered Rows" = Table.SelectRows(Source, each ([EffortType] = "C###")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Role", "Resource", "Effort", "Facility", "ResAvailability", "ResMaxAvail"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Effort", "C###"}})
in
    #"Renamed Columns";

shared UnitL1PathTABLE = let
  
    
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
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;