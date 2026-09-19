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

shared UnitL1PathTABLE = // Version 25.01 local-first ResidentialCare
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
    WorkbookFolder = Text.BeforeDelimiter(RawFilePath, "\", {0, RelativePosition.FromEnd}),
    IsLocalPath = (Text.Contains(WorkbookFolder, ":\") or Text.StartsWith(WorkbookFolder, "\\")) and not Text.StartsWith(Text.Lower(WorkbookFolder), "http"),
    RootPath = if IsLocalPath then WorkbookFolder else error "FilePathUrl did not resolve to a local path. Open/save this workbook from the active local sync folder before refreshing.",
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(RawFilePath, "[", "]") otherwise Text.AfterDelimiter(RawFilePath, "\", {0, RelativePosition.FromEnd}),
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"FilePathUrl", RawFilePath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;


