// Power Query from: EffortOutcomeLogXY.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\EOW\EffortOutcomeLogXY.xlsx
// Extracted: 2026-05-21T00:47:42.393Z

section Section1;

shared Folder = let
    Source = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Folder", type text}}),
    Folder1 = #"Changed Type"{0}[Folder]
in
    Folder1;

shared EffortOutcomesAG0 = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\E-O-I\EffortOutcomes.xlsx"), null, true),
    EffortOutcomesAG0_Table = Source{[Item="EffortOutcomesAG0",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortOutcomesAG0_Table,{{"Date", type date}, {"WARDS", type text}, {"Role", type text}, {"Shifts", type text}, {"Allocation", type number}, {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}, {"Epn", type number}, {"EpX", type number}, {"Ein", type number}, {"Ipn", type any}, {"Ipx", type number}, {"Iin", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"CapacityX"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "Point"}})
in
    #"Renamed Columns";

shared #"EffortOutcomesAG1-1DS" = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\E-O-I\EffortOutcomes.xlsx"), null, true),
    EffortOutcomesAG1_1DayShift_Table = Source{[Item="EffortOutcomesAG1_1DayShift",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(EffortOutcomesAG1_1DayShift_Table,{{"Facility", type text}, {"Role", type text}, {"Shift", type text}, {"Demand", type number}, {"Capacity", type number}, {"CapacityMaxHC", Int64.Type}, {"CapacityX", type number}, {"Allocation", type number}, {"DATESHIFT", type text}, {"ShiftDurations.ShiftDuration", Int64.Type}, {"Date", type date}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}, {"Epn", type number}, {"EpX", type number}, {"Ein", type number}, {"Ipn", type number}, {"Ipx", type number}, {"Iin", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"CapacityX"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "Point"}})
in
    #"Renamed Columns";

shared #"Null Matrix Check" = let
    Source = EffortOutcomesAG0,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date", "WARDS", "Role", "Shifts", "Demand"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[Role]), "Role", "Demand", List.Sum)
in
    #"Pivoted Column";

shared EffortOutcomesAG1_1DayShift = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\2. Calculations\E-O-I\EffortOutcomes.xlsx"), null, true),
    EffortOutcomesAG1_1DayShift_Table = Source{[Item="EffortOutcomesAG1_1DayShift",Kind="Table"]}[Data]
in
    EffortOutcomesAG1_1DayShift_Table;