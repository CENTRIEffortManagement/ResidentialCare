// Power Query from: AllocationChange.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\Change\AllocationChange.xlsx
// Extracted: 2026-05-21T00:48:07.522Z

section Section1;

shared #"Changed Allocation" = let
    Source = Table_ChangesAllRoles,
    #"Added VERSION" = Table.AddColumn(Source, "Version", each "ChangedAllocation"),
    #"Divided FTE" = Table.TransformColumns(#"Added VERSION", {{"Value", each _ / ShiftDuration, type number}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Divided FTE",{{"Value", "ChangedAllocFTE"}})
in
    #"Renamed Columns1";

shared RevisedAllocated = let
    Source = Table.NestedJoin(Allocation, {"Facility", "Role", "Date", "ShiftPeriod"}, #"Changed Allocation", {"Facility", "Role", "Date", "Shift"}, "Changed Allocation", JoinKind.FullOuter),
    #"Expanded Changed Allocation1" = Table.ExpandTableColumn(Source, "Changed Allocation", {"Role", "ChangedAllocFTE", "Shift", "Date", "Facility"}, {"Role.1", "ChangedAllocFTE", "Shift", "Date.1", "Facility.1"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded Changed Allocation1",null,0,Replacer.ReplaceValue,{"Allocation", "ChangedAllocFTE"}),
    #"Inserted Addition" = Table.AddColumn(#"Replaced Value", "RevisedAllocation", each [ChangedAllocFTE] + [Allocation], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Addition",{"Allocation", "Version" }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Role", "RoleX"}, {"ShiftPeriod", "ShiftPeriodX"}, {"Date", "DateX"}, {"Facility", "FacilityX"}}),
    #"Combine ROLE" = Table.AddColumn(#"Renamed Columns", "Role", each if [RoleX] = null then [Role.1]
else [RoleX]),
    #"Combine DATE" = Table.AddColumn(#"Combine ROLE", "Date", each if [Date.1] = null then [DateX] else [Date.1]),
    #"Combine SHIFTS" = Table.AddColumn(#"Combine DATE", "ShiftPeriod", each if [ShiftPeriodX] = null then [Shift]
else [ShiftPeriodX]),
    #"Added Custom" = Table.AddColumn(#"Combine SHIFTS", "Facility", each if [FacilityX] = null then [Facility.1] else [FacilityX]),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"ShiftPeriodX", "RoleX", "DateX", "Role.1", "ChangedAllocFTE", "Shift", "Date.1", "Facility.1", "FacilityX"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns1",{{"Role", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"Date", type date}})
in
    #"Changed Type";

shared Table_ChangesAllRoles = let
     Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder & "\MAIT\2. Calculations\Changes.xlsm"), null, true),
    ChangesAllRoles = Source{[Item="ChangesAllRoles",Kind="Table"]}[Data],
    #"Replaced Value" = Table.ReplaceValue(ChangesAllRoles,"RN","Registered Nurse",Replacer.ReplaceText,{"Role"})
in
    #"Replaced Value";

shared Shifts = let
         Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder & "\2. Calculations\SETTINGS.xlsx"), null, true),
    Shifts_Table = Source{[Item="Shifts",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Shifts_Table,{{"Shifts", type text}}),
    #"Added Index" = Table.AddIndexColumn(#"Changed Type", "Index", 1, 1, Int64.Type)
in
    #"Added Index";

shared Allocation = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\MAIT\2. Calculations\Allocation.xlsx"), null, true),
    Table_Allocation_Table = Source{[Item="Table_Allocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_Allocation_Table,{{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"Role", type text}, {"RoleShiftEffort", type number}, {"Allocation", type number}}),
    #"Added FACILITY" = Table.AddColumn(#"Changed Type", "Facility", each "MAIT"),
    #"Added Version" = Table.AddColumn(#"Added FACILITY", "Version", each "Baseline"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Version",{"RoleShiftEffort"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Allocation", "Allocation"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([ShiftDate],DateSHIFT)),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"ShiftDate"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns1",{{"Role", Order.Ascending}, {"Date", Order.Ascending}})
in
    #"Sorted Rows";

shared ShiftDuration = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\MAIT\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared DateSHIFT = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMrJQio0FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", Int64.Type}}),
    Column1 = #"Changed Type"{0}[Column1]
in
    Column1;