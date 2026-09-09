// Power Query from: Settings Data.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx
// Extracted: 2026-09-08T23:22:31.834Z

section Section1;

shared #"Dates Listed" = let
    Source = DateRoleShiftAllocation,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Allocation"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Date", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared #"Min Date" = let
    Source = #"Dates Listed",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] >= DateFrom)),
    Custom1 = Table.Min(#"Filtered Rows","Date"),
    Date = Custom1[Date]
in
    Date;

shared DateList = let
    Source = #"Dates Listed",
    #"Changed Type4" = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Custom1 = #"Changed Type4"[Date],
    Custom2 = List.Max(#"Changed Type4" [Date]),
    #"Converted to Table" = #table(1, {{Custom2}}),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Max"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Min", each #"Min Date"),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom1",{{"Max", type date}, {"Min", type date}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"Max", Int64.Type}, {"Min", Int64.Type}}),
    #"Added Custom3" = Table.AddColumn(#"Changed Type2", "Custom", each {[Min]..[Max]}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom3",{{"Max", type date}, {"Min", type date}}),
    #"Expanded Custom" = Table.ExpandListColumn(#"Changed Type1", "Custom"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Custom",{"Max", "Min"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Custom", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"Custom", "Date"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared PermutationDimensions = let
    Source = #"DateList",
    #"Added SHIFTS" = Table.AddColumn(Source, "Shifts", each Shifts),
    #"Expanded Shifts" = Table.ExpandListColumn(#"Added SHIFTS", "Shifts"),
    #"Added Index" = Table.AddIndexColumn(#"Expanded Shifts", "Period", 1, 1, Int64.Type),
    #"ADD ROLES" = Table.AddColumn(#"Added Index", "RolesList", each Roles),
    #"Expanded RolesList" = Table.ExpandListColumn(#"ADD ROLES", "RolesList"),
    #"Sorted Rows" = Table.Sort(#"Expanded RolesList",{{"Date", Order.Ascending}, {"Period", Order.Ascending}, {"RolesList", Order.Ascending}})
in
    #"Sorted Rows";

shared DateNameRoleShiftAllocation = let
    Source = AllocationExtracted,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Custom", each (if [End]<[Start]
then (1-[Start])+[End]
else [End]-[Start])*24),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each ([End] <> null)),
    #"Added Custom2" = Table.AddColumn(#"Filtered Rows", "TimeWorked", each if [Custom] >6 then [Custom] - .5
else [Custom]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom2", "Shift", each if [Start] < 0.604 then "AM" else if [Start] > 0.89 then "NIGHT" else "PM"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom1",{ "Start", "End", "Custom"})
in
    #"Removed Columns1";

shared DateRoleShiftAllocation = let
    Source = DateNameRoleShiftAllocation,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Role", "Shift"}, {{"Allocation", each List.Sum([TimeWorked]), type number}})
in
    #"Grouped Rows";

shared AllocationExtracted = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Date", type date}, {"Start", type time}, {"End", type time}, {"Break", Int64.Type}, {"Hours", type number}, {"Name", type text}, {"Code", Int64.Type}, {"Role", type text}})
in
    #"Changed Type";

shared Roles = let
    Source = Excel.CurrentWorkbook(){[Name="Roles"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ROLES", type text}}),
    ROLES = #"Changed Type"[ROLES]
in
    ROLES;

shared Shifts = let
    Source = Excel.CurrentWorkbook(){[Name="Shifts"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}}),
    #"Added Index1" = Table.AddIndexColumn(#"Changed Type", "Index.1", 1, 1, Int64.Type),
    Shifts1 = #"Added Index1"[Shifts]
in
    Shifts1;

shared DateFrom = let
    Source = Excel.CurrentWorkbook(){[Name="DateFrom"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateFrom", type date}}),
    DateFrom1 = #"Changed Type"{0}[DateFrom]
in
    DateFrom1;