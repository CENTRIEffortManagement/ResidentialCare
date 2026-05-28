// Power Query from: 1-AllocationExtracted.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\1. Input\1-AllocationExtracted.xlsx
// Extracted: 2026-05-21T00:46:30.569Z

section Section1;

shared RosteredDays = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", Int64.Type}}),
    #"Calculated Distinct Count" = List.NonNullCount(List.Distinct(#"Changed Type"[Date]))
in
    #"Calculated Distinct Count";

shared RosterStart = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Calculated Maximum" = List.Min(#"Removed Other Columns"[Date])
in
    #"Calculated Maximum";

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Code", Int64.Type}, {"Name", type text}, {"Date", type date}, {"Start", type time}, {"Finish", type time}, {"Break", Int64.Type}, {"Break Time", type time}, {"Hours", type number},  {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Period", type text}, {"Event", type text}, {"Function", type text}, {"Comments", type text}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type1", {"Date"}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Removed Errors", {{"Code", 9999}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Replaced Errors",{"Code", "Name", "Date", "Start", "Finish", "Break", "Break Time", "Hours", "Location", "Department", "Area", "Role"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Name] <> "AIN 01 E4" and [Name] <> "AIN AM1 Healthcare HQ" and [Name] <> "AIN AM2 Healthcare HQ" and [Name] <> "AIN AM3 Healthcare HQ" and [Name] <> "AIN AM4 Healthcare HQ" and [Name] <> "AIN AM5 Healthcare HQ" and [Name] <> "AIN AM6 Healthcare HQ" and [Name] <> "AIN NS1 Healthcare HQ" and [Name] <> "AIN PM1 Healthcare HQ" and [Name] <> "AIN PM2 Healthcare HQ" and [Name] <> "RN 01 E4"))
in
    #"Filtered Rows";

shared AllocationExtraction = let
    Source = #"Allocation Prepare",
    #"Renamed Columns2" = Table.RenameColumns(Source,{{"Finish", "End"}}),
    #"Filtered ROLES" = Table.SelectRows(#"Renamed Columns2", each ([Role] = "Assistant In Nursing" or [Role] = "Assistant in Nursing (Certificate IV)" or [Role] = "In Charge" or [Role] = "In Charge On Call" or [Role] = "Registered Nurse" )),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered ROLES",{{"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Role", each if Text.Contains([RoleX], "Registered") then "RN" else if Text.Contains([RoleX], "In Charge") then "RN" else if [RoleX] = "Assistant In Nursing" then "AIN" else if [RoleX] = "Assistant in Nursing (Certificate IV)" then "AINC4" else "??"),
    #"Added Custom" = Table.AddColumn(#"Added Conditional Column", "Unit", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RoleX"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([Name] <> "AIN 01 E4" and [Name] <> "AIN AM1 Healthcare HQ" and [Name] <> "AIN AM2 Healthcare HQ" and [Name] <> "AIN AM3 Healthcare HQ" and [Name] <> "AIN AM4 Healthcare HQ" and [Name] <> "AIN AM5 Healthcare HQ" and [Name] <> "AIN AM6 Healthcare HQ" and [Name] <> "AIN NS1 Healthcare HQ" and [Name] <> "AIN PM1 Healthcare HQ" and [Name] <> "AIN PM2 Healthcare HQ" and [Name] <> "RN 01 E4"))
in
    #"Filtered Rows";

shared AllocationExtracted = let
    Source = Table.NestedJoin(AllocationExtraction, {"Name"}, ResDoubleRoles, {"Name"}, "ResDoubleRoles", JoinKind.LeftOuter),
    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(Source, "ResDoubleRoles", {"MultiRole"}, {"MultiRole"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded ResDoubleRoles",{{"Name", "NameX"}}),
    #"Inserted MULTIROLENAMES" = Table.AddColumn(#"Renamed Columns", "Name", each if [MultiRole] <> null 
then
Text.Combine({[NameX], [Role]}, " (")&")"
else 
[NameX]),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted MULTIROLENAMES",{"NameX", "MultiRole"})
in
    #"Removed Columns";

shared DayAdjustment = let
    Source = Excel.CurrentWorkbook(){[Name="DayAdjustment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayAdjustment", Int64.Type}}),
    DayAdjustment1 = #"Changed Type"{0}[DayAdjustment]
in
    DayAdjustment1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\1. Incoming\241115 ASH Roster\Ashburn Working roster 7 to 20 October 24.xls"), null, true),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Start", type time}, {"Date", type date}})
in
    #"Changed Type";

shared #"04062023Carer" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\03. Incoming\230613 Maitland Data Emailed\Roster  - 22.5 -4.6.23xlsx.xlsx"), null, true),
    #"04062023_Sheet" = Source{[Item="04062023",Kind="Sheet"]}[Data],
    #"Filtered Rows" = Table.SelectRows(#"04062023_Sheet", each true)
in
    #"Filtered Rows";

shared #"RN Master 23-5-2022 (4)" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\03. Incoming\230613 Maitland Data Emailed\Roster  - 22.5 -4.6.23xlsx.xlsx"), null, true),
    #"RN Master 23-5-2022 (4)_Sheet" = Source{[Item="RN Master 23-5-2022 (4)",Kind="Sheet"]}[Data]
in
    #"RN Master 23-5-2022 (4)_Sheet";

shared AllocationExtractedCheck = let
    Source = AllocationExtraction,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom2" = Table.AddColumn(#"Changed Type1", "Subtraction", each if [End] < [Start]
then ([End]) + (1-[Start])
else  [End]-[Start]),
    #"Filtered Rows3" = Table.SelectRows(#"Added Custom2", each true),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows3", each true),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows2",{{"Subtraction", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Role"}, {{"Days", each List.Sum([Subtraction]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "AllocationTime", each [Days]*24/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "AllocatedEffort", each [AllocationTime]*0.9366)
in
    #"Added Custom1";

shared AllocatedStaffList = let
    Source = AllocationExtracted,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Role", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows";

shared ResDoubleRoles = let
    Source = AllocationExtraction,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Count] = 2)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "MultiRole", each "MultiRole")
in
    #"Added Custom";

shared ResRolesProportion = let
    Source = Table.NestedJoin(AllocationExtraction, {"Name"}, ResDoubleRoles, {"Name"}, "ResDoubleRoles", JoinKind.LeftOuter),
    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(Source, "ResDoubleRoles", {"Name"}, {"Name.1"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDoubleRoles", each ([Name.1] <> null)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Name", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Name", "Date", "Start", "Hours", "Role"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"Name", "Role"}, {{"Hours", each List.Sum([Hours]), type nullable number}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", List.Distinct(#"Grouped Rows"[Role]), "Role", "Hours", List.Sum),
    #"Inserted Addition" = Table.AddColumn(#"Pivoted Column", "Addition", each [AIN] + [AINC4], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Addition",{{"Addition", "Total"}}),
    #"Inserted Division" = Table.AddColumn(#"Renamed Columns", "AINC4HrsAvilPref", each [AINC4] / [Total], type number),
    #"Rounded Off" = Table.TransformColumns(#"Inserted Division",{{"AINC4HrsAvilPref", each Number.Round(_, 1), type number}})
in
    #"Rounded Off";