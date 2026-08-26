// Power Query from: Demand-MasterRoster Manual Read.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\1. Input\Demand-MasterRoster Manual Read.xlsx
// Extracted: 2026-08-26T09:55:30.984Z

section Section1;

shared #"IMPORT Master" = let
    Source = Excel.Workbook(File.Contents("C:\Users\alexp\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\ResidentialCare\Whiddon\DATExx\UNITS\Unit1\1. Input\Master Roster.xlsx"), null, true),
    Combined_Sheet = Source{[Item="Combined",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(Combined_Sheet, [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Master Template", type text}, {"Template", type text}, {"Location", type text}, {"Department", type text}, {"Role", type text}, {"Area", type text}, {"Employee Code", Int64.Type}, {"Employee Name", type text}, {"Week No", Int64.Type}, {"Week Day", type text}, {"Start Time", type time}, {"End Time", type time}, {"Roster Hours", type number}, {"Cost", type number}, {"MinRosterHours", type number}, {"MaxRosterHours", Int64.Type}, {"Event", type text}, {"Break Length", Int64.Type}, {"Break Start Time", type time}, {"Paid Break Length", Int64.Type}, {"Paid Break Start Time", type time}, {"Shift Definition", type text}, {"Shift Net Length", type number}, {"Shift Type", type text}, {"Non Attended", type logical}})
in
    #"Changed Type1";

shared #"INPUT SHiftEnd" = let
    Source = Excel.CurrentWorkbook(){[Name="Table7"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shift", type text}, {"End", type time}})
in
    #"Changed Type";

shared #"ShftEnd Prepare" = let
    Source = #"INPUT SHiftEnd",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type)
in
    #"Added Index";

shared #"Master Prepare" = let
    Source = #"IMPORT Master",
    LOCATION = Table.TransformColumns(Source, {{"Location", each Text.Start(_, 2), type text}}),
    #"Removed Other Columns" = Table.SelectColumns(LOCATION,{"Location", "Role", "Employee Code", "Week No", "Week Day", "Start Time", "End Time", "Roster Hours"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Role", "RoleX"}}),
    ROLES = Table.AddColumn(#"Renamed Columns", "Role", each if Text.Contains([RoleX], "REGN") then "RN" else if [RoleX] = "Asst in Nursing Med Comp" then "AINC4" else if Text.Contains([RoleX], "Asst") then "AIN" else if Text.Contains([RoleX], "Enrolled Nurse") then "EN" else if Text.Contains([RoleX], "Wellbeing & Lifestyle Officer") then "WLO" else if Text.Contains([RoleX], "Therapy") then "TA" else [RoleX], type text),
    #"Filtered ROLES" = Table.SelectRows(ROLES, each ([Role] = "AIN" or [Role] = "EN" or [Role] = "RN" or [Role] = "TA" or [Role] = "Wellbeing & Lifestyle Officer" or [Role] = "WLO")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered ROLES",{"RoleX"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each true),
    #"Filtered FACILITIES" = Table.SelectRows(#"Filtered Rows", each ([Location] <> "NR")),
    SHIFTS = Table.AddColumn(#"Filtered FACILITIES", "Shift", each if [Start Time] <= StartAM then "NS" else if [Start Time] <= StartPM then "AM" else if [Start Time] <= StartNS then "PM" else if [Start Time] <= #time(23, 59, 59) then "NS" else "ERROR"),
    #"Merged Queries" = Table.NestedJoin(SHIFTS, {"Shift"}, #"ShftEnd Prepare", {"Shift"}, "ShftEnd Prepare", JoinKind.LeftOuter),
    #"Expanded ShftEnd Prepare" = Table.ExpandTableColumn(#"Merged Queries", "ShftEnd Prepare", {"Index"}, {"ShiftIndex"})
in
    #"Expanded ShftEnd Prepare";

shared LocRoleWeekDaysHours = let
    Source = #"Master Prepare",
    #"added DAYOFWEEKS#" = Table.AddColumn(Source, "DayOfWeek", each List.PositionOf(
    {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},
    [Week Day]
) + 1),
    #"Sorted Rows" = Table.Sort(#"added DAYOFWEEKS#",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Grouped Rows" = Table.Group(#"Sorted Rows", {"Location", "Week No", "Role", "DayOfWeek", "Week Day", "Shift", "ShiftIndex"}, {{"Hours", each List.Sum([Roster Hours]), type nullable number}}),
    #"Sorted Rows1" = Table.Sort(#"Grouped Rows",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Inserted Merged Column1" = Table.AddColumn(#"Sorted Rows1", "WeekDayShift", each Text.Combine({Text.From([DayOfWeek], "en-AU"), [Week Day], [Shift]}, ""), type text),
    #"Inserted Merged Column" = Table.AddColumn(#"Inserted Merged Column1", "WeekAndDay", each Text.Combine({Text.From([Week No], "en-AU"), [Week Day]}, ""), type text),
    #"Inserted Merged Column2" = Table.AddColumn(#"Inserted Merged Column", "DayShift", each Text.Combine({[Week Day], [Shift]}, ""), type text)
in
    #"Inserted Merged Column2";

shared StartPM = let
    Source = #"ShftEnd Prepare",
    End = Source{0}[End]
in
    End;

shared StartNS = let
    Source = #"ShftEnd Prepare",
    End = Source{1}[End]
in
    End;

shared StartAM = let
    Source = #"ShftEnd Prepare",
    End = Source{2}[End]
in
    End;

shared LocRoleWeekDaysHoursMATRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "WeekDayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[WeekDayShift]), "WeekDayShift", "Hours", List.Sum),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVEMATTRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}, {"ShiftIndex", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "DayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[DayShift]), "DayShift", "Hours", List.Average),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVETABLE = let
    Source = LocRoleWeekDaysShiftAVEMATTRIX,
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Location", "Role"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

shared LocRoleTOTAL = let
    Source = LocRoleWeekDaysShiftAVETABLE,
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Hours", each List.Sum([Value]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared #"LocRoleDayShift%" = let
    Source = Table.NestedJoin(LocRoleWeekDaysShiftAVETABLE, {"Location", "Role"}, LocRoleTOTAL, {"Location", "Role"}, "LocRoleTOTAL", JoinKind.LeftOuter),
    #"Expanded LocRoleTOTAL" = Table.ExpandTableColumn(Source, "LocRoleTOTAL", {"Hours"}, {"Hours"}),
    #"Inserted Division" = Table.AddColumn(#"Expanded LocRoleTOTAL", "LocRoleDayShift%", each [Value] / [Hours], type number),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Division",{{"LocRoleDayShift%", Percentage.Type}})
in
    #"Changed Type";

shared #"LocRoleDayShift%CHECK" = let
    Source = #"LocRoleDayShift%",
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Check", each List.Sum([#"LocRoleDayShift%"]), type number}})
in
    #"Grouped Rows";

shared #"LocRoleDatShift%MATRIX" = let
    Source = #"LocRoleDayShift%",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Value", "Hours"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Attribute]), "Attribute", "LocRoleDayShift%", List.Sum)
in
    #"Pivoted Column";

shared #"LocRoleWeekDaysHours (2)" = let
    Source = LocRoleWeekDaysHours
in
    Source;