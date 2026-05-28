// Power Query from: 2-DemandExtract.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\1. Input\2-DemandExtract.xlsx
// Extracted: 2026-05-21T00:46:34.875Z

section Section1;

shared StartAM = let
    Source = ShiftStart,
    StartAM1 = Source{0}[StartAM]
in
    StartAM1;

shared StartPM = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartPM]
in
    StartPM1;

shared StartNIGHT = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartNIGHT]
in
    StartPM1;

shared #"EndNIGHT(7)" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMtAzsjQ0MzMzV4qNBQA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Time = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Time", type number}}),
    Time = #"Changed Type"{0}[Time]
in
    Time;

shared DemandStartDays = let
    Source = IMPORTMasterSchedule,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Column8"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Column8] <> null and [Column8] <> "Start Time")),
    #"Calculated Earliest" = List.Min(#"Filtered Rows"[Column8]),
    #"Converted to Table" = #table(1, {{#"Calculated Earliest"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Converted to Table",{{"Column1", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "DemandStart"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "AddDays", each Table_RosterStart-[DemandStart]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"AddDays", Int64.Type}}),
    #"Subtracted from Column" = Table.TransformColumns(#"Changed Type1", {{"AddDays", each _ - 1, type number}}),
    AddDays = #"Subtracted from Column"{0}[AddDays]
in
    AddDays;

shared ShiftOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t, Column2 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "Shift"}, {"Column2", "Order"}})
in
    #"Renamed Columns";

shared DayOrderDate = let
    Source = ExtractMatrix,
    #"Kept First Rows" = Table.FirstN(Source,1),
    #"Removed Columns" = Table.RemoveColumns(#"Kept First Rows",{"Role", "Start Time", "End Time", "Shift Length"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"Unit"}, "Attribute", "Value"),
    #"Removed Columns1" = Table.RemoveColumns(#"Unpivoted Other Columns",{"Unit"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns1", "Index", 1, 1, Int64.Type),
    #"Removed Columns2" = Table.RemoveColumns(#"Added Index",{"Value"})
in
    #"Removed Columns2";

shared #"ShiftUnitDemandHRS-wide B" = let
    Source = ExtractMatrix,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Role] = "Carer" or [Role] = "Endorsed Enrolled Nurse" or [Role] = "Registered Nurse")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Start Time", Order.Ascending}}),
    #"Inserted SHIFTDURATION" = Table.AddColumn(#"Sorted Rows", "ShiftDuration", each Time.Minute([Shift Length])/60 + Time.Hour([Shift Length])),
    #"Reordered Columns1" = Table.ReorderColumns(#"Inserted SHIFTDURATION",{"Unit", "Role", "Start Time", "End Time", "Shift Length", "ShiftDuration", "A-Mon", "A-Tues", "A-Wed", "A-Thur", "A-Fri", "A-Sat", "A-Sun", "B-Mon", "B-Tues", "B-Wed", "B-Thur", "B-Fri", "B-Sat", "B-Sun"}),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Reordered Columns1", "Start Time", "Start Time - Copy"),
    #"Changed Type1" = Table.TransformColumnTypes(#"Duplicated Column1",{{"Start Time - Copy", type time}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type1",{{"Start Time - Copy", type number}}),
    #"Added SHIFT" = Table.AddColumn(#"Changed Type2", "Shift", each if [#"Start Time - Copy"] >= StartAM
and [#"Start Time - Copy"] < StartPM
then "AM"

else if [#"Start Time - Copy"] >= StartPM
and [#"Start Time - Copy"] < StartNIGHT
then "PM"

else "NIGHT"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added SHIFT",{"Shift Length", "Start Time - Copy"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns1", {"ShiftDuration", "End Time", "Start Time", "Role", "Unit","Shift"}, "RosterDay", "DemandFTE"),
    InsertedDEMANDHRS = Table.AddColumn(#"Unpivoted Other Columns", "DemandHrs", each [ShiftDuration] * [DemandFTE], Int64.Type),
    #"Duplicated Column" = Table.DuplicateColumn(InsertedDEMANDHRS, "Start Time", "Start Time - Copy"),
    #"Changed Type" = Table.TransformColumnTypes(#"Duplicated Column",{{"Start Time - Copy", type number}}),
    BUFFER = Table.Buffer(     #"Changed Type"),
    #"Grouped Rows" = Table.Group(BUFFER, {"Unit", "Role", "Start Time", "End Time", "RosterDay", "Shift", "ShiftDuration"}, {{"DemandFTE", each List.Sum([DemandFTE]), type number}, {"DemandHRS", each List.Sum([DemandHrs]), type number}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Grouped Rows",{"RosterDay", "Shift", "Unit", "Role", "Start Time", "End Time", "DemandFTE", "DemandHRS"}),
    #"Merged DAYORDER" = Table.NestedJoin(#"Reordered Columns", {"RosterDay"}, DayOrderDate, {"Attribute"}, "DayOrder", JoinKind.LeftOuter),
    #"Expanded DayOrder" = Table.ExpandTableColumn(#"Merged DAYORDER", "DayOrder", {"Index"}, {"Index"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded DayOrder",{{"Role", Order.Ascending}, {"Index", Order.Ascending}}),
    #"Merged SHIFTORDER" = Table.NestedJoin(#"Sorted Rows1", {"Shift"}, ShiftOrder, {"Shift"}, "ShiftOrder", JoinKind.LeftOuter),
    #"Expanded ShiftOrder" = Table.ExpandTableColumn(#"Merged SHIFTORDER", "ShiftOrder", {"Order"}, {"Order"}),
    #"Added Custom1" = Table.AddColumn(#"Expanded ShiftOrder", "Facility", each "A"),
    #"Added STARTTIME" = Table.AddColumn(#"Added Custom1", "StartTime", each Date.AddDays([Start Time],[Index]+DemandStartDays)),
    #"Inserted Time" = Table.AddColumn(#"Added STARTTIME", "Time", each DateTime.Time([End Time]), type time),
    #"Changed Type5" = Table.TransformColumnTypes(#"Inserted Time",{{"Time", type number}}),
    #"Added ENDTIME" = Table.AddColumn(#"Changed Type5", "EndTime", each if [Shift] = "NIGHT"
and
[Time] < #"EndNIGHT(7)"  
then Date.AddDays([End Time],[Index]+DemandStartDays+1) 
else
Date.AddDays([End Time],[Index]+DemandStartDays)),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Added ENDTIME", {{"RosterDay", each Text.AfterDelimiter(_, "-"), type text}}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Extracted Text After Delimiter",{{"Start Time", type datetime}, {"End Time", type datetime}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type3",{"Start Time", "End Time", "Time"}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Removed Columns",{{"StartTime", type datetime}, {"EndTime", type datetime}}),
    #"Filtered Rows1" = Table.SelectRows(#"Changed Type4", each true)
in
    #"Filtered Rows1";

shared DemandWeeksHrsHRS = let
    Source = #"ShiftUnitDemandHRS - Master",
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role"}, {{"DemandRosterHRS", each List.Sum([DemandHRS]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DemandWeekHrs", each [DemandRosterHRS]/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "CareDemandWeekHrsApprox", each [DemandWeekHrs]*.9366)
in
    #"Added Custom1";

shared ExtractMatrix = let
    Source = IMPORTMasterSchedule,
    #"Removed Top Rows" = Table.Skip(Source,8),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers",{"Column1", "Column2", "Shift Label", "Column25", "Ogranisation path", "Department"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Mon", "A-Mon"}, {"Tues", "A-Tues"}, {"Wed", "A-Wed"}, {"Thur", "A-Thur"}, {"Fri", "A-Fri"}, {"Sat", "A-Sat"}, {"Sun", "A-Sun"}, {"Mon2", "B-Mon"}, {"Tues3", "B-Tues"}, {"Wed4", "B-Wed"}, {"Thur5", "B-Thur"}, {"Fri6", "B-Fri"}, {"Sat7", "B-Sat"}, {"Sun8", "B-Sun"}, {"Job Description", "Role"}, {"Area /Wing if applicable", "Unit"}})
in
    #"Renamed Columns";

shared #"ShiftUnitDemandHRS-no time" = let
    Source = #"ShiftUnitDemandHRS - Master",
    #"Grouped Rows" = Table.Group(Source, {"Shift", "Role", "Facility", "Unit", "Date"}, {{"DemandFTE", each List.Sum([DemandFTE]), type number}, {"DemandHRS", each List.Sum([DemandHRS]), type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Date", "DateX"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX],28)),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"DateX"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Date", type date}})
in
    #"Changed Type";

shared #"ShiftUnitDemandHRS - Master" = let
    Source = #"ShiftUnitDemandHRS-wide B",
    BUFFER = Table.Buffer(Source),
    #"Grouped Rows" = Table.Group(BUFFER, {"Shift", "Role", "Facility", "StartTime", "EndTime", "Unit", "ShiftDuration"}, {{"DemandFTE", each List.Sum([DemandFTE]), type number}, {"DemandHRS", each List.Sum([DemandHRS]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"StartTime", Order.Ascending}}),
    #"Inserted Date" = Table.AddColumn(#"Sorted Rows", "Date", each DateTime.Date([StartTime]), type date)
in
    #"Inserted Date";

shared IMPORTMasterSchedule = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\03. Incoming\230602-Data Meeting\Benhome_Master Roster_Shift Templates_21.11.22.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Sheet1_Sheet,{{"Column1", type any}, {"Column2", type any}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type any}, {"Column11", type any}, {"Column12", type any}, {"Column13", type any}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}, {"Column22", type any}, {"Column23", type any}, {"Column24", type any}, {"Column25", type any}})
in
    #"Changed Type";

shared #"IMPORT RoleShiftDemandANACC" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\1. Input\Demand-MasterRoster Manual Read.xlsx"), null, true),
    RoleShiftDemandANACC_Table = Source{[Item="RoleShiftDemandANACC",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RoleShiftDemandANACC_Table,{{"Role", type text}, {"AM", type number}, {"PM", type number}, {"NIGHT", type number}, {"Total", type number}})
in
    #"Changed Type";

shared ShiftStart = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftStart_Table = Source{[Item="ShiftStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftStart_Table,{{"StartAM", type number}, {"StartPM", type number}, {"StartNIGHT", type number}})
in
    #"Changed Type";

shared Table_RosterStart = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\1. Input\1-AllocationExtracted.xlsx"), null, true),
    Table_RosterStart_Table = Source{[Item="Table_RosterStart",Kind="Table"]}[Data],
    RosterStart = Table_RosterStart_Table{0}[RosterStart]
in
    RosterStart;

shared Table_RosteredDays = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\1. Input\1-AllocationExtracted.xlsx"), null, true),
    Table_RosteredDays_Table = Source{[Item="Table_RosteredDays",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_RosteredDays_Table,{{"RosteredDays", Int64.Type}})
in
    #"Changed Type";

shared ShiftDuration = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared #"IMPORT ShiftPeriod" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}, {"DurationOfShifts", type number}})
in
    #"Changed Type";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\01 Christadelphian Aged Care\240715\FACILITIES\ASHB\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared #"RoleShiftDemandANACC Prepare" = let
    Source = #"IMPORT RoleShiftDemandANACC",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"Role"}, "Shift", "DemandFTE"),
    #"Merged Queries" = Table.NestedJoin(#"Unpivoted Other Columns", {"Shift", "Role"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "IMPORT ShiftPeriod", JoinKind.LeftOuter),
    #"Expanded IMPORT ShiftPeriod" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ShiftPeriod", {"StartTime", "DurationOfShifts"}, {"StartTime", "DurationOfShifts"}),
    #"Inserted Multiplication" = Table.AddColumn(#"Expanded IMPORT ShiftPeriod", "DemandHRS", each [DurationOfShifts] * [DemandFTE], type number),
    #"Sorted Rows" = Table.Sort(#"Inserted Multiplication",{{"Shift", Order.Ascending}})
in
    #"Sorted Rows";

shared #"Permutation DateTimeRoleShift" = let
    Source = #"IMPORT PermutationDimensions",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"RolesList", "Roles"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns", {"Shifts", "Roles"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "IMPORT ShiftPeriod", JoinKind.LeftOuter),
    #"Expanded IMPORT ShiftPeriod1" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ShiftPeriod", {"StartDay", "EndDay"}, {"StartDay", "EndDay"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Expanded IMPORT ShiftPeriod1",{{"StartDay", "Start"}, {"EndDay", "End"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns2",{{"Start", type time}, {"End", type time}}),
    #"START TIMEDATE" = Table.AddColumn(#"Changed Type", "StartTime", each [Date] & [Start], type datetime),
    #"END TIMEDATE" = Table.AddColumn(#"START TIMEDATE", "EndTime", each if [Shifts] = "NIGHT" then Date.AddDays([Date],1) & [End]
else [Date]&[End]),
    #"Changed Type1" = Table.TransformColumnTypes(#"END TIMEDATE",{{"EndTime", type datetime}}),
    #"Added UNIT" = Table.AddColumn(#"Changed Type1", "Unit", each null),
    #"Added FACILITY" = Table.AddColumn(#"Added UNIT", "Facility", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added FACILITY",{"Start", "End"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Shifts", "Shift"}, {"Roles", "Role"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns1",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared ShiftUnitDemandHRS = let
    Source = #"Permutation DateTimeRoleShift",
    #"Merged Queries" = Table.NestedJoin(Source, {"Role", "Shift"}, #"RoleShiftDemandANACC Prepare", {"Role", "Shift"}, "RoleShiftDemandANACC Prepare", JoinKind.LeftOuter),
    #"Expanded RoleShiftDemandANACC Prepare" = Table.ExpandTableColumn(#"Merged Queries", "RoleShiftDemandANACC Prepare", {"DemandFTE", "DemandHRS", "DurationOfShifts"}, {"DemandFTE", "DemandHRS", "DurationOfShifts"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded RoleShiftDemandANACC Prepare",{{"Period", Order.Ascending}, {"Role", Order.Ascending}, {"StartTime", Order.Ascending}, {"EndTime", Order.Ascending}})
in
    #"Sorted Rows1";