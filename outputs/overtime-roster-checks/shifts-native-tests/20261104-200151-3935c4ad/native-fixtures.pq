section OvertimeShiftsIntegrationNative;
shared OvertimeShiftsIntegrationNativeTests = () as table => let
EvaluateBaseline = (Inputs as record) as record => let
UnitL1PathTABLE = Inputs[PathTable],
#"FilePath - 2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value,
#"FilePath - 1Input" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value,
#"IMPORT Table_AllocationExtracted" = Inputs[Roster],
#"IMPORT Intervals" = Inputs[IntervalsNavigation],
StaffList = let
    Source = #"IMPORT Table_AllocationExtracted",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Date", "End", "Unit", "Start"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates",
MealBreakTime = let
    Source = #"EXTRACT MealBreak",
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1,
MealBreakStart = let
    Source = #"EXTRACT MealBreak",
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart,
MinShiftGap = let
    Source = #"IMPORT Settings Data",
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap,
ShiftLength = let
    Source = #"IMPORT Settings Data",
    ShiftLent_Table = Source{[Item="ShiftLent",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftLent_Table,{{"MinHrs", Int64.Type}, {"ShortHrs", Int64.Type}, {"StdHrs", Int64.Type}, {"1.5OT", type number}, {"2.0OT", Int64.Type}})
in
    #"Changed Type",
StartNames = let
    Source = Table.NestedJoin(Table_Intervals, {"StartInterval", "Role"}, #"EXTRACT Table_StartTime_Duration", {"StartDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(Source, "StartTime+Duration", {"Name"}, {"Name"}),
    STARTLABEL = Table.AddColumn(#"Expanded StartTime+Duration", "TimeType", each "Start"),
    #"Renamed Columns" = Table.RenameColumns(STARTLABEL,{{"StartInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"EndPrecise", "EndInterval"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Date", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows",
EndNames = let
    Source = Table.NestedJoin(Table_Intervals, {"EndPrecise", "Role"}, #"EXTRACT Table_StartTime_Duration", {"EndDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndInterval", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type",{{"EndInterval", Order.Ascending}}),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(#"Sorted Rows1", "StartTime+Duration", {"Name"}, {"Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded StartTime+Duration",{{"EndInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"StartInterval", "EndPrecise"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "TimeType", each "ShiftEnd")
in
    #"Added Custom",
#"NAMES+INTERVALS" = let
    Source = Table.Combine({StartNames, EndNames}),
    #"Removed Columns1" = Table.RemoveColumns(Source,{"ShiftPeriod"}),
    #"Replaced NULL-->X" = Table.ReplaceValue(#"Removed Columns1",null,"x",Replacer.ReplaceValue,{"Name"}),
    BUFFER = Table.Buffer(#"Replaced NULL-->X"),
    SORTNAMESDATES = Table.Sort(BUFFER,{{"Name", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(SORTNAMESDATES, "Index", 1, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 2, 1, Int64.Type),
    #"Merged INDEX" = Table.NestedJoin(#"Added Index1", {"Index.1"}, #"Added Index1", {"Index"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded INTERVALID" = Table.ExpandTableColumn(#"Merged INDEX", "Added Index1", {"IntervalID"}, {"Added Index1.IntervalID"}),
    FILTERNULLINTERVALID = Table.SelectRows(#"Expanded INTERVALID", each ([Added Index1.IntervalID] <> null)),
    FILTERSTART = Table.SelectRows(FILTERNULLINTERVALID, each [TimeType] = "Start"),
    REMOVEINDEXS = Table.RemoveColumns(FILTERSTART,{"Index", "Index.1"}),
    INTERVALRANGE = Table.AddColumn(REMOVEINDEXS, "Interval Range", each if [Added Index1.IntervalID]-[IntervalID] <=0 
then 0 
else [Added Index1.IntervalID]-[IntervalID]),
    INTERVALLIST = Table.AddColumn(INTERVALRANGE, "AllocatedIntervals", each List.Numbers([IntervalID],[Interval Range]+1)),
    #"Expanded AllocatedIntervals" = Table.ExpandListColumn(INTERVALLIST, "AllocatedIntervals"),
    BUFFER1 = Table.Buffer(#"Expanded AllocatedIntervals"),
    #"Merged INTERVALS" = Table.NestedJoin(BUFFER1, {"AllocatedIntervals", "Role"}, Table_Intervals, {"IntervalID", "Role"}, "Intervals", JoinKind.LeftOuter),
    #"Expanded Intervals" = Table.ExpandTableColumn(#"Merged INTERVALS", "Intervals", {"StartInterval", "IntervalID", "EndInterval", "Duration", "ShiftPeriod"}, {"StartInterval", "IntervalID.1", "EndInterval", "Duration.1", "ShiftPeriod"}),
    #"Sorted Rows3" = Table.Sort(#"Expanded Intervals",{{"Name", Order.Ascending}, {"StartInterval", Order.Ascending}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Sorted Rows3",{{"Duration.1", "IntervalDurationTemp"}}),
    #"Sorted Rows2" = Table.Sort(#"Renamed Columns1",{{"Date", Order.Ascending}, {"StartInterval", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows2",{{"StartInterval", "IntervalStart"}, {"EndInterval", "IntervalEnd"}, {"IntervalID", "ShiftStartInterval"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Added Index1.IntervalID", "Interval Range", "Date", "Duration", "ShiftStartInterval"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"IntervalStart", "IntervalEnd",  "Name", "AllocatedIntervals"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Name", Order.Ascending}, {"IntervalStart", Order.Ascending}}),
    #"Merged STAFFSHIFTS" = Table.NestedJoin(#"Sorted Rows", {"Name", "IntervalStart"}, StaffDoubleShifts, {"Name", "DateTime"}, "StaffDoubleShifts", JoinKind.LeftOuter),
    #"Expanded StaffDoubleShifts" = Table.ExpandTableColumn(#"Merged STAFFSHIFTS", "StaffDoubleShifts", {"Double Shift", "Effective Duration", "RealDuration", "ShiftType"}, {"Double Shift", "Effective Duration", "RealDuration", "ShiftType"}),
    #"Filled Down" = Table.FillDown(#"Expanded StaffDoubleShifts",{"Effective Duration", "ShiftType", "RealDuration"}),
    #"Sorted Rows1" = Table.Sort(#"Filled Down",{{"IntervalStart", Order.Ascending}}),
    BUFFER2 = Table.Buffer(#"Sorted Rows1")
in
    BUFFER2,
#"NAMES+INTERVALSLIST" = let
    Source = #"NAMES+INTERVALS",
    BUFFER1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER1, {"IntervalStart", "Role"}, #"Table Intervals-IntervalAssocShift", {"StartInterval", "Role"}, "Table Intervals-IntervalAssocShift", JoinKind.LeftOuter),
    #"Expanded Table Intervals-IntervalAssocShift" = Table.ExpandTableColumn(#"Merged Queries", "Table Intervals-IntervalAssocShift", {"IntervalAssociatedShift"}, {"IntervalAssociatedShift"}),
    #"Unpivoted Only Selected Columns" = Table.Unpivot(#"Expanded Table Intervals-IntervalAssocShift", {"IntervalStart", "IntervalEnd"}, "Attribute", "Value"),
    #"Changed Type" = Table.TransformColumnTypes(#"Unpivoted Only Selected Columns",{{"Value", type datetime}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{  {"Value", "TimeDate"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Name", "AllocatedIntervals", "ShiftPeriod", "Role",  "TimeDate", "IntervalDurationTemp",  "Double Shift", "Effective Duration", "ShiftType"}),
    TIME = Table.AddColumn(#"Reordered Columns", "Time", each DateTime.Time([TimeDate]), type time),
    #"Changed Type1" = Table.TransformColumnTypes(TIME,{{"Time", type number}}),
    SHIFTDATE = Table.AddColumn(#"Changed Type1", "ShiftDate", each if [Time] > 0 and [Time] < AMP
then Date.AddDays(DateTime.Date([TimeDate]), -1)
else  DateTime.Date([TimeDate])),
    #"Changed Type2" = Table.TransformColumnTypes(SHIFTDATE,{{"ShiftDate", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type2",{"Time"}),
    BUFFER = Table.Buffer(#"Removed Columns"),
    SORTEDDATETIME = Table.Sort(BUFFER,{{"TimeDate", Order.Ascending}}),
    ROLESALL = Table.AddColumn(SORTEDDATETIME, "Role1", each RolesList),
    #"Expanded RoleList" = Table.ExpandListColumn(ROLESALL, "Role1"),
    STAFFCOUNT = Table.AddColumn(#"Expanded RoleList", "StaffCount", each if [Role1] = [Role] then 1 else 0),
    #"Replaced NULL X" = Table.ReplaceValue(STAFFCOUNT,"x",null,Replacer.ReplaceValue,{"Name"}),
    #"Filtered NAMES ONLY" = Table.SelectRows(#"Replaced NULL X", each ([Name] <> null))
in
    #"Filtered NAMES ONLY",
StaffShifts = let
    Source = #"EXTRACT Table_StartTime_Duration",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"StartDateTime", "StartOutter"}, {"EndDateTime", "EndOutter"}}),
    STARTOUTTER = Table.AddColumn(#"Renamed Columns", "StartInner", each [StartOutter] + #duration(0,0,1,0)),
    ENDINNER = Table.AddColumn(STARTOUTTER, "EndInner", each [EndOutter]-#duration(0,0,1,0)),
    #"Reordered Columns" = Table.ReorderColumns(ENDINNER,{"Name", "Role",  "StartOutter", "StartInner", "EndOutter", "Duration", "Shift"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"StartInner", type datetime}, {"EndInner", type datetime}}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Name", "Role",  "Duration", "Shift"}, "Attribute", "Value"),
    #"Renamed Columns1" = Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "IntervalPoint"}, {"Value", "DateTime"}}),
    STAFFCOUNT = Table.AddColumn(#"Renamed Columns1", "StaffCount", each if [IntervalPoint] = "StartOuter" then 0 else if [IntervalPoint] = "StartInner" then 1 else if [IntervalPoint] = "EndInner" then 1 else 0),
    BUFFER = Table.Buffer(STAFFCOUNT)
in
    BUFFER,
ShiftIndex = let
    Source = StaffShifts,
    #"Removed Columns" = Table.RemoveColumns(Source,{"StaffCount", "Role", "Duration", "Shift"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([IntervalPoint] = "StartOutter")),
    SHIFTINDEX = Table.AddIndexColumn(#"Filtered Rows", "ShiftIndex", 1, 1, Int64.Type)
in
    SHIFTINDEX,
MinHrs = let
    Source = ShiftLength,
    MinHrs = Source{0}[MinHrs]
in
    MinHrs,
ShortHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[ShortHrs]
in
    ShortHrs,
StdHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[StdHrs]
in
    ShortHrs,
#"15OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[1.5OT]
in
    ShortHrs,
#"20OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[2.0OT]
in
    ShortHrs,
ShiftPeriod = let
    Source = #"IMPORT Settings Data",
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type",
ShiftGapIndices = let
    Source = #"EXTRACT Table_StartTime_Duration",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Name", Order.Ascending}, {"StartDateTime", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"StartDateTime", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name"}, {{"Count", each _, type table [Name=text, Role=nullable text, WARD=nullable text, StartDateTime=datetime, EndDateTime=nullable datetime, Duration=number, Shift=text]}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "Custom", each Table.AddIndexColumn([Count],"Index",0,1)),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Count"}),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Removed Columns", "Custom", {"Role", "WARD", "StartDateTime", "EndDateTime", "Duration", "Shift", "Index"}, {"Custom.Role", "Custom.WARD", "Custom.StartDateTime", "Custom.EndDateTime", "Custom.Duration", "Custom.Shift", "Custom.Index"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded Custom",{{"Custom.StartDateTime", type datetime}}),
    #"Grouped Rows1" = Table.Group(#"Changed Type1", {"Name"}, {{"Count2", each _, type table [Name=text, Custom.Role=text, Custom.WARD=text, Custom.StartDateTime=datetime, Custom.EndDateTime=datetime, Custom.Duration=number, Custom.Shift=text, Custom.Index=number]}}),
    #"Added Custom1" = Table.AddColumn(#"Grouped Rows1", "Index2", each Table.AddIndexColumn([Count2],"Index2",1,1)),
    #"Expanded Index2" = Table.ExpandTableColumn(#"Added Custom1", "Index2", {"Custom.Role", "Custom.WARD", "Custom.StartDateTime", "Custom.EndDateTime", "Custom.Duration", "Custom.Shift", "Custom.Index", "Index2"}, {"Index2.Custom.Role", "Index2.Custom.WARD", "Index2.Custom.StartDateTime", "Index2.Custom.EndDateTime", "Index2.Custom.Duration", "Index2.Custom.Shift", "Index2.Custom.Index", "Index2.Index2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Index2",{"Count2", "Index2.Custom.WARD"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns1",{{"Index2.Custom.StartDateTime", type number}})
in
    #"Changed Type2",
ShiftGap = let
    Source = Table.NestedJoin(ShiftGapIndices, {"Name", "Index2.Index2"}, ShiftGapIndices, {"Name", "Index2.Custom.Index"}, "StaffShifts (2)", JoinKind.LeftOuter),
    #"Sorted Rows" = Table.Sort(Source,{{"Name", Order.Ascending}, {"Index2.Custom.StartDateTime", Order.Ascending}}),
    BUFFER = Table.Buffer(#"Sorted Rows"),
    #"Expanded StaffShifts (2)" = Table.ExpandTableColumn(BUFFER, "StaffShifts (2)", {"Index2.Custom.StartDateTime", "Index2.Custom.Duration"}, {"StaffShifts (2).Index2.Custom.StartDateTime", "StaffShifts (2).Index2.Custom.Duration"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded StaffShifts (2)",{{"Index2.Custom.StartDateTime", type datetime}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{ "Index2.Custom.Shift", "Index2.Custom.Index", "Index2.Index2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns",{{"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type2",null,0,Replacer.ReplaceValue,{"StaffShifts (2).Index2.Custom.Duration"}),
    #"Replaced Value" = Table.ReplaceValue(#"Replaced Value1",null,0,Replacer.ReplaceValue,{"StaffShifts (2).Index2.Custom.StartDateTime"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Replaced Value",{{"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    SHIFTGAP = Table.AddColumn(#"Changed Type3", "Subtraction", each if [#"StaffShifts (2).Index2.Custom.Duration"] = 0
then 99
else

[Index2.Custom.EndDateTime] - [#"StaffShifts (2).Index2.Custom.StartDateTime"]),
    #"Renamed Columns" = Table.RenameColumns(SHIFTGAP,{{"Subtraction", "EndStartDifference"}}),
    BUFFER2 = Table.Buffer(#"Renamed Columns")
in
    BUFFER2,
EffectiveDuration = let
    Source = ShiftGap,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndStartDifference", type number}, {"Index2.Custom.StartDateTime", type datetime}, {"Index2.Custom.EndDateTime", type datetime}, {"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    DOUBLESHIFT = Table.AddColumn(#"Changed Type", "DoubleShift1", each if [EndStartDifference] = 99 
or 
[EndStartDifference]*-1 >= MinShiftGap/24
then null
else "DoubleShift1"),
    COMBINEDDURATIONS = Table.AddColumn(DOUBLESHIFT, "CombinedDuration", each if [DoubleShift1] = "DoubleShift1" 
then [Index2.Custom.Duration] +  [#"StaffShifts (2).Index2.Custom.Duration"]
else [Index2.Custom.Duration]),
    EFFECTIVEDURATION = Table.AddColumn(COMBINEDDURATIONS, "EffectiveDuration", each if [CombinedDuration] < MealBreakStart/24
then [CombinedDuration] 

else if [CombinedDuration] > 2*MealBreakStart/24
then [CombinedDuration] - 2*MealBreakTime/24

else [CombinedDuration] - MealBreakTime/24),
    #"Added Index" = Table.AddIndexColumn(EFFECTIVEDURATION, "Index", 0, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 1, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index1",{"StaffShifts (2).Index2.Custom.StartDateTime", "StaffShifts (2).Index2.Custom.Duration", "EndStartDifference"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER,
DoubleShifts = let
    Source = EffectiveDuration,
    #"Merged Queries1" = Table.NestedJoin(Source, {"Index"}, Source, {"Index.1"}, "Source", JoinKind.LeftOuter),
    #"Expanded Source" = Table.ExpandTableColumn(#"Merged Queries1", "Source", {"CombinedDuration", "DoubleShift1", "EffectiveDuration"}, {"Source.CombinedDuration", "Source.DoubleShift1", "Source.EffectiveDuration"}),
    #"DOUBLE SHIFT" = Table.AddColumn(#"Expanded Source", "Double Shift", each if [DoubleShift1] = "DoubleShift1" then "DoubleShift1" else if [Source.DoubleShift1] = "DoubleShift1" then "DoubleShift2" else null),
    REALDURATION = Table.AddColumn(#"DOUBLE SHIFT", "RealDuration", each if [Double Shift] = "DoubleShift1" then [CombinedDuration] else if [Double Shift] = "DoubleShift2" then [Source.CombinedDuration] else [CombinedDuration]),
    #"EFFECTIVE SHIFT" = Table.AddColumn(REALDURATION, "Effective Duration", each if [Double Shift] = "DoubleShift1" then [EffectiveDuration] else if [Double Shift] = "DoubleShift2" then [Source.EffectiveDuration] else [EffectiveDuration])
in
    #"EFFECTIVE SHIFT",
StaffDoubleShifts = let
    Source = Table.NestedJoin(StaffShifts, {"Name", "DateTime"}, ShiftIndex, {"Name", "DateTime"}, "EffectiveShift.1", JoinKind.LeftOuter),
    #"Expanded EffectiveShift.2" = Table.ExpandTableColumn(Source, "EffectiveShift.1", {"DateTime"}, {"EffectiveShift.1.DateTime"}),
    #"Sorted Rows" = Table.Sort(#"Expanded EffectiveShift.2",{{"Name", Order.Ascending}, {"DateTime", Order.Ascending}}),
    #"Filled Down" = Table.FillDown(#"Sorted Rows",{"EffectiveShift.1.DateTime"}),
    #"Merged Queries" = Table.NestedJoin(#"Filled Down", {"Name", "EffectiveShift.1.DateTime"}, DoubleShifts, {"Name", "Index2.Custom.StartDateTime"}, "DoubleShifts", JoinKind.LeftOuter),
    #"Expanded DoubleShifts" = Table.ExpandTableColumn(#"Merged Queries", "DoubleShifts", {"Double Shift", "Effective Duration", "RealDuration"}, {"Double Shift", "Effective Duration", "RealDuration"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded DoubleShifts",{"EffectiveShift.1.DateTime"}),
    SHIFTTYPE = Table.AddColumn(#"Removed Columns", "ShiftType", each if [RealDuration] <= 2/24 then "0-4Hrs" else if [RealDuration] <= 7.5/24 then "4-7.5Hrs" else if [RealDuration] <= 8.5/24 then "7.5-8.5Hrs" else if [RealDuration] <= 10/24 then "8.5-10Hrs" else if [RealDuration] <= 12/24 then "10-12Hrs" else ">12Hrs")
in
    SHIFTTYPE,
AMP = let
    Source = ShiftPeriod,
    StartDay = Source{0}[StartDay]
in
    StartDay,
PMP = let
    Source = ShiftPeriod,
    StartDay = Source{1}[StartDay]
in
    StartDay,
NIGHTP = let
    Source = ShiftPeriod,
    StartDay = Source{2}[StartDay]
in
    StartDay,
RolesList = let
    Source = #"LINK Role Analysis",
    ROLES = Source[ROLES]
in
    ROLES,
ShiftPoints = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCs7ITCsJLkksKlGK1YFyXfNSlGJjAQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ShiftPoints = _t]),
    ShiftPoints1 = Source[ShiftPoints]
in
    ShiftPoints1,
#"EXTRACT Table_StartTime_Duration" = let

    Source = #"IMPORT Intervals",
    Table_StartTime_Duration_Table = Source{[Item="Table_StartTime_Duration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_StartTime_Duration_Table,{{"Date", type date}, {"Name", type text}, {"Role", type text}, {"Start", type number}, {"End", type number}, {"StartDateTime", type datetime}, {"EndDateTime", type datetime}, {"Duration", type number}, {"Shift", type text}, {"EffectiveDuration", type number}})
in
    #"Changed Type",
#"EXTRACT Table Intervals" = let

    Source = #"IMPORT Intervals",
    Intervals_Table = Source{[Item="Intervals",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(Intervals_Table,{{"StartInterval", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}, {"EndInterval", type datetime}, {"EndPrecise", type datetime}, {"Duration", type number}, {"Role", type text}, {"ShiftPeriod", type text}, {"IntervalAssociatedShift", type text}})
in
    #"Changed Type1",
#"EXTRACT Table_IntervalsList" = let

    Source = #"IMPORT Intervals",
    Table_IntervalsList_Table = Source{[Item="IntervalsList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_IntervalsList_Table,{{"IntervalID", Int64.Type}, {"DayDate", type date}, {"TimeDate", type datetime}, {"Time", type number}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Attribute", type text}, {"ShiftDate", type date}})
in
    #"Changed Type",
#"IMPORT Settings Data" = Inputs[SettingsNavigation],
#"EXTRACT MealBreak" = let

    Source = #"IMPORT Settings Data",
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type",
#"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type",
#"MANUAL Shifts" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Shifts = _t, Order = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}})
in
    #"Changed Type",
#"LINK Role Analysis" = Inputs[Roles],
Table_Intervals = let
        Source1 = #"EXTRACT Table Intervals",
    #"Removed Columns" = Table.RemoveColumns(Source1,{"IntervalAssociatedShift"})
in
    #"Removed Columns",
#"Table Intervals-IntervalAssocShift" = let
    Source = #"EXTRACT Table Intervals"
in
    Source,
IntervalsList = let
    Source = #"EXTRACT Table_IntervalsList"
in
    Source,
#"NAMES+INTERVALSMATRIX" = let
    Source = #"NAMES+INTERVALS",
    #"Filtered Rows1" = Table.SelectRows(Source, each [IntervalDurationTemp] > 0.1875),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"ShiftPeriod"}, #"MANUAL Shifts", {"Shifts"}, "MANUAL Shifts", JoinKind.LeftOuter),
    #"Expanded MANUAL Shifts" = Table.ExpandTableColumn(#"Merged Queries", "MANUAL Shifts", {"Order"}, {"Order"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded MANUAL Shifts", each ([Name] <> "x")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"DayDate", "Role", "IntervalStart", "Name", "ShiftPeriod", "Order"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"DayDate", "ShiftPeriod", "Order", "Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows1" = Table.Sort(#"Grouped Rows",{{"DayDate", Order.Ascending}, {"Order", Order.Ascending}}),
    BUFFER = Table.Buffer(#"Sorted Rows1"),
    #"Merged Queries1" = Table.NestedJoin(BUFFER, {"DayDate", "ShiftPeriod", "Role"}, #"EXTRACT PermutationDimensions", {"Date", "Shifts", "RolesList"}, "IMPORT PermutationDimensions", JoinKind.FullOuter),
    #"Expanded IMPORT PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries1", "IMPORT PermutationDimensions", {"Date", "RolesList", "Shifts"}, {"Date", "RolesList", "Shifts"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded IMPORT PermutationDimensions",{{"ShiftPeriod", "ShiftPeriodX"}, {"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ShiftPeriod", each if [ShiftPeriodX] = null then [Shifts] else [ShiftPeriodX]),
    #"Added ROLE" = Table.AddColumn(#"Added Conditional Column", "Role", each if [RoleX] = null then [RolesList] else [RoleX]),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Added ROLE", {{"DayDate", type text}}, "en-AU"),{"DayDate", "ShiftPeriod"},Combiner.CombineTextByDelimiter(":", QuoteStyle.None),"Merged"),
    #"Added Custom" = Table.AddColumn(#"Merged Columns", "Allocated", each "X"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Order", "Count", "ShiftPeriodX", "RoleX", "RolesList", "Shifts", "Date"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Merged]), "Merged", "Allocated")
in
    #"Pivoted Column1"
in [Points = StaffDoubleShifts, Intervals = #"NAMES+INTERVALSLIST"],
EvaluateNew = (Inputs as record) as record => let
UnitL1PathTABLE = Inputs[PathTable],
#"FilePath - 2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value,
#"FilePath - 1Input" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value,
#"IMPORT Table_AllocationExtracted" = Inputs[Roster],
#"IMPORT Intervals" = Inputs[IntervalsNavigation],
StaffList = let
    Source = #"IMPORT Table_AllocationExtracted",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Date", "End", "Unit", "Start"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates",
MealBreakTime = let
    Source = #"EXTRACT MealBreak",
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1,
MealBreakStart = let
    Source = #"EXTRACT MealBreak",
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart,
MinShiftGap = let
    Source = #"IMPORT Settings Data",
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap,
ShiftLength = let
    Source = #"IMPORT Settings Data",
    ShiftLent_Table = Source{[Item="ShiftLent",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftLent_Table,{{"MinHrs", Int64.Type}, {"ShortHrs", Int64.Type}, {"StdHrs", Int64.Type}, {"1.5OT", type number}, {"2.0OT", Int64.Type}})
in
    #"Changed Type",
StartNames = let
    Source = Table.NestedJoin(Table_Intervals, {"StartInterval", "Role"}, RosterShiftsWithChecks, {"StartDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    ShiftAttributes = List.RemoveItems(Table.ColumnNames(RosterShiftsWithChecks), {"Code", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(Source, "StartTime+Duration", ShiftAttributes, ShiftAttributes),
    STARTLABEL = Table.AddColumn(#"Expanded StartTime+Duration", "TimeType", each "Start"),
    #"Renamed Columns" = Table.RenameColumns(STARTLABEL,{{"StartInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"EndPrecise", "EndInterval"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Date", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows",
EndNames = let
    Source = Table.NestedJoin(Table_Intervals, {"EndPrecise", "Role"}, RosterShiftsWithChecks, {"EndDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndInterval", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type",{{"EndInterval", Order.Ascending}}),
    ShiftAttributes = List.RemoveItems(Table.ColumnNames(RosterShiftsWithChecks), {"Code", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(#"Sorted Rows1", "StartTime+Duration", ShiftAttributes, ShiftAttributes),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded StartTime+Duration",{{"EndInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"StartInterval", "EndPrecise"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "TimeType", each "ShiftEnd")
in
    #"Added Custom",
#"NAMES+INTERVALS" = let
    Source = Table.Combine({StartNames, EndNames}),
    #"Removed Columns1" = Table.RemoveColumns(Source,{"ShiftPeriod"}),
    #"Replaced NULL-->X" = Table.ReplaceValue(#"Removed Columns1",null,"x",Replacer.ReplaceValue,{"Name"}),
    BUFFER = Table.Buffer(#"Replaced NULL-->X"),
    SORTNAMESDATES = Table.Sort(BUFFER,{{"Name", Order.Ascending}, {"Date", Order.Ascending}}),
    // A shared display name does not identify the other end of this physical shift.
    PairedOwnShiftFinish = Table.NestedJoin(SORTNAMESDATES, {"ShiftFinish", "Role"}, Table_Intervals, {"EndPrecise", "Role"}, "OwnFinishInterval", JoinKind.LeftOuter),
    WithShiftEndIntervalID = Table.AddColumn(PairedOwnShiftFinish, "ShiftEndIntervalID", each
        if [LocalShiftKey] = null then null
        else if Table.RowCount([OwnFinishInterval]) = 1 then [OwnFinishInterval]{0}[IntervalID]
        else if [TimeType] = "Start" then error Error.Record("Shift interval boundary", "A shift requires exactly one finish interval for its own role and original finish.", [LocalShiftKey = [LocalShiftKey]])
        else null, Int64.Type),
    RemovedFinishLookup = Table.RemoveColumns(WithShiftEndIntervalID, {"OwnFinishInterval"}),
    FILTERNULLINTERVALID = Table.SelectRows(RemovedFinishLookup, each ([ShiftEndIntervalID] <> null)),
    FILTERSTART = Table.SelectRows(FILTERNULLINTERVALID, each [TimeType] = "Start"),
    INTERVALRANGE = Table.AddColumn(FILTERSTART, "Interval Range", each if [ShiftEndIntervalID] < [IntervalID]
        then error "A shift's finish interval precedes its start interval."
        else [ShiftEndIntervalID] - [IntervalID]),
    INTERVALLIST = Table.AddColumn(INTERVALRANGE, "AllocatedIntervals", each List.Numbers([IntervalID],[Interval Range]+1)),
    #"Expanded AllocatedIntervals" = Table.ExpandListColumn(INTERVALLIST, "AllocatedIntervals"),
    BUFFER1 = Table.Buffer(#"Expanded AllocatedIntervals"),
    #"Merged INTERVALS" = Table.NestedJoin(BUFFER1, {"AllocatedIntervals", "Role"}, Table_Intervals, {"IntervalID", "Role"}, "Intervals", JoinKind.LeftOuter),
    #"Expanded Intervals" = Table.ExpandTableColumn(#"Merged INTERVALS", "Intervals", {"StartInterval", "IntervalID", "EndInterval", "Duration", "ShiftPeriod"}, {"StartInterval", "IntervalID.1", "EndInterval", "Duration.1", "ShiftPeriod"}),
    #"Sorted Rows3" = Table.Sort(#"Expanded Intervals",{{"Name", Order.Ascending}, {"StartInterval", Order.Ascending}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Sorted Rows3",{{"Duration.1", "IntervalDurationTemp"}}),
    #"Sorted Rows2" = Table.Sort(#"Renamed Columns1",{{"Date", Order.Ascending}, {"StartInterval", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows2",{{"StartInterval", "IntervalStart"}, {"EndInterval", "IntervalEnd"}, {"IntervalID", "ShiftStartInterval"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"ShiftEndIntervalID", "Interval Range", "Date", "Duration", "ShiftStartInterval"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"IntervalStart", "IntervalEnd",  "Name", "AllocatedIntervals"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Name", Order.Ascending}, {"IntervalStart", Order.Ascending}}),
    JoinedOwnLegacyMeasures = Table.NestedJoin(#"Sorted Rows", {"LocalShiftKey"}, LegacyShiftMeasures, {"LocalShiftKey"}, "OwnLegacyMeasures", JoinKind.LeftOuter),
    ExpandedOwnLegacyMeasures = Table.ExpandTableColumn(JoinedOwnLegacyMeasures, "OwnLegacyMeasures", {"Double Shift", "Effective Duration", "RealDuration", "ShiftType", "LegacyShiftStatus", "LegacyShiftReason"}, {"LegacyDoubleShiftAtPoint", "Effective Duration", "RealDuration", "ShiftType", "LegacyShiftStatus", "LegacyShiftReason"}),
    // Preserve the old point-only Double Shift marker; the other three measures describe the whole shift.
    WithLegacyDoubleShiftMarker = Table.AddColumn(ExpandedOwnLegacyMeasures, "Double Shift", each if List.Contains({[ShiftStart], [ShiftStart] + #duration(0, 0, 1, 0), [ShiftFinish] - #duration(0, 0, 1, 0), [ShiftFinish]}, [IntervalStart]) then [LegacyDoubleShiftAtPoint] else null),
    RemovedTemporaryMarker = Table.RemoveColumns(WithLegacyDoubleShiftMarker, {"LegacyDoubleShiftAtPoint"}),
    #"Sorted Rows1" = Table.Sort(RemovedTemporaryMarker,{{"IntervalStart", Order.Ascending}}),
    BUFFER2 = Table.Buffer(#"Sorted Rows1")
in
    BUFFER2,
#"NAMES+INTERVALSLIST" = let
    Source = #"NAMES+INTERVALS",
    BUFFER1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER1, {"IntervalStart", "Role"}, #"Table Intervals-IntervalAssocShift", {"StartInterval", "Role"}, "Table Intervals-IntervalAssocShift", JoinKind.LeftOuter),
    #"Expanded Table Intervals-IntervalAssocShift" = Table.ExpandTableColumn(#"Merged Queries", "Table Intervals-IntervalAssocShift", {"IntervalAssociatedShift"}, {"IntervalAssociatedShift"}),
    #"Unpivoted Only Selected Columns" = Table.Unpivot(#"Expanded Table Intervals-IntervalAssocShift", {"IntervalStart", "IntervalEnd"}, "Attribute", "Value"),
    #"Changed Type" = Table.TransformColumnTypes(#"Unpivoted Only Selected Columns",{{"Value", type datetime}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{  {"Value", "TimeDate"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Name", "AllocatedIntervals", "ShiftPeriod", "Role",  "TimeDate", "IntervalDurationTemp",  "Double Shift", "Effective Duration", "ShiftType"}),
    TIME = Table.AddColumn(#"Reordered Columns", "Time", each DateTime.Time([TimeDate]), type time),
    #"Changed Type1" = Table.TransformColumnTypes(TIME,{{"Time", type number}}),
    SHIFTDATE = Table.AddColumn(#"Changed Type1", "ShiftDate", each if [Time] > 0 and [Time] < AMP
then Date.AddDays(DateTime.Date([TimeDate]), -1)
else  DateTime.Date([TimeDate])),
    #"Changed Type2" = Table.TransformColumnTypes(SHIFTDATE,{{"ShiftDate", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type2",{"Time"}),
    BUFFER = Table.Buffer(#"Removed Columns"),
    SORTEDDATETIME = Table.Sort(BUFFER,{{"TimeDate", Order.Ascending}}),
    ROLESALL = Table.AddColumn(SORTEDDATETIME, "Role1", each RolesList),
    #"Expanded RoleList" = Table.ExpandListColumn(ROLESALL, "Role1"),
    STAFFCOUNT = Table.AddColumn(#"Expanded RoleList", "StaffCount", each if [Role1] = [Role] then 1 else 0),
    #"Replaced NULL X" = Table.ReplaceValue(STAFFCOUNT,"x",null,Replacer.ReplaceValue,{"Name"}),
    #"Filtered NAMES ONLY" = Table.SelectRows(#"Replaced NULL X", each ([Name] <> null))
in
    #"Filtered NAMES ONLY",
StaffShifts = let
    Source = RosterShiftsWithChecks,
    #"Removed Other Columns" = Table.RemoveColumns(Source, {"Code"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"StartDateTime", "StartOutter"}, {"EndDateTime", "EndOutter"}}),
    STARTOUTTER = Table.AddColumn(#"Renamed Columns", "StartInner", each [StartOutter] + #duration(0,0,1,0)),
    ENDINNER = Table.AddColumn(STARTOUTTER, "EndInner", each [EndOutter]-#duration(0,0,1,0)),
    #"Reordered Columns" = Table.ReorderColumns(ENDINNER,{"Name", "Role",  "StartOutter", "StartInner", "EndOutter", "Duration", "Shift"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"StartInner", type datetime}, {"EndInner", type datetime}}),
    #"Unpivoted Columns" = Table.Unpivot(#"Changed Type", {"StartOutter", "StartInner", "EndOutter", "EndInner"}, "Attribute", "Value"),
    #"Renamed Columns1" = Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "IntervalPoint"}, {"Value", "DateTime"}}),
    STAFFCOUNT = Table.AddColumn(#"Renamed Columns1", "StaffCount", each if [IntervalPoint] = "StartOuter" then 0 else if [IntervalPoint] = "StartInner" then 1 else if [IntervalPoint] = "EndInner" then 1 else 0),
    BUFFER = Table.Buffer(STAFFCOUNT)
in
    BUFFER,
ShiftIndex = let
    Source = StaffShifts,
    #"Removed Columns" = Table.SelectColumns(Source, {"Name", "IntervalPoint", "DateTime"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([IntervalPoint] = "StartOutter")),
    SHIFTINDEX = Table.AddIndexColumn(#"Filtered Rows", "ShiftIndex", 1, 1, Int64.Type)
in
    SHIFTINDEX,
MinHrs = let
    Source = ShiftLength,
    MinHrs = Source{0}[MinHrs]
in
    MinHrs,
ShortHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[ShortHrs]
in
    ShortHrs,
StdHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[StdHrs]
in
    ShortHrs,
#"15OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[1.5OT]
in
    ShortHrs,
#"20OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[2.0OT]
in
    ShortHrs,
ShiftPeriod = let
    Source = #"IMPORT Settings Data",
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type",
ShiftGapIndices = let
    Source = #"EXTRACT Table_StartTime_Duration",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Name", Order.Ascending}, {"StartDateTime", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"StartDateTime", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name"}, {{"Count", each _, type table [Name=text, Role=nullable text, WARD=nullable text, StartDateTime=datetime, EndDateTime=nullable datetime, Duration=number, Shift=text]}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "Custom", each Table.AddIndexColumn([Count],"Index",0,1)),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Count"}),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Removed Columns", "Custom", {"Role", "WARD", "StartDateTime", "EndDateTime", "Duration", "Shift", "Index"}, {"Custom.Role", "Custom.WARD", "Custom.StartDateTime", "Custom.EndDateTime", "Custom.Duration", "Custom.Shift", "Custom.Index"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded Custom",{{"Custom.StartDateTime", type datetime}}),
    #"Grouped Rows1" = Table.Group(#"Changed Type1", {"Name"}, {{"Count2", each _, type table [Name=text, Custom.Role=text, Custom.WARD=text, Custom.StartDateTime=datetime, Custom.EndDateTime=datetime, Custom.Duration=number, Custom.Shift=text, Custom.Index=number]}}),
    #"Added Custom1" = Table.AddColumn(#"Grouped Rows1", "Index2", each Table.AddIndexColumn([Count2],"Index2",1,1)),
    #"Expanded Index2" = Table.ExpandTableColumn(#"Added Custom1", "Index2", {"Custom.Role", "Custom.WARD", "Custom.StartDateTime", "Custom.EndDateTime", "Custom.Duration", "Custom.Shift", "Custom.Index", "Index2"}, {"Index2.Custom.Role", "Index2.Custom.WARD", "Index2.Custom.StartDateTime", "Index2.Custom.EndDateTime", "Index2.Custom.Duration", "Index2.Custom.Shift", "Index2.Custom.Index", "Index2.Index2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded Index2",{"Count2", "Index2.Custom.WARD"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns1",{{"Index2.Custom.StartDateTime", type number}})
in
    #"Changed Type2",
ShiftGap = let
    Source = Table.NestedJoin(ShiftGapIndices, {"Name", "Index2.Index2"}, ShiftGapIndices, {"Name", "Index2.Custom.Index"}, "StaffShifts (2)", JoinKind.LeftOuter),
    #"Sorted Rows" = Table.Sort(Source,{{"Name", Order.Ascending}, {"Index2.Custom.StartDateTime", Order.Ascending}}),
    BUFFER = Table.Buffer(#"Sorted Rows"),
    #"Expanded StaffShifts (2)" = Table.ExpandTableColumn(BUFFER, "StaffShifts (2)", {"Index2.Custom.StartDateTime", "Index2.Custom.Duration"}, {"StaffShifts (2).Index2.Custom.StartDateTime", "StaffShifts (2).Index2.Custom.Duration"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded StaffShifts (2)",{{"Index2.Custom.StartDateTime", type datetime}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{ "Index2.Custom.Shift", "Index2.Custom.Index", "Index2.Index2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns",{{"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type2",null,0,Replacer.ReplaceValue,{"StaffShifts (2).Index2.Custom.Duration"}),
    #"Replaced Value" = Table.ReplaceValue(#"Replaced Value1",null,0,Replacer.ReplaceValue,{"StaffShifts (2).Index2.Custom.StartDateTime"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Replaced Value",{{"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    SHIFTGAP = Table.AddColumn(#"Changed Type3", "Subtraction", each if [#"StaffShifts (2).Index2.Custom.Duration"] = 0
then 99
else

[Index2.Custom.EndDateTime] - [#"StaffShifts (2).Index2.Custom.StartDateTime"]),
    #"Renamed Columns" = Table.RenameColumns(SHIFTGAP,{{"Subtraction", "EndStartDifference"}}),
    BUFFER2 = Table.Buffer(#"Renamed Columns")
in
    BUFFER2,
EffectiveDuration = let
    Source = ShiftGap,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndStartDifference", type number}, {"Index2.Custom.StartDateTime", type datetime}, {"Index2.Custom.EndDateTime", type datetime}, {"StaffShifts (2).Index2.Custom.StartDateTime", type datetime}}),
    DOUBLESHIFT = Table.AddColumn(#"Changed Type", "DoubleShift1", each if [EndStartDifference] = 99 
or 
[EndStartDifference]*-1 >= MinShiftGap/24
then null
else "DoubleShift1"),
    COMBINEDDURATIONS = Table.AddColumn(DOUBLESHIFT, "CombinedDuration", each if [DoubleShift1] = "DoubleShift1" 
then [Index2.Custom.Duration] +  [#"StaffShifts (2).Index2.Custom.Duration"]
else [Index2.Custom.Duration]),
    EFFECTIVEDURATION = Table.AddColumn(COMBINEDDURATIONS, "EffectiveDuration", each if [CombinedDuration] < MealBreakStart/24
then [CombinedDuration] 

else if [CombinedDuration] > 2*MealBreakStart/24
then [CombinedDuration] - 2*MealBreakTime/24

else [CombinedDuration] - MealBreakTime/24),
    #"Added Index" = Table.AddIndexColumn(EFFECTIVEDURATION, "Index", 0, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 1, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index1",{"StaffShifts (2).Index2.Custom.StartDateTime", "StaffShifts (2).Index2.Custom.Duration", "EndStartDifference"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER,
DoubleShifts = let
    Source = EffectiveDuration,
    #"Merged Queries1" = Table.NestedJoin(Source, {"Index"}, Source, {"Index.1"}, "Source", JoinKind.LeftOuter),
    #"Expanded Source" = Table.ExpandTableColumn(#"Merged Queries1", "Source", {"CombinedDuration", "DoubleShift1", "EffectiveDuration"}, {"Source.CombinedDuration", "Source.DoubleShift1", "Source.EffectiveDuration"}),
    #"DOUBLE SHIFT" = Table.AddColumn(#"Expanded Source", "Double Shift", each if [DoubleShift1] = "DoubleShift1" then "DoubleShift1" else if [Source.DoubleShift1] = "DoubleShift1" then "DoubleShift2" else null),
    REALDURATION = Table.AddColumn(#"DOUBLE SHIFT", "RealDuration", each if [Double Shift] = "DoubleShift1" then [CombinedDuration] else if [Double Shift] = "DoubleShift2" then [Source.CombinedDuration] else [CombinedDuration]),
    #"EFFECTIVE SHIFT" = Table.AddColumn(REALDURATION, "Effective Duration", each if [Double Shift] = "DoubleShift1" then [EffectiveDuration] else if [Double Shift] = "DoubleShift2" then [Source.EffectiveDuration] else [EffectiveDuration])
in
    #"EFFECTIVE SHIFT",
StaffDoubleShifts = let
    Source = Table.NestedJoin(StaffShifts, {"Name", "DateTime"}, ShiftIndex, {"Name", "DateTime"}, "EffectiveShift.1", JoinKind.LeftOuter),
    #"Expanded EffectiveShift.2" = Table.ExpandTableColumn(Source, "EffectiveShift.1", {"DateTime"}, {"EffectiveShift.1.DateTime"}),
    #"Sorted Rows" = Table.Sort(#"Expanded EffectiveShift.2",{{"Name", Order.Ascending}, {"DateTime", Order.Ascending}}),
    #"Filled Down" = Table.FillDown(#"Sorted Rows",{"EffectiveShift.1.DateTime"}),
    #"Merged Queries" = Table.NestedJoin(#"Filled Down", {"Name", "EffectiveShift.1.DateTime"}, DoubleShifts, {"Name", "Index2.Custom.StartDateTime"}, "DoubleShifts", JoinKind.LeftOuter),
    #"Expanded DoubleShifts" = Table.ExpandTableColumn(#"Merged Queries", "DoubleShifts", {"Double Shift", "Effective Duration", "RealDuration"}, {"Double Shift", "Effective Duration", "RealDuration"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded DoubleShifts",{"EffectiveShift.1.DateTime"}),
    SHIFTTYPE = Table.AddColumn(#"Removed Columns", "ShiftType", each if [RealDuration] <= 2/24 then "0-4Hrs" else if [RealDuration] <= 7.5/24 then "4-7.5Hrs" else if [RealDuration] <= 8.5/24 then "7.5-8.5Hrs" else if [RealDuration] <= 10/24 then "8.5-10Hrs" else if [RealDuration] <= 12/24 then "10-12Hrs" else ">12Hrs")
in
    SHIFTTYPE,
AMP = let
    Source = ShiftPeriod,
    StartDay = Source{0}[StartDay]
in
    StartDay,
PMP = let
    Source = ShiftPeriod,
    StartDay = Source{1}[StartDay]
in
    StartDay,
NIGHTP = let
    Source = ShiftPeriod,
    StartDay = Source{2}[StartDay]
in
    StartDay,
RolesList = let
    Source = #"LINK Role Analysis",
    ROLES = Source[ROLES]
in
    ROLES,
ShiftPoints = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCs7ITCsJLkksKlGK1YFyXfNSlGJjAQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ShiftPoints = _t]),
    ShiftPoints1 = Source[ShiftPoints]
in
    ShiftPoints1,
#"EXTRACT Table_StartTime_Duration" = let

    Source = #"IMPORT Intervals",
    Table_StartTime_Duration_Table = Source{[Item="Table_StartTime_Duration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_StartTime_Duration_Table,{{"Date", type date}, {"Name", type text}, {"Role", type text}, {"Start", type number}, {"End", type number}, {"StartDateTime", type datetime}, {"EndDateTime", type datetime}, {"Duration", type number}, {"Shift", type text}, {"EffectiveDuration", type number}})
in
    #"Changed Type",
#"EXTRACT Table Intervals" = let

    Source = #"IMPORT Intervals",
    Intervals_Table = Source{[Item="Intervals",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(Intervals_Table,{{"StartInterval", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}, {"EndInterval", type datetime}, {"EndPrecise", type datetime}, {"Duration", type number}, {"Role", type text}, {"ShiftPeriod", type text}, {"IntervalAssociatedShift", type text}})
in
    #"Changed Type1",
#"EXTRACT Table_IntervalsList" = let

    Source = #"IMPORT Intervals",
    Table_IntervalsList_Table = Source{[Item="IntervalsList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_IntervalsList_Table,{{"IntervalID", Int64.Type}, {"DayDate", type date}, {"TimeDate", type datetime}, {"Time", type number}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Attribute", type text}, {"ShiftDate", type date}})
in
    #"Changed Type",
#"IMPORT Settings Data" = Inputs[SettingsNavigation],
#"EXTRACT MealBreak" = let

    Source = #"IMPORT Settings Data",
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type",
#"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type",
#"MANUAL Shifts" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Shifts = _t, Order = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}})
in
    #"Changed Type",
#"LINK Role Analysis" = Inputs[Roles],
Table_Intervals = let
        Source1 = #"EXTRACT Table Intervals",
    #"Removed Columns" = Table.RemoveColumns(Source1,{"IntervalAssociatedShift"})
in
    #"Removed Columns",
#"Table Intervals-IntervalAssocShift" = let
    Source = #"EXTRACT Table Intervals"
in
    Source,
#"IMPORT OvertimeRosterChecks" = Inputs[Checks],
RosterShiftsWithChecks = let
    Source = #"EXTRACT Table_StartTime_Duration",
    PlotInputs = Table.SelectColumns(Source, {"Code", "Name", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    WithLocalRowNumber = Table.AddIndexColumn(PlotInputs, "LocalShiftRowNumber", 0, 1, Int64.Type),
    LocalUnit = Text.From(UnitL1PathTABLE{[#"Variable Name"="Unit"]}[Value]),
    WithSourceUnit = Table.AddColumn(WithLocalRowNumber, "SourceUnit", each LocalUnit, type text),
    WithAllocationCode = Table.AddColumn(WithSourceUnit, "AllocationCode", each if [Code] = null then null else Text.Trim(Text.From([Code], "en-AU")), type nullable text),
    WithShiftStart = Table.AddColumn(WithAllocationCode, "ShiftStart", each [StartDateTime], type datetime),
    WithShiftFinish = Table.AddColumn(WithShiftStart, "ShiftFinish", each [EndDateTime], type datetime),
    // This execution-local row identity also exists for unmatched Master Roster rows. SourceShiftID
    // remains the stable published-duty identity used across workbooks and payroll counters.
    WithLocalShiftKey = Table.AddColumn(WithShiftFinish, "LocalShiftKey", each Binary.ToText(Json.FromValue({[SourceUnit], [AllocationCode], [Role], [ShiftStart], [ShiftFinish], [LocalShiftRowNumber]}), BinaryEncoding.Base64), type text),
    RemovedLocalRowNumber = Table.RemoveColumns(WithLocalShiftKey, {"LocalShiftRowNumber"}),
    CoreSource = #"IMPORT OvertimeRosterChecks",
    MatchKeys = {"SourceUnit", "AllocationCode", "Role", "ShiftStart", "ShiftFinish"},
    RequiredCoreColumns = List.Combine({MatchKeys, {"SourceShiftID"}}),
    CoreContract = if Table.HasColumns(CoreSource, RequiredCoreColumns) then CoreSource
        else error "OvertimeRosterChecks_ShiftResults is missing its required shift identity columns.",
    // Source provenance stays in the core helper rather than repeating wide text on plotting/interval rows.
    CoreAttributes = List.RemoveItems(Table.ColumnNames(CoreContract), List.Combine({MatchKeys, {"SourceFile", "SourceRow", "SourceName", "SourceDate"}})),
    CoreColumns = List.Combine({MatchKeys, CoreAttributes}),
    NarrowCoreResults = Table.SelectColumns(CoreContract, CoreColumns),
    CoreResults = Table.Buffer(NarrowCoreResults),
    JoinedPublishedShifts = Table.NestedJoin(RemovedLocalRowNumber, MatchKeys, CoreResults, MatchKeys, "PublishedShiftMatches", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(JoinedPublishedShifts, "OvertimeMatchCount", each if [AllocationCode] = null or [AllocationCode] = "" or [Role] = null
        then 0 else Table.RowCount([PublishedShiftMatches]), Int64.Type),
    WithMatchStatus = Table.AddColumn(WithMatchCount, "OvertimeMatchStatus", each if [OvertimeMatchCount] = 1 then "Matched"
        else if [OvertimeMatchCount] = 0 then "Unmatched" else "Ambiguous", type text),
    WithMatchReason = Table.AddColumn(WithMatchStatus, "OvertimeMatchReason", each if [OvertimeMatchCount] = 1 then null
        else if [OvertimeMatchCount] = 0 then "No uniquely identified published shift; existing plotting row retained."
        else "Multiple published shifts match this unit, employee code, role and original boundaries.", type nullable text),
    UnknownResult = Record.FromList(List.Transform(CoreAttributes, each if Text.EndsWith(_, "Status") then "Unknown" else null), CoreAttributes),
    // Select a record only after uniqueness is established; duplicate matches never multiply output rows.
    WithPublishedResult = Table.AddColumn(WithMatchReason, "PublishedShiftResult", each if [OvertimeMatchCount] = 1
        then Record.SelectFields([PublishedShiftMatches]{0}, CoreAttributes) else UnknownResult, type record),
    RemovedNestedMatches = Table.RemoveColumns(WithPublishedResult, {"PublishedShiftMatches"}),
    ExpandedCheckAttributes = Table.ExpandRecordColumn(RemovedNestedMatches, "PublishedShiftResult", CoreAttributes, CoreAttributes),
    NumericAttributes = {"GrossHours", "BreakMinutes", "NetRosterHours", "QualifyingNetHours", "GapHours", "GapWallHours", "DailyHoursLimit", "DailyExcessHours", "PayrollFortnightCount", "FortnightHoursBefore", "FortnightHoursAfter", "FortnightExcessHours", "DutyFreeDays", "DaysShort", "OvertimeHours"},
    LogicalAttributes = List.Select(CoreAttributes, each Text.EndsWith(_, "Flag") or List.Contains({"OvertimeExempt", "BrokenShiftException"}, _)),
    TextAttributes = List.Select(CoreAttributes, each Text.EndsWith(_, "Status") or List.Contains({"SourceShiftID", "EmployeeID", "WorkerEmployeeID", "EmployeeKey", "EmployerKey", "SourceRole", "RosterShiftPeriod", "AgreementStream", "Grade", "EmploymentType", "PreviousSourceShiftID", "TriggerSummary", "ResultReasons"}, _)),
    DateAttributes = List.Intersect({CoreAttributes, {"PayrollFortnightStart", "PayrollFortnightEnd"}}),
    AttributeTypes = List.Combine({List.Transform(List.Intersect({CoreAttributes, NumericAttributes}), each {_, type number}),
        List.Transform(LogicalAttributes, each {_, type logical}), List.Transform(TextAttributes, each {_, type text}),
        List.Transform(DateAttributes, each {_, type date}),
        if List.Contains(CoreAttributes, "PreviousShiftFinish") then {{"PreviousShiftFinish", type datetime}} else {}}),
    TypedCheckAttributes = Table.TransformColumnTypes(ExpandedCheckAttributes, AttributeTypes, "en-AU")
in
    TypedCheckAttributes,
LegacyShiftMeasures = let
    LegacyColumns = {"Double Shift", "Effective Duration", "RealDuration", "ShiftType"},
    ClassificationColumns = {"Effective Duration", "RealDuration", "ShiftType"},
    Source = StaffDoubleShifts,
    NarrowLegacyPoints = Table.SelectColumns(Source, List.Combine({{"LocalShiftKey", "Name", "IntervalPoint"}, LegacyColumns})),
    ClassificationPoints = Table.SelectColumns(NarrowLegacyPoints, List.Combine({{"LocalShiftKey", "Name"}, ClassificationColumns})),
    DistinctLegacyResults = Table.Distinct(ClassificationPoints),
    GroupedShiftResults = Table.Group(DistinctLegacyResults, {"LocalShiftKey", "Name"}, {{"LegacyResultCount", each Table.RowCount(_), Int64.Type}, {"CandidateMeasures", each Table.SelectColumns(_, ClassificationColumns), type table}}),
    NameIdentities = Table.SelectColumns(RosterShiftsWithChecks, {"Name", "AllocationCode"}),
    NameEmployeeCounts = Table.Group(NameIdentities, {"Name"}, {{"NameIdentityKnown", each
        let Codes = List.Distinct([AllocationCode]), DisplayName = _{0}[Name]
        in List.Count(Codes) = 1 and Codes{0} <> null and Codes{0} <> "" and DisplayName <> null and Text.Trim(DisplayName) <> "", type logical}}),
    JoinedNameIdentities = Table.NestedJoin(GroupedShiftResults, {"Name"}, NameEmployeeCounts, {"Name"}, "NameIdentity", JoinKind.LeftOuter),
    ExpandedIdentityCounts = Table.ExpandTableColumn(JoinedNameIdentities, "NameIdentity", {"NameIdentityKnown"}),
    WithMeasureValidity = Table.AddColumn(ExpandedIdentityCounts, "LegacyMeasuresValid", each if [LegacyResultCount] <> 1 then false else
        let Candidate = [CandidateMeasures]{0}, Real = Candidate[RealDuration], Effective = Candidate[#"Effective Duration"], Band = Candidate[ShiftType]
        in Value.Is(Real, type number) and Value.Is(Effective, type number) and not Number.IsNaN(Real) and not Number.IsNaN(Effective)
            and Number.Abs(Real) <> #infinity and Number.Abs(Effective) <> #infinity and Real >= 0 and Effective >= 0
            and Value.Is(Band, type text) and Text.Trim(Band) <> "", type logical),
    WithLegacyStatus = Table.AddColumn(WithMeasureValidity, "LegacyShiftStatus", each if [LegacyResultCount] = 1 and [NameIdentityKnown] and [LegacyMeasuresValid] then "Known" else "Unknown", type text),
    WithLegacyReason = Table.AddColumn(WithLegacyStatus, "LegacyShiftReason", each if not [NameIdentityKnown] then "Legacy Name grouping does not identify exactly one nonblank employee code."
        else if [LegacyResultCount] <> 1 then "Legacy plotting points disagree about this physical shift's classification."
        else if not [LegacyMeasuresValid] then "Legacy duration or classification is invalid." else null, type nullable text),
    UnknownMeasures = Record.FromList(List.Repeat({null}, List.Count(ClassificationColumns)), ClassificationColumns),
    WithUniqueMeasures = Table.AddColumn(WithLegacyReason, "UniqueMeasures", each if [LegacyShiftStatus] = "Known" then [CandidateMeasures]{0} else UnknownMeasures, type record),
    ExpandedUniqueMeasures = Table.ExpandRecordColumn(WithUniqueMeasures, "UniqueMeasures", ClassificationColumns, ClassificationColumns),
    // A shared end/start timestamp may have another point's Double Shift marker. Keep the
    // physical shift's own outer-start marker separate from its whole-shift classification.
    OwnStartPoints = Table.SelectRows(NarrowLegacyPoints, each [IntervalPoint] = "StartOutter"),
    OwnStartMarkers = Table.Group(OwnStartPoints, {"LocalShiftKey"}, {{"OwnDoubleShiftMarkers", each List.Distinct([#"Double Shift"]), type list}}),
    JoinedStartMarkers = Table.NestedJoin(ExpandedUniqueMeasures, {"LocalShiftKey"}, OwnStartMarkers, {"LocalShiftKey"}, "StartMarker", JoinKind.LeftOuter),
    ExpandedStartMarkers = Table.ExpandTableColumn(JoinedStartMarkers, "StartMarker", {"OwnDoubleShiftMarkers"}),
    WithDoubleShiftMarker = Table.AddColumn(ExpandedStartMarkers, "Double Shift", each if [LegacyShiftStatus] = "Known" and List.Count([OwnDoubleShiftMarkers]) = 1 then [OwnDoubleShiftMarkers]{0} else null, type nullable text),
    PublishedMeasures = Table.SelectColumns(WithDoubleShiftMarker, List.Combine({{"LocalShiftKey"}, LegacyColumns, {"LegacyShiftStatus", "LegacyShiftReason"}}))
in
    PublishedMeasures,
OvertimeRosterChecks = let
    Source = StaffShifts,
    JoinedLegacyClassification = Table.NestedJoin(Source, {"LocalShiftKey"}, LegacyShiftMeasures, {"LocalShiftKey"}, "LegacyClassification", JoinKind.LeftOuter),
    ExpandedLegacyClassification = Table.ExpandTableColumn(JoinedLegacyClassification, "LegacyClassification", {"ShiftType", "LegacyShiftStatus", "LegacyShiftReason"}),
    PlottingColumns = {"Name", "Role", "DateTime", "StaffCount", "ShiftType"},
    CheckColumns = List.RemoveItems(Table.ColumnNames(RosterShiftsWithChecks), {"Code", "Name", "Role", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    PublishedColumns = List.Combine({PlottingColumns, CheckColumns, {"LegacyShiftStatus", "LegacyShiftReason"}}),
    SelectedPlotAndChecks = Table.SelectColumns(ExpandedLegacyClassification, PublishedColumns),
    TypedPlotColumns = Table.TransformColumnTypes(SelectedPlotAndChecks, {{"Name", type text}, {"Role", type text}, {"DateTime", type datetime}, {"StaffCount", Int64.Type}, {"ShiftType", type text}})
in
    TypedPlotColumns,
IntervalsList = let
    Source = #"EXTRACT Table_IntervalsList"
in
    Source,
#"NAMES+INTERVALSMATRIX" = let
    Source = #"NAMES+INTERVALS",
    #"Filtered Rows1" = Table.SelectRows(Source, each [IntervalDurationTemp] > 0.1875),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"ShiftPeriod"}, #"MANUAL Shifts", {"Shifts"}, "MANUAL Shifts", JoinKind.LeftOuter),
    #"Expanded MANUAL Shifts" = Table.ExpandTableColumn(#"Merged Queries", "MANUAL Shifts", {"Order"}, {"Order"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded MANUAL Shifts", each ([Name] <> "x")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"DayDate", "Role", "IntervalStart", "Name", "ShiftPeriod", "Order"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"DayDate", "ShiftPeriod", "Order", "Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows1" = Table.Sort(#"Grouped Rows",{{"DayDate", Order.Ascending}, {"Order", Order.Ascending}}),
    BUFFER = Table.Buffer(#"Sorted Rows1"),
    #"Merged Queries1" = Table.NestedJoin(BUFFER, {"DayDate", "ShiftPeriod", "Role"}, #"EXTRACT PermutationDimensions", {"Date", "Shifts", "RolesList"}, "IMPORT PermutationDimensions", JoinKind.FullOuter),
    #"Expanded IMPORT PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries1", "IMPORT PermutationDimensions", {"Date", "RolesList", "Shifts"}, {"Date", "RolesList", "Shifts"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded IMPORT PermutationDimensions",{{"ShiftPeriod", "ShiftPeriodX"}, {"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ShiftPeriod", each if [ShiftPeriodX] = null then [Shifts] else [ShiftPeriodX]),
    #"Added ROLE" = Table.AddColumn(#"Added Conditional Column", "Role", each if [RoleX] = null then [RolesList] else [RoleX]),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Added ROLE", {{"DayDate", type text}}, "en-AU"),{"DayDate", "ShiftPeriod"},Combiner.CombineTextByDelimiter(":", QuoteStyle.None),"Merged"),
    #"Added Custom" = Table.AddColumn(#"Merged Columns", "Allocated", each "X"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Order", "Count", "ShiftPeriodX", "RoleX", "RolesList", "Shifts", "Date"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Merged]), "Merged", "Allocated")
in
    #"Pivoted Column1"
in [Points = StaffDoubleShifts, Intervals = #"NAMES+INTERVALSLIST", Prepared = RosterShiftsWithChecks, Published = OvertimeRosterChecks],
EvaluateConsumer = (Rows as table) as table => let #"EXTRACT NAMES_INTERVALSLIST" = Rows, ResourceIntervalAllocation = let
    Source = #"EXTRACT NAMES_INTERVALSLIST",
    #"Filtered Rows" = Table.SelectRows(Source, each ([RealDuration] <> 0)),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"IntervalDurationTemp", "IntervalDuration"}}),
    RESEFFECTIVERATIO = Table.AddColumn(#"Renamed Columns", "ResEffectiveRatio-FTE", each if [StaffCount] = 0 then 0
else [Effective Duration] / [RealDuration]),
    EFFECTIVEINTERVALEFFORT = Table.AddColumn(RESEFFECTIVERATIO, "ResEffectiveIntervalEffort", each [#"ResEffectiveRatio-FTE"]*[IntervalDuration]),
    #"Removed Columns" = Table.RemoveColumns(EFFECTIVEINTERVALEFFORT,{ "TimeType"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"TimeDate", type datetime}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"ShiftDate", type date}}),
    BUFFER = Table.Buffer(#"Changed Type1")
in
    BUFFER in ResourceIntervalAllocation,
Fixtures = (EvaluateBaseline as function, EvaluateNew as function, EvaluateConsumer as function) as table =>
let
    RosterShift = (code as number, name as text, role as text, start as datetime, finish as datetime, shift as text) as record =>
        [Code = code, Name = name, Role = role, Date = Date.From(start), Start = Number.From(Time.From(start)),
            End = Number.From(Time.From(finish)), StartDateTime = start, EndDateTime = finish,
            Duration = Duration.TotalDays(finish - start), Shift = shift,
            EffectiveDuration = Duration.TotalDays(finish - start) - 0.5 / 24],
    RosterRecords = {
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 18, 0, 0), "AM"),
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 3, 14, 0, 0), #datetime(2026, 11, 3, 22, 0, 0), "PM"),
        RosterShift(101, "Ava", "AIN", #datetime(2026, 11, 4, 22, 0, 0), #datetime(2026, 11, 5, 6, 0, 0), "NIGHT"),
        RosterShift(202, "Ben", "RN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(900, "Master unmatched", "AIN", #datetime(2026, 11, 4, 6, 0, 0), #datetime(2026, 11, 4, 14, 0, 0), "AM")
    },
    Roster = Table.FromRecords(RosterRecords),
    IntervalTable = (roster as table) as table =>
        let
            Boundaries = List.Sort(List.Distinct(List.Combine(List.Transform(Table.ToRecords(roster), each {[StartDateTime], [EndDateTime]})))),
            IntervalRecords = List.Combine(List.Transform({0..List.Count(Boundaries) - 2}, (index) =>
                List.Transform({"AIN", "RN"}, (role) =>
                    [StartInterval = Boundaries{index}, EndPrecise = Boundaries{index + 1},
                        EndInterval = Boundaries{index + 1} - #duration(0, 0, 1, 0),
                        DayDate = Date.From(Boundaries{index}), IntervalID = index + 1,
                        Duration = Duration.TotalDays(Boundaries{index + 1} - Boundaries{index}),
                        Role = role, ShiftPeriod = "AM", IntervalAssociatedShift = "AM"])))
        in Table.FromRecords(IntervalRecords),
    CheckRecord = (row as record, ordinal as number) as record =>
        [SourceUnit = "BD", AllocationCode = Text.From(row[Code]), Role = row[Role],
            ShiftStart = row[StartDateTime], ShiftFinish = row[EndDateTime], SourceShiftID = "S" & Text.From(ordinal),
            EmployeeID = Text.From(row[Code]), RosterShiftPeriod = row[Shift],
            DailyStatus = if ordinal = 1 then "Flag" else "Pass", DailyExcessHours = if ordinal = 1 then 1.5 else 0,
            MinimumBreakFlag = if ordinal = 3 then null else false,
            GapStatus = if ordinal = 3 then "Unknown" else "Known",
            CoverageStatus = if ordinal = 4 then "Unknown" else "Confirmed",
            SourceFile = "synthetic-source", SourceRow = ordinal, SourceName = row[Name], SourceDate = row[Date]],
    CheckRecords = List.Transform({0..3}, each CheckRecord(RosterRecords{_}, _ + 1)),
    Checks = Table.FromRecords(CheckRecords),
    Inputs = (checks as table, optional suppliedRoster as nullable table) as record => [
        Roster = if suppliedRoster = null then Roster else suppliedRoster,
        Checks = checks,
        PathTable = #table({"Variable Name", "Value"}, {{"Unit", "BD"}, {"Root Path", "fixture/2. Calculations"}}),
        IntervalsNavigation = #table({"Item", "Kind", "Data"}, {
            {"Table_StartTime_Duration", "Table", if suppliedRoster = null then Roster else suppliedRoster},
            {"Intervals", "Table", IntervalTable(if suppliedRoster = null then Roster else suppliedRoster)}}),
        SettingsNavigation = #table({"Item", "Kind", "Data"}, {
            {"Meals", "Table", #table({"MealBreakTime", "MealBreakstart"}, {{0.5, 5}})},
            {"ShiftGap", "Table", #table({"MinGap"}, {{8}})},
            {"ShiftLent", "Table", #table({"MinHrs", "ShortHrs", "StdHrs", "1.5OT", "2.0OT"}, {{4, 6, 8, 1.5, 2}})},
            {"ShiftPeriod", "Table", #table({"ShiftPeriod", "StartTime", "StartDay"}, {{"AM", 6 / 24, 6 / 24}, {"PM", 14 / 24, 14 / 24}, {"NIGHT", 22 / 24, 22 / 24}})}}),
        Roles = #table({"ROLES"}, {{"AIN"}, {"RN"}})
    ],
    Baseline = EvaluateBaseline(Inputs(Checks)),
    Current = EvaluateNew(Inputs(Checks)),
    PointColumns = {"Name", "Role", "Duration", "Shift", "IntervalPoint", "DateTime", "StaffCount", "Double Shift", "Effective Duration", "RealDuration", "ShiftType"},
    PointSort = {{"Name", Order.Ascending}, {"DateTime", Order.Ascending}, {"IntervalPoint", Order.Ascending}, {"Role", Order.Ascending}},
    IntervalSort = {{"Name", Order.Ascending}, {"Role", Order.Ascending}, {"TimeDate", Order.Ascending}, {"AllocatedIntervals", Order.Ascending}, {"Attribute", Order.Ascending}, {"Role1", Order.Ascending}},
    EqualProjection = (left as table, right as table, columns as list, sort as list) as logical =>
        Table.ToRows(Table.Sort(Table.SelectColumns(left, columns), sort)) = Table.ToRows(Table.Sort(Table.SelectColumns(right, columns), sort)),
    OldIntervals = Baseline[Intervals],
    NewIntervals = Current[Intervals],
    OldConsumer = EvaluateConsumer(OldIntervals),
    NewConsumer = EvaluateConsumer(NewIntervals),
    Metadata = {"SourceUnit", "AllocationCode", "ShiftStart", "ShiftFinish", "SourceShiftID", "EmployeeID", "RosterShiftPeriod",
        "DailyStatus", "DailyExcessHours", "MinimumBreakFlag", "GapStatus", "CoverageStatus", "OvertimeMatchCount", "OvertimeMatchStatus", "OvertimeMatchReason"},
    SourcePoints = (sourceID as text) as table => Table.SelectRows(Current[Points], each [SourceShiftID] = sourceID),
    SourceIntervals = (sourceID as text) as table => Table.SelectRows(NewIntervals, each [SourceShiftID] = sourceID),
    WithAmbiguity = Table.FromRecords(CheckRecords & {Record.TransformFields(CheckRecords{0}, {{"SourceShiftID", each "duplicate"}})}),
    Ambiguous = EvaluateNew(Inputs(WithAmbiguity)),
    OverlapRows = {
        RosterShift(301, "Shared Name", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(302, "Shared Name", "AIN", #datetime(2026, 11, 2, 7, 0, 0), #datetime(2026, 11, 2, 15, 0, 0), "AM")},
    SameStartRows = {
        RosterShift(401, "Identical Start", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 14, 0, 0), "AM"),
        RosterShift(402, "Identical Start", "AIN", #datetime(2026, 11, 2, 6, 0, 0), #datetime(2026, 11, 2, 15, 0, 0), "AM")},
    AdjacentRows = {
        RosterShift(501, "Adjacent worker", "AIN", #datetime(2026, 11, 6, 6, 0, 0), #datetime(2026, 11, 6, 10, 0, 0), "AM"),
        RosterShift(501, "Adjacent worker", "AIN", #datetime(2026, 11, 6, 11, 0, 0), #datetime(2026, 11, 6, 15, 0, 0), "AM")},
    OverlapChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(OverlapRows{_}, _ + 1))),
    SameStartChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(SameStartRows{_}, _ + 1))),
    Overlap = EvaluateNew(Inputs(OverlapChecks, Table.FromRecords(OverlapRows))),
    SameStart = EvaluateNew(Inputs(SameStartChecks, Table.FromRecords(SameStartRows))),
    AdjacentChecks = Table.FromRecords(List.Transform({0..1}, each CheckRecord(AdjacentRows{_}, _ + 1))),
    Adjacent = EvaluateNew(Inputs(AdjacentChecks, Table.FromRecords(AdjacentRows))),
    AdjacentBaseline = EvaluateBaseline(Inputs(AdjacentChecks, Table.FromRecords(AdjacentRows))),
    CanonicalPointCounts = (result as record) as logical =>
        let grouped = Table.Group(result[Published], {"SourceShiftID"}, {{"Count", each Table.RowCount(_), Int64.Type}})
        in Table.RowCount(result[Published]) = 8 and Table.RowCount(grouped) = 2 and List.Distinct(grouped[Count]) = {4},
    IntervalsStayWithinOwnShift = (result as record, sourceRows as list) as logical =>
        let
            Actual = EvaluateConsumer(result[Intervals]),
            WithinBounds = Table.RowCount(Table.SelectRows(Actual, each [TimeDate] < [ShiftStart] or [TimeDate] >= [ShiftFinish])) = 0,
            Starts = Table.SelectRows(Actual, each [Attribute] = "IntervalStart" and [Role1] = [Role]),
            Totals = Table.Group(Starts, {"SourceShiftID"}, {{"ElapsedHours", each List.Sum([IntervalDuration]) * 24, type number}}),
            OriginalHours = List.Transform(sourceRows, each Duration.TotalHours([EndDateTime] - [StartDateTime]))
        in WithinBounds and Table.RowCount(Totals) = 2 and List.Sort(Totals[ElapsedHours]) = List.Sort(OriginalHours),
    AlterFirstKey = (key as text, replacement as any) as table => Table.FromRecords({Record.TransformFields(CheckRecords{0}, {{key, each replacement}})} & List.Skip(CheckRecords)),
    Check = (name as text, run as function, predicate as function) as record =>
        let
            Attempt = try run(),
            Test = if Attempt[HasError] then [HasError = true, Error = Attempt[Error]] else try predicate(Attempt[Value]),
            Passed = not Test[HasError] and Test[Value] = true,
            Details = if Test[HasError] then try Test[Error][Message] otherwise "Evaluation failed"
                else if Passed then null else "Unexpected fixture result"
        in [Case = name, Passed = Passed, Details = Details],
    Cases = {
        Check("Five source shifts retain exactly four plotting points each", () => Current[Points], each Table.RowCount(_) = 20),
        Check("All eleven legacy plotting columns preserve exact baseline values", () => Current[Points], each EqualProjection(Baseline[Points], _, PointColumns, PointSort)),
        Check("SourceShiftID and checks copied to all four points", () => {1..4}, each List.AllTrue(List.Transform(_, (index) => Table.RowCount(SourcePoints("S" & Text.From(index))) = 4))),
        Check("Different shifts of same employee retain independent statuses", () => {SourcePoints("S1"), SourcePoints("S2")}, each
            List.Distinct(_{0}[DailyStatus]) = {"Flag"} and List.Distinct(_{1}[DailyStatus]) = {"Pass"}
            and List.Distinct(_{0}[DailyExcessHours]) = {1.5} and List.Distinct(_{1}[DailyExcessHours]) = {0}),
        Check("Same-time different employees retain independent status", () => SourcePoints("S4"), each List.Distinct(_[CoverageStatus]) = {"Unknown"} and List.Distinct(_[DailyStatus]) = {"Pass"}),
        Check("Nullable flags are preserved without fill-down", () => SourcePoints("S3"), each List.Distinct(_[MinimumBreakFlag]) = {null} and List.Distinct(_[GapStatus]) = {"Unknown"}),
        Check("Unmatched Master row stays as four unknown plotting points", () => Table.SelectRows(Current[Points], each [Name] = "Master unmatched"), each
            Table.RowCount(_) = 4 and List.Distinct(_[OvertimeMatchStatus]) = {"Unmatched"}
            and List.Distinct(_[SourceShiftID]) = {null} and List.Distinct(_[DailyStatus]) = {"Unknown"}
            and List.Distinct(_[DailyExcessHours]) = {null}),
        Check("Ambiguous match retains one prepared row with unknown checks", () => Table.SelectRows(Ambiguous[Prepared], each [AllocationCode] = "101" and [ShiftStart] = #datetime(2026, 11, 2, 6, 0, 0)), each
            Table.RowCount(_) = 1 and _{0}[OvertimeMatchCount] = 2 and _{0}[OvertimeMatchStatus] = "Ambiguous"
            and _{0}[SourceShiftID] = null and _{0}[DailyStatus] = "Unknown" and _{0}[DailyExcessHours] = null),
        Check("Ambiguous match causes no plotting or interval row growth", () => Ambiguous, each
            Table.RowCount(_[Points]) = Table.RowCount(Baseline[Points]) and Table.RowCount(_[Intervals]) = Table.RowCount(OldIntervals)),
        Check("NAMES+INTERVALSLIST legacy keys and values preserve baseline", () => NewIntervals, each
            EqualProjection(OldIntervals, _, Table.ColumnNames(OldIntervals), IntervalSort)),
        Check("NAMES+INTERVALSLIST carries all new checks before endpoint and Role1 expansion", () => NewIntervals, each Table.HasColumns(_, Metadata)),
        Check("Interval rows preserve status by source shift without fill-down", () => {SourceIntervals("S1"), SourceIntervals("S2"), SourceIntervals("S3")}, each
            List.AllTrue(List.Transform(_, (rows) => Table.RowCount(rows) > 0))
            and List.Distinct(_{0}[DailyStatus]) = {"Flag"} and List.Distinct(_{1}[DailyStatus]) = {"Pass"}
            and List.Distinct(_{2}[MinimumBreakFlag]) = {null}),
        Check("Actual ResourceIntervalAllocation keeps all old metrics and row grain", () => NewConsumer, each
            EqualProjection(OldConsumer, _, Table.ColumnNames(OldConsumer), IntervalSort)),
        Check("Actual ResourceIntervalAllocation retains added check columns", () => NewConsumer, each Table.HasColumns(_, Metadata)),
        Check("Published OvertimeRosterChecks exposes five plotting fields and metadata", () => Current[Published], each
            List.FirstN(Table.ColumnNames(_), 5) = {"Name", "Role", "DateTime", "StaffCount", "ShiftType"}
            and Table.RowCount(_) = 20 and Table.HasColumns(_, Metadata)),
        Check("Wide source provenance is not repeated on plotting or interval rows", () => Current, each
            List.IsEmpty(List.Intersect({Table.ColumnNames(_[Points]), {"SourceFile", "SourceRow", "SourceName", "SourceDate"}}))
            and List.IsEmpty(List.Intersect({Table.ColumnNames(_[Intervals]), {"SourceFile", "SourceRow", "SourceName", "SourceDate"}}))),
        Check("Same display name overlapping distinct employees retain four canonical points each", () => Overlap, CanonicalPointCounts),
        Check("Same display name overlapping intervals retain each source shift's own bounds and elapsed hours", () => Overlap, each IntervalsStayWithinOwnShift(_, OverlapRows)),
        Check("Same display name legacy classifications remain visibly unknown", () => Overlap[Published], each
            List.Distinct(_[LegacyShiftStatus]) = {"Unknown"} and List.Distinct(_[ShiftType]) = {null}),
        Check("Same display name with identical starts causes no canonical point multiplication", () => SameStart, CanonicalPointCounts),
        Check("Same display name with identical starts retains exact original interval bounds", () => SameStart, each IntervalsStayWithinOwnShift(_, SameStartRows)),
        Check("Same-name distinct employee check results are not cross-attributed", () => Overlap[Published], each
            List.Distinct(Table.SelectRows(_, (row) => row[AllocationCode] = "301")[DailyStatus]) = {"Flag"}
            and List.Distinct(Table.SelectRows(_, (row) => row[AllocationCode] = "302")[DailyStatus]) = {"Pass"}),
        Check("Adjacent double-shift legs retain known classification", () => Adjacent[Published], each
            Table.RowCount(_) = 8 and List.Distinct(_[LegacyShiftStatus]) = {"Known"}),
        Check("Adjacent double-shift legs preserve baseline plotting values and bands", () => Adjacent[Published], each
            EqualProjection(AdjacentBaseline[Points], _, {"Name", "Role", "DateTime", "StaffCount", "ShiftType"},
                {{"Name", Order.Ascending}, {"Role", Order.Ascending}, {"DateTime", Order.Ascending}}))
    },
    KeyCases = List.Transform({
        [Key = "SourceUnit", Value = "TE"], [Key = "AllocationCode", Value = "other"], [Key = "Role", Value = "RN"],
        [Key = "ShiftStart", Value = #datetime(2026, 11, 2, 6, 1, 0)], [Key = "ShiftFinish", Value = #datetime(2026, 11, 2, 18, 1, 0)]},
        (change) => Check("Exact five-part key rejects changed " & change[Key],
            () => EvaluateNew(Inputs(AlterFirstKey(change[Key], change[Value])))[Prepared],
            each let first = Table.SelectRows(_, (row) => row[Name] = "Ava" and row[StartDateTime] = #datetime(2026, 11, 2, 6, 0, 0)){0}
                in first[OvertimeMatchCount] = 0 and first[OvertimeMatchStatus] = "Unmatched" and first[DailyStatus] = "Unknown")),
    Output = Table.FromRecords(Cases & KeyCases)
in
    Output
in Fixtures(EvaluateBaseline, EvaluateNew, EvaluateConsumer);
