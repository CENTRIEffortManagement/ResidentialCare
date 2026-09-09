// Power Query from: ShiftsRFBI.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\ShiftsRFBI.xlsx
// Extracted: 2026-09-09T01:30:51.700Z

section Section1;

shared Table_AllocationExtracted = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 &  "\1. Input\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="Table_AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Date", type date}, {"Start", type time}, {"End", type time}, {"Role", type text}, {"Unit", type text}}),
    BUFFEER = Table.Buffer(#"Changed Type")
in
    BUFFEER;

shared MealBreak = let
    
    
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 & "\2. Calculations\SETTINGS.xlsx"), null, true),
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared StaffList = let
    Source = Table_AllocationExtracted,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Date", "End", "Unit", "Start"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates";

shared MealBreakTime = let
    Source = MealBreak,
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared MealBreakStart = let
    Source = MealBreak,
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MinShiftGap = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Resrodel\Customers - Documents\IndoChinese Aged Care\4. Analysis\221223\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap;

shared ShiftLength = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Resrodel\Customers - Documents\IndoChinese Aged Care\4. Analysis\221223\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftLent_Table = Source{[Item="ShiftLent",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftLent_Table,{{"MinHrs", Int64.Type}, {"ShortHrs", Int64.Type}, {"StdHrs", Int64.Type}, {"1.5OT", type number}, {"2.0OT", Int64.Type}})
in
    #"Changed Type";

shared StartNames = let
    Source = Table.NestedJoin(Table_Intervals, {"StartInterval"}, Table_StartTime_Duration, {"StartDateTime"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(Source, "StartTime+Duration", {"Name", "Role"}, {"Name", "Role"}),
    STARTLABEL = Table.AddColumn(#"Expanded StartTime+Duration", "TimeType", each "Start"),
    #"Renamed Columns" = Table.RenameColumns(STARTLABEL,{{"StartInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"EndPrecise", "EndInterval"})
in
    #"Removed Columns";

shared EndNames = let
    Source = Table.NestedJoin(Table_Intervals, {"EndPrecise"}, Table_StartTime_Duration, {"EndDateTime"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndInterval", type datetime}}),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(#"Changed Type", "StartTime+Duration", {"Name", "Role"}, {"Name", "Role"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded StartTime+Duration",{{"EndInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"StartInterval", "EndPrecise"}),
    #"Added TIMETYPE" = Table.AddColumn(#"Removed Columns", "TimeType", each "ShiftEnd")
in
    #"Added TIMETYPE";

shared #"NAMES+INTERVALS" = let
    Source = Table.Combine({StartNames, EndNames}),
    #"Removed Columns1" = Table.RemoveColumns(Source,{"ShiftPeriod"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns1",null,"x",Replacer.ReplaceValue,{"Name"}),
    SORTNAMESDATES = Table.Sort(#"Replaced Value",{{"Name", Order.Ascending}, {"Date", Order.Ascending}}),
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
    #"Merged INTERVALS" = Table.NestedJoin(#"Expanded AllocatedIntervals", {"AllocatedIntervals"}, Table_Intervals, {"IntervalID"}, "Intervals", JoinKind.LeftOuter),
    #"Expanded Intervals" = Table.ExpandTableColumn(#"Merged INTERVALS", "Intervals", {"StartInterval", "IntervalID", "EndInterval", "Duration", "ShiftPeriod"}, {"StartInterval", "IntervalID.1", "EndInterval", "Duration.1", "ShiftPeriod"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Expanded Intervals",{{"Duration.1", "IntervalDurationTemp"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns1",{{"StartInterval", "IntervalStart"}, {"EndInterval", "IntervalEnd"}, {"IntervalID", "ShiftStartInterval"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Added Index1.IntervalID", "Interval Range", "Date", "Duration", "ShiftStartInterval"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"IntervalStart", "IntervalEnd",  "Name", "AllocatedIntervals"}),
    #"Merged STAFFSHIFTS" = Table.NestedJoin(#"Reordered Columns", {"Name", "IntervalStart"}, StaffDoubleShifts, {"Name", "DateTime"}, "StaffDoubleShifts", JoinKind.LeftOuter),
    #"Expanded StaffDoubleShifts" = Table.ExpandTableColumn(#"Merged STAFFSHIFTS", "StaffDoubleShifts", {"Double Shift", "Effective Duration", "RealDuration", "ShiftType"}, {"Double Shift", "Effective Duration", "RealDuration", "ShiftType"}),
    #"Filled Down" = Table.FillDown(#"Expanded StaffDoubleShifts",{"Effective Duration", "ShiftType", "RealDuration"})
in
    #"Filled Down";

shared #"NAMES+INTERVALSLIST" = let
    Source = #"NAMES+INTERVALS",
    #"Unpivoted Only Selected Columns" = Table.Unpivot(Source, {"IntervalStart", "IntervalEnd"}, "Attribute", "Value"),
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
    #"Filtered Rows" = Table.SelectRows(#"Replaced NULL X", each ([Name] <> null)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"TimeDate", Order.Ascending}})
in
    #"Sorted Rows";

shared StaffShifts = let
    Source = Table_StartTime_Duration,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "Unit", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"StartDateTime", "StartOutter"}, {"EndDateTime", "EndOutter"}}),
    STARTOUTTER = Table.AddColumn(#"Renamed Columns", "StartInner", each [StartOutter] + #duration(0,0,1,0)),
    ENDINNER = Table.AddColumn(STARTOUTTER, "EndInner", each [EndOutter]-#duration(0,0,1,0)),
    #"Reordered Columns1" = Table.ReorderColumns(ENDINNER,{"Name", "Role", "Unit", "StartOutter", "StartInner", "EndInner", "EndOutter", "Duration", "Shift"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Reordered Columns1",{"Name", "Role", "Unit", "StartOutter", "StartInner", "EndOutter", "Duration", "Shift"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"StartInner", type datetime}, {"EndInner", type datetime}}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Name", "Role", "Unit", "Duration", "Shift"}, "Attribute", "Value"),
    #"Renamed Columns1" = Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "IntervalPoint"}, {"Value", "DateTime"}}),
    STAFFCOUNT = Table.AddColumn(#"Renamed Columns1", "StaffCount", each if [IntervalPoint] = "StartOuter" then 0 else if [IntervalPoint] = "StartInner" then 1 else if [IntervalPoint] = "EndInner" then 1 else 0),
    BUFFER = Table.Buffer(STAFFCOUNT)
in
    BUFFER;

shared ShiftIndex = let
    Source = StaffShifts,
    #"Removed Columns" = Table.RemoveColumns(Source,{"StaffCount", "Role", "Unit", "Duration", "Shift"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([IntervalPoint] = "StartOutter")),
    SHIFTINDEX = Table.AddIndexColumn(#"Filtered Rows", "ShiftIndex", 1, 1, Int64.Type)
in
    SHIFTINDEX;

shared MinHrs = let
    Source = ShiftLength,
    MinHrs = Source{0}[MinHrs]
in
    MinHrs;

shared ShortHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[ShortHrs]
in
    ShortHrs;

shared StdHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[StdHrs]
in
    ShortHrs;

shared #"15OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[1.5OT]
in
    ShortHrs;

shared #"20OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[2.0OT]
in
    ShortHrs;

shared ShiftPeriod = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Resrodel\Customers - Documents\IndoChinese Aged Care\4. Analysis\221223\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type";

shared ShiftGapIndices = let
    Source = Table_StartTime_Duration,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "Unit", "StartDateTime", "EndDateTime", "Duration", "Shift"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"StartDateTime", type number}}),
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
    #"Changed Type2";

shared ShiftGap = let
    Source = Table.NestedJoin(ShiftGapIndices, {"Name", "Index2.Index2"}, ShiftGapIndices, {"Name", "Index2.Custom.Index"}, "StaffShifts (2)", JoinKind.LeftOuter),
    #"Expanded StaffShifts (2)" = Table.ExpandTableColumn(Source, "StaffShifts (2)", {"Index2.Custom.StartDateTime", "Index2.Custom.Duration"}, {"StaffShifts (2).Index2.Custom.StartDateTime", "StaffShifts (2).Index2.Custom.Duration"}),
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
    BUFFER2;

shared EffectiveDuration = let
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
    BUFFER;

shared DoubleShifts = let
    Source = EffectiveDuration,
    #"Merged Queries1" = Table.NestedJoin(Source, {"Index"}, Source, {"Index.1"}, "Source", JoinKind.LeftOuter),
    #"Expanded Source" = Table.ExpandTableColumn(#"Merged Queries1", "Source", {"CombinedDuration", "DoubleShift1", "EffectiveDuration"}, {"Source.CombinedDuration", "Source.DoubleShift1", "Source.EffectiveDuration"}),
    #"DOUBLE SHIFT" = Table.AddColumn(#"Expanded Source", "Double Shift", each if [DoubleShift1] = "DoubleShift1" then "DoubleShift1" else if [Source.DoubleShift1] = "DoubleShift1" then "DoubleShift2" else null),
    REALDURATION = Table.AddColumn(#"DOUBLE SHIFT", "RealDuration", each if [Double Shift] = "DoubleShift1" then [CombinedDuration] else if [Double Shift] = "DoubleShift2" then [Source.CombinedDuration] else [CombinedDuration]),
    #"EFFECTIVE SHIFT" = Table.AddColumn(REALDURATION, "Effective Duration", each if [Double Shift] = "DoubleShift1" then [EffectiveDuration] else if [Double Shift] = "DoubleShift2" then [Source.EffectiveDuration] else [EffectiveDuration])
in
    #"EFFECTIVE SHIFT";

shared StaffDoubleShifts = let
    Source = Table.NestedJoin(StaffShifts, {"Name", "DateTime"}, ShiftIndex, {"Name", "DateTime"}, "EffectiveShift.1", JoinKind.LeftOuter),
    #"Expanded EffectiveShift.2" = Table.ExpandTableColumn(Source, "EffectiveShift.1", {"DateTime"}, {"EffectiveShift.1.DateTime"}),
    #"Filled Down" = Table.FillDown(#"Expanded EffectiveShift.2",{"EffectiveShift.1.DateTime"}),
    #"Merged Queries" = Table.NestedJoin(#"Filled Down", {"Name", "EffectiveShift.1.DateTime"}, DoubleShifts, {"Name", "Index2.Custom.StartDateTime"}, "DoubleShifts", JoinKind.LeftOuter),
    #"Expanded DoubleShifts" = Table.ExpandTableColumn(#"Merged Queries", "DoubleShifts", {"Double Shift", "Effective Duration", "RealDuration"}, {"Double Shift", "Effective Duration", "RealDuration"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded DoubleShifts",{"EffectiveShift.1.DateTime"}),
    SHIFTTYPE = Table.AddColumn(#"Removed Columns", "ShiftType", each if [RealDuration] <= 2/24 then "0-4Hrs" else if [RealDuration] <= 7.5/24 then "4-7.5Hrs" else if [RealDuration] <= 8.5/24 then "7.5-8.5Hrs" else if [RealDuration] <= 10/24 then "8.5-10Hrs" else if [RealDuration] <= 12/24 then "10-12Hrs" else ">12Hrs")
in
    SHIFTTYPE;

shared Shifts = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRVitWJVgqAUH6e7h4hSrGxAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Shifts = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}})
in
    #"Changed Type";

shared AMP = let
    Source = ShiftPeriod,
    StartDay = Source{0}[StartDay]
in
    StartDay;

shared PMP = let
    Source = ShiftPeriod,
    StartDay = Source{1}[StartDay]
in
    StartDay;

shared NIGHTP = let
    Source = ShiftPeriod,
    StartDay = Source{2}[StartDay]
in
    StartDay;

shared RolesList = let
    Source = Roles,
    ROLES = Source[ROLES]
in
    ROLES;

shared ShiftPoints = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCs7ITCsJLkksKlGK1YFyXfNSlGJjAQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ShiftPoints = _t]),
    ShiftPoints1 = Source[ShiftPoints]
in
    ShiftPoints1;

shared Roles = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 &  "\2. Calculations\SETTINGS.xlsx"), null, true),
    Roles_Table = Source{[Item="Roles",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Roles_Table,{{"ROLES", type text}})
in
    #"Changed Type";

shared Table_StartTime_Duration = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 & "\2. Calculations\Intervals.xlsx"), null, true),
    Table_StartTime_Duration_Table = Source{[Item="Table_StartTime_Duration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_StartTime_Duration_Table,{{"Date", type date}, {"Unit", type text}, {"Name", type text}, {"Role", type text}, {"Start", type number}, {"End", type number}, {"MEDS", type text}, {"StartDateTime", type datetime}, {"EndDateTime", type datetime}, {"Duration", type number}, {"Shift", type text}, {"EffectiveDuration", type number}})
in
    #"Changed Type";

shared Table_Intervals = let
        Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 & "\2. Calculations\Intervals.xlsx"), null, true),
    Table_Intervals_Table = Source{[Item="Table_Intervals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_Intervals_Table,{{"StartInterval", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}, {"EndPrecise", type datetime}, {"EndInterval", type datetime}, {"Duration", type number}, {"ShiftPeriod", type text}})
in
    #"Changed Type";

shared Table_IntervalsList = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 & "\2. Calculations\Intervals.xlsx"), null, true),
    Table_IntervalsList_Table = Source{[Item="Table_IntervalsList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_IntervalsList_Table,{{"IntervalID", Int64.Type}, {"DayDate", type date}, {"TimeDate", type datetime}, {"Time", type number}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Attribute", type text}, {"ShiftDate", type date}})
in
    #"Changed Type";

shared IntervalsList = let
    Source = Table_IntervalsList
in
    Source;