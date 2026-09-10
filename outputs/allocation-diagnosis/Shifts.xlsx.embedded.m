section Section1;

shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
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
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(RawFilePath),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    SharePointCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SharepointRootUrl], type text),
    SharePointDocumentsCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type text),
    LocalCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SyncedFolderRootPath], type text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(
        MatchCandidates_WithLength,
        each [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    BestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(LocalFullPath, "[", "]") otherwise Text.AfterDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

shared #"FilePath - 2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"FilePath - 1Input" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;

shared #"IMPORT Table_AllocationExtracted" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Code", Int64.Type}, {"Date", type date}, {"Start", type datetime}, {"RoleShift", type text}, {"End", type datetime}, {"Break", Int64.Type}, {"Hours", type number}, {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Unit", type any}, {"Name", type text}})
in
    #"Changed Type1";

shared #"IMPORT Intervals" = let

    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Intervals.xlsx"), null, true)
in
    Source;

shared StaffList = let
    Source = #"IMPORT Table_AllocationExtracted",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Date", "End", "Unit", "Start"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates";

shared MealBreakTime = let
    Source = #"EXTRACT MealBreak",
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared MealBreakStart = let
    Source = #"EXTRACT MealBreak",
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MinShiftGap = let
    Source = #"IMPORT Settings Data",
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap;

shared ShiftLength = let
    Source = #"IMPORT Settings Data",
    ShiftLent_Table = Source{[Item="ShiftLent",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftLent_Table,{{"MinHrs", Int64.Type}, {"ShortHrs", Int64.Type}, {"StdHrs", Int64.Type}, {"1.5OT", type number}, {"2.0OT", Int64.Type}})
in
    #"Changed Type";

shared StartNames = let
    Source = Table.NestedJoin(Table_Intervals, {"StartInterval", "Role"}, #"EXTRACT Table_StartTime_Duration", {"StartDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(Source, "StartTime+Duration", {"Name"}, {"Name"}),
    STARTLABEL = Table.AddColumn(#"Expanded StartTime+Duration", "TimeType", each "Start"),
    #"Renamed Columns" = Table.RenameColumns(STARTLABEL,{{"StartInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"EndPrecise", "EndInterval"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Date", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows";

shared EndNames = let
    Source = Table.NestedJoin(Table_Intervals, {"EndPrecise", "Role"}, #"EXTRACT Table_StartTime_Duration", {"EndDateTime", "Role"}, "StartTime+Duration", JoinKind.LeftOuter),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EndInterval", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type",{{"EndInterval", Order.Ascending}}),
    #"Expanded StartTime+Duration" = Table.ExpandTableColumn(#"Sorted Rows1", "StartTime+Duration", {"Name"}, {"Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded StartTime+Duration",{{"EndInterval", "Date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"StartInterval", "EndPrecise"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "TimeType", each "ShiftEnd")
in
    #"Added Custom";

shared #"NAMES+INTERVALS" = let
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
    BUFFER2;

shared #"NAMES+INTERVALSLIST" = let
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
    #"Filtered NAMES ONLY";

shared StaffShifts = let
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
    BUFFER;

shared ShiftIndex = let
    Source = StaffShifts,
    #"Removed Columns" = Table.RemoveColumns(Source,{"StaffCount", "Role", "Duration", "Shift"}),
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
    Source = #"IMPORT Settings Data",
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type";

shared ShiftGapIndices = let
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
    #"Changed Type2";

shared ShiftGap = let
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
    #"Sorted Rows" = Table.Sort(#"Expanded EffectiveShift.2",{{"Name", Order.Ascending}, {"DateTime", Order.Ascending}}),
    #"Filled Down" = Table.FillDown(#"Sorted Rows",{"EffectiveShift.1.DateTime"}),
    #"Merged Queries" = Table.NestedJoin(#"Filled Down", {"Name", "EffectiveShift.1.DateTime"}, DoubleShifts, {"Name", "Index2.Custom.StartDateTime"}, "DoubleShifts", JoinKind.LeftOuter),
    #"Expanded DoubleShifts" = Table.ExpandTableColumn(#"Merged Queries", "DoubleShifts", {"Double Shift", "Effective Duration", "RealDuration"}, {"Double Shift", "Effective Duration", "RealDuration"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded DoubleShifts",{"EffectiveShift.1.DateTime"}),
    SHIFTTYPE = Table.AddColumn(#"Removed Columns", "ShiftType", each if [RealDuration] <= 2/24 then "0-4Hrs" else if [RealDuration] <= 7.5/24 then "4-7.5Hrs" else if [RealDuration] <= 8.5/24 then "7.5-8.5Hrs" else if [RealDuration] <= 10/24 then "8.5-10Hrs" else if [RealDuration] <= 12/24 then "10-12Hrs" else ">12Hrs")
in
    SHIFTTYPE;

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
    Source = #"EXTRACT Roles",
    ROLES = Source[ROLES]
in
    ROLES;

shared ShiftPoints = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCs7ITCsJLkksKlGK1YFyXfNSlGJjAQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ShiftPoints = _t]),
    ShiftPoints1 = Source[ShiftPoints]
in
    ShiftPoints1;

shared #"EXTRACT Table_StartTime_Duration" = let

    Source = #"IMPORT Intervals",
    Table_StartTime_Duration_Table = Source{[Item="Table_StartTime_Duration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_StartTime_Duration_Table,{{"Date", type date}, {"Name", type text}, {"Role", type text}, {"Start", type number}, {"End", type number}, {"StartDateTime", type datetime}, {"EndDateTime", type datetime}, {"Duration", type number}, {"Shift", type text}, {"EffectiveDuration", type number}})
in
    #"Changed Type";

shared #"EXTRACT Table Intervals" = let

    Source = #"IMPORT Intervals",
    Intervals_Table = Source{[Item="Intervals",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(Intervals_Table,{{"StartInterval", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}, {"EndInterval", type datetime}, {"EndPrecise", type datetime}, {"Duration", type number}, {"Role", type text}, {"ShiftPeriod", type text}, {"IntervalAssociatedShift", type text}})
in
    #"Changed Type1";

shared #"EXTRACT Table_IntervalsList" = let

    Source = #"IMPORT Intervals",
    Table_IntervalsList_Table = Source{[Item="IntervalsList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_IntervalsList_Table,{{"IntervalID", Int64.Type}, {"DayDate", type date}, {"TimeDate", type datetime}, {"Time", type number}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Attribute", type text}, {"ShiftDate", type date}})
in
    #"Changed Type";

shared #"IMPORT Settings Data" = let

    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true)
in
    Source;

shared #"EXTRACT MealBreak" = let

    Source = #"IMPORT Settings Data",
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared #"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared #"EXTRACT Roles" = let

    Source = #"IMPORT Settings Data",
    Roles_Table = Source{[Item="Roles",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Roles_Table,{{"ROLES", type text}})
in
    #"Changed Type";

shared #"MANUAL Shifts" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Shifts = _t, Order = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}})
in
    #"Changed Type";

shared Table_Intervals = let
        Source1 = #"EXTRACT Table Intervals",
    #"Removed Columns" = Table.RemoveColumns(Source1,{"IntervalAssociatedShift"})
in
    #"Removed Columns";

shared #"Table Intervals-IntervalAssocShift" = let
    Source = #"EXTRACT Table Intervals"
in
    Source;

shared IntervalsList = let
    Source = #"EXTRACT Table_IntervalsList"
in
    Source;

[ Description = "Trimmed to 4.5 hoours min." ]
shared #"NAMES+INTERVALSMATRIX" = let
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
    #"Pivoted Column1";