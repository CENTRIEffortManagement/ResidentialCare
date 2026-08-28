// Power Query from: Intervals.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Intervals.xlsx
// Extracted: 2026-05-18T06:14:23.971Z

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

shared #"FilePath-2Calculations" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    Folder = #"Renamed Columns"{0}[Folder]
in
    Folder;

shared ShiftStart = let


    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations"&"\Settings Data.xlsx"), null, true),
    ShiftStart_Table = Source{[Item="ShiftStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftStart_Table,{{"StartAM", type number}, {"StartPM", type number}, {"StartNIGHT", type number}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"StartAM", each Number.Round(_, 3), type number}, {"StartPM", each Number.Round(_, 3), type number}, {"StartNIGHT", each Number.Round(_, 3), type number}})
in
    #"Rounded Off";

shared #"IMPORT Roles" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations" & "\Settings Data.xlsx"), null, true),
    Roles_Table = Source{[Item="Roles",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Roles_Table,{{"ROLES", type text}})
in
    #"Changed Type";

shared #"IMPORT ShiftPeriod" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations" & "\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data]
in
    ShiftPeriod_Table;

shared DateFrom = let

    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations"&"\Settings Data.xlsx"), null, true),

    DateFrom_Table = Source{[Item="DateFrom",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(DateFrom_Table,{{"DateFrom", type date}}),
    DateFrom1 = #"Changed Type"{0}[DateFrom]
in
    DateFrom1;

shared RolesList = let
    Source = #"IMPORT Roles",
    ROLES = Source[ROLES]
in
    ROLES;

shared ShiftPoints = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCs7ITCsJLkksKlGK1YFyXfNSlGJjAQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ShiftPoints = _t]),
    ShiftPoints1 = Source[ShiftPoints]
in
    ShiftPoints1;

shared #"R1-AMP" = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{0}[StartDay]
in
    StartDay;

shared #"R1-PMP" = let
    Source = #"IMPORT ShiftPeriod",
    #"Rounded Off" = Table.TransformColumns(Source,{{"StartDay", each Number.Round(_, 4), type number}}),
    StartDay = #"Rounded Off"{1}[StartDay]
in
    StartDay;

shared #"R1-NIGHTP" = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{2}[StartDay]
in
    StartDay;

shared Start_AM = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{0}[StartDay]
in
    StartDay;

shared StartPM = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{1}[StartDay]
in
    StartDay;

shared StartNIGHT = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{2}[StartDay]
in
    StartDay;

shared #"ShiftPeriod END" = let
    Source = #"IMPORT ShiftPeriod",
    #"Removed Columns" = Table.RemoveColumns(Source,{"StartTime", "DurationOfShifts"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"Role", "ShiftPeriod"}, "PointType", "TimeInDay"),
    #"Filtered Rows" = Table.SelectRows(#"Unpivoted Columns", each ([PointType] = "EndDay")),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows",{"PointType"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns1", List.Distinct(#"Removed Columns1"[ShiftPeriod]), "ShiftPeriod", "TimeInDay")
in
    #"Pivoted Column";

shared MealBreakStart = let
    Source = MealBreak,
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MealBreakTime = let
    Source = MealBreak,
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared DemandIntervals = let
    Source = #"IMPORT ShiftUnitDemandHRS",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Shift", "Role", "StartTime", "EndTime"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Sorted Rows1" = Table.Sort(#"Removed Duplicates",{{"Role", Order.Ascending}}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Sorted Rows1", {"Shift","Role"
}, "Attribute", "ShiftPointTime"),
    #"Sorted Rows" = Table.Sort(#"Unpivoted Columns",{{"ShiftPointTime", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Attribute"}),
    #"Removed Duplicates1" = Table.Distinct(#"Removed Columns1"),
    #"Added SHIFTPOINT" = Table.AddColumn(#"Removed Duplicates1", "IntervalType", each "ShiftPoint"),
    #"Renamed Columns" = Table.RenameColumns(#"Added SHIFTPOINT",{{"IntervalType", "TimeType"}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Renamed Columns",{{"ShiftPointTime", "DateTime"}}),
    #"Inserted Date" = Table.AddColumn(#"Renamed Columns1", "Date", each DateTime.Date([DateTime]), type date),
    #"Renamed Columns2" = Table.RenameColumns(#"Inserted Date",{{"Date", "DayDate"}}),
    BUFFER = Table.Buffer(#"Renamed Columns2"),
    #"Sorted Rows3" = Table.Sort(BUFFER,{{"DateTime", Order.Ascending}}),
    #"Sorted Rows2" = Table.Sort(#"Sorted Rows3",{{"Role", Order.Ascending}, {"DateTime", Order.Ascending}})
in
    #"Sorted Rows2";

shared #"StartTime+Duration" = let
    Source = Table_AllocationExtracted,
    FILTERROLES = Table.SelectRows(Source, each true),
    #"Changed Type1" = Table.TransformColumnTypes(FILTERROLES,{{"Start", type time}, {"End", type time}, {"Date", type date}}),
    #"Inserted Merged Date and Time" = Table.AddColumn(#"Changed Type1", "StartDateTime", each [Date] & [Start], type datetime),
    #"Inserted Merged Date and Time1" = Table.AddColumn(#"Inserted Merged Date and Time", "EndDateTime", each if [End]<[Start]
then (Date.AddDays([Date],1) & [End])
else

[Date] & [End]),
    #"Changed Type2" = Table.TransformColumnTypes(#"Inserted Merged Date and Time1",{{"EndDateTime", type datetime}}),
    DURATION = Table.AddColumn(#"Changed Type2", "Duration", each [EndDateTime] - [StartDateTime]
/* (if [End]<[Start]
then (1-[Start])+[End]
else [End]-[Start])*24 */),
    #"Changed Type" = Table.TransformColumnTypes(DURATION,{{"EndDateTime", type datetime}, {"Start", type number}, {"End", type number}, {"Duration", type number}}),
    SHIFT = Table.AddColumn(#"Changed Type", "Shift", each if [Start] >= Start_AM 
and 
[Start] < StartPM 
then "AM"

else if [Start]>= StartPM 
and 
[Start]< StartNIGHT 
then "PM"

else "NIGHT"),
    WORKTIME = Table.AddColumn(SHIFT, "EffectiveDuration", each if [Duration] > MealBreakStart 
then [Duration] - MealBreakTime
else [Duration])
in
    WORKTIME;

shared STARTS = let
    Source = #"StartTime+Duration",
    #"STARTS+ENDSCOLS" = Table.SelectColumns(Source,{"StartDateTime", "Shift"}),
    #"Removed Duplicates" = Table.Distinct(#"STARTS+ENDSCOLS"),
    STARTTIMETYPE = Table.AddColumn(#"Removed Duplicates", "TimeType", each "Start"),
    #"Renamed Columns" = Table.RenameColumns(STARTTIMETYPE,{{"StartDateTime", "DateTime"}}),
    #"Inserted Date" = Table.AddColumn(#"Renamed Columns", "DayDate", each DateTime.Date([DateTime]), type date),
    #"Sorted Rows" = Table.Sort(#"Inserted Date",{{"DateTime", Order.Ascending}})
in
    #"Sorted Rows";

shared ENDS = let
    Source = #"StartTime+Duration",
    ENDSCOLS = Table.SelectColumns(Source,{"EndDateTime", "Shift"}),
    #"Removed Duplicates" = Table.Distinct(ENDSCOLS),
    ENDS = #"Removed Duplicates",
    ENDTIMETYPE = Table.AddColumn(ENDS, "TimeType", each "End"),
    #"Renamed Columns" = Table.RenameColumns(ENDTIMETYPE,{{"EndDateTime", "DateTime"}}),
    #"Inserted Date" = Table.AddColumn(#"Renamed Columns", "DayDate", each DateTime.Date([DateTime]), type date),
    #"Sorted Rows" = Table.Sort(#"Inserted Date",{{"DateTime", Order.Ascending}})
in
    #"Sorted Rows";

[ Description = "Rostered Name EXCLUDED - reconnect later" ]
shared #"RoleIntervalID !!" = let
    Source = Table.Combine({DemandIntervals,STARTS,ENDS}),
    #"Sorted Rows" = Table.Sort(Source,{{"Role", Order.Ascending}, {"DateTime", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Shift", "Role", "TimeType"}),
    #"Sorted Rows1" = Table.Sort(#"Removed Columns1",{{"DateTime", Order.Ascending}}),
    #"Removed Duplicates1" = Table.Distinct(#"Sorted Rows1", {"DateTime"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Duplicates1", "IntervalID", 1, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

[ Description = "Rostered Name EXCLUDED - reconnect later" ]
shared Intervals = let
    Source = #"RoleIntervalID !!",
    BUFFER1 = Table.Buffer(Source),
    #"Added Index" = Table.AddIndexColumn(BUFFER1, "Index", 0, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 1, 1, Int64.Type),
    #"Merged CURRENTINDEX" = Table.NestedJoin(#"Added Index1", {"Index.1"}, #"Added Index1", {"Index"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded Added Index1" = Table.ExpandTableColumn(#"Merged CURRENTINDEX", "Added Index1", {"DateTime"}, {"Added Index1.DateTime"}),
    #"Added Remove Tail" = Table.AddColumn(#"Expanded Added Index1", "Custom", each if [Added Index1.DateTime] = null then "Remove Tail" else if [Added Index1.DateTime] < [DateTime] then "Remove tail" else null),
    #"Filtered REMOVETAIL" = Table.SelectRows(#"Added Remove Tail", each [Custom] = null),
    #"Removed Columns2" = Table.RemoveColumns(#"Filtered REMOVETAIL",{"Custom"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns2",{{"Added Index1.DateTime", "EndPrecise"}, {"DateTime", "StartInterval"}}),
    FILTERNULLEND = Table.SelectRows(#"Renamed Columns", each ([EndPrecise] <> null)),
    #"Removed Columns" = Table.RemoveColumns(FILTERNULLEND,{"Index", "Index.1"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Removed Columns",{{"StartInterval", type datetime}, {"EndPrecise", type datetime}}),
    #"END-1" = Table.AddColumn(#"Changed Type3", "EndInterval", each [EndPrecise]-#duration(0,0,1,0)),
    #"Changed Type1" = Table.TransformColumnTypes(#"END-1",{{"EndInterval", type datetime}, {"StartInterval", type datetime}, {"DayDate", type date}, {"EndPrecise", type datetime}}),
    DURATIONTEMP = Table.AddColumn(#"Changed Type1", "Durationtemp", each [EndPrecise]-[StartInterval]),
    #"Changed Type" = Table.TransformColumnTypes(DURATIONTEMP,{{"Durationtemp", type number}}),
    DURATION = Table.AddColumn(#"Changed Type", "Duration", each [Durationtemp]),

    #"Changed Type2" = Table.TransformColumnTypes(DURATION,{{"Duration", type number}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type2",{"Durationtemp"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns1",{ {"IntervalID", Order.Ascending}}),
    #"Inserted START" = Table.AddColumn(#"Sorted Rows", "StartTime", each DateTime.Time([StartInterval]), type time),
    #"Inserted END" = Table.AddColumn(#"Inserted START", "EndTime", each DateTime.Time([EndInterval]), type time),
    #"Changed Type6" = Table.TransformColumnTypes(#"Inserted END",{{"StartTime", type number}, {"EndTime", type number}}),
    BUFFER = Table.Buffer(#"Changed Type6"),
    #"Added ROLES" = Table.AddColumn(BUFFER, "Roles", each #"IMPORT Roles"),
    #"Expanded Roles" = Table.ExpandTableColumn(#"Added ROLES", "Roles", {"ROLES"}, {"ROLES"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Expanded Roles",{{"ROLES", "Role"}}),
    #"Merged Queries1" = Table.NestedJoin(#"Renamed Columns1", {"Role"}, #"ShiftPeriod END", {"Role"}, "IMPORT ShiftPeriod", JoinKind.LeftOuter),
    #"Expanded IMPORT ShiftPeriod1" = Table.ExpandTableColumn(#"Merged Queries1", "IMPORT ShiftPeriod", {"AM", "PM", "NIGHT"}, {"AM", "PM", "NIGHT"}),
    #"Rounded Off" = Table.TransformColumns(#"Expanded IMPORT ShiftPeriod1",{{"StartTime", each Number.Round(_, 6), type number}, {"EndTime", each Number.Round(_, 6), type number}, {"AM", each Number.Round(_, 6), type number}, {"PM", each Number.Round(_, 6), type number}, {"NIGHT", each Number.Round(_, 6), type number}}),
    #"Added SHIFTPERIOD" = Table.AddColumn(#"Rounded Off", "ShiftPeriod", each if [StartTime] >= [NIGHT] and [EndTime] < [AM] and [StartTime] <[EndTime]
then "AM"
else 
if [StartTime] >= [AM] and [EndTime] < [PM] and [StartTime] <[EndTime]
then "PM"
else 
if ([StartTime] >= [PM] or [StartTime] >0)

then "NIGHT"
else 
"ERROR"),
    #"Sorted Rows1" = Table.Sort(#"Added SHIFTPERIOD",{{"ShiftPeriod", Order.Ascending}}),
    #"Added INTERVALASSOCIATEDSHIFT" = Table.AddColumn(#"Sorted Rows1", "IntervalAssociatedShift", each      [ShiftPeriod]),

    #"Removed Columns3" = Table.RemoveColumns(#"Added INTERVALASSOCIATEDSHIFT",{"StartTime", "EndTime"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns3",{"StartInterval", "DayDate", "IntervalID", "EndInterval", "EndPrecise", "Duration", "ShiftPeriod"}),
    #"Removed Columns4" = Table.RemoveColumns(#"Reordered Columns",{"AM", "PM", "NIGHT"})
in
    #"Removed Columns4"


/* if [Durationtemp] > 0.34       

then  
DateTime.Time([EndPrecise]) - DateTime.Time([StartInterval]) else [Durationtemp] */


           /* if 
            [StartTime] > (AMP-0.0625)
            and 
            ([StartTime] < PMP)
        then "AM"

        else if 
            [StartTime] >= PMP
            and 
            [StartTime] < NIGHTP
        then "PM"

        else "NIGHT")
        */;

shared IntervalsList = let
    Source = Intervals,
    #"Reordered Columns" = Table.ReorderColumns(Source,{"StartInterval",  "DayDate", "EndInterval", "EndPrecise", "Duration", "IntervalID"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"EndPrecise"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"DayDate", "Duration", "IntervalID","ShiftPeriod","IntervalAssociatedShift","Role"
 }, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Columns",{{"Value", "TimeDate"}}),
    TIME = Table.AddColumn(#"Renamed Columns", "Time", each DateTime.Time([TimeDate]), type time),
    #"Changed Type" = Table.TransformColumnTypes(TIME,{{"Time", type number}}),
    SHIFTDATE = Table.AddColumn(#"Changed Type", "ShiftDate", each if [Time] > 0 and [Time] < #"R1-AMP"
then Date.AddDays(DateTime.Date([TimeDate]), -1)
else  DateTime.Date([TimeDate])),
    #"Changed Type1" = Table.TransformColumnTypes(SHIFTDATE,{{"ShiftDate", type date}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"IntervalID", Order.Ascending}, {"TimeDate", Order.Ascending}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Sorted Rows",{"IntervalID", "DayDate", "TimeDate", "Time", "Duration", "ShiftPeriod", "Attribute", "ShiftDate"})
in
    #"Reordered Columns1";

shared #"IMPORT ShiftUnitDemandHRS" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-1Input" & "\2-DemandExtract.xlsx"), null, true),
    #"Sorted Rows" = Table.Sort(Source,{{"Name", Order.Ascending}}),
    ShiftUnitDemandHRS_Sheet = #"Sorted Rows"{[Item="ShiftUnitDemandHRS",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ShiftUnitDemandHRS_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Day", Int64.Type}, {"Shift", type text}, {"Period", Int64.Type}, {"Role", type text}, {"StartTime", type datetime}, {"EndTime", type datetime}, {"Unit", type any}, {"Facility", type any}, {"DemandFTE", type number}, {"DemandHRS", type number}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type",{{"StartTime", Order.Ascending}})
in
    #"Sorted Rows1";

shared MealBreak = let
    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations"&"\Settings Data.xlsx"), null, true),
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared Table_AllocationExtracted = let
Source = Excel.Workbook(File.Contents(#"FilePath-1Input"& "\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Code", Int64.Type}, {"Date", type date}, {"Start", type datetime}, {"End", type datetime}, {"Break", Int64.Type}, {"Hours", type number}, {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Unit", type any}, {"Name", type text}})
in
    #"Changed Type";

shared #"FilePath-1Input" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","2. Calculations","1. Input",Replacer.ReplaceText,{"Folder"}),
    Folder = #"Replaced Value"{0}[Folder]
in
    Folder;
