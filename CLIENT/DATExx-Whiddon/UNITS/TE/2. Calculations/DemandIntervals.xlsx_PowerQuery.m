// Power Query from: DemandIntervals.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\DemandIntervals.xlsx
// Extracted: 2026-09-24T08:43:15.027Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook's calculation folder through canonical FilePathUrl and CentriSyncPaths.
shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        PathInput = if Table.HasColumns(Source, {"FilePath"}) then Table.SelectColumns(Source, {"FilePath"})
            else if Table.ColumnNames(Source) = {"Column1"} then Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must contain a FilePath column or be a single named cell.",
        ValidatedRows = if Table.RowCount(PathInput) = 1 then PathInput
            else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedRows)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if not Value.Is(RawFilePathValue, type text) then
        error "FilePathUrl[FilePath] must contain the saved workbook path as text."
        else if Text.Trim(RawFilePathValue) = "" then
        error "FilePathUrl[FilePath] is blank. Save this workbook and recalculate its CELL filename formula."
        else Text.Trim(RawFilePathValue),
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

shared FilePath = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Filtered Rows", {{"Value", each Text.BeforeDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
    Value = #"Extracted Text Before Delimiter"{0}[Value]
in
    Value;

shared #"IMPORT Table_ShiftDemandHRS" = let

    
    Source = Excel.Workbook(File.Contents(FilePath &  "\1. Input\2-DemandExtract.xlsx"), null, true),
    ShiftUnitDemandHRS_Sheet = Source{[Item="ShiftUnitDemandHRS",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ShiftUnitDemandHRS_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Day", Int64.Type}, {"Shift", type text}, {"Period", Int64.Type}, {"Role", type text}, {"StartTime", type datetime}, {"EndTime", type datetime}, {"Unit", type any}, {"Facility", type any}, {"DemandFTE", type number}, {"DemandHRS", type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"DurationOfShifts", "ShiftDuration"}})
in
    #"Renamed Columns";

shared Table_IntervalIDList = let

    Source = Excel.Workbook(File.Contents(FilePath & "\2. Calculations\Intervals.xlsx"), null, true),
    RoleIntervalID_Table = Source{[Item="RoleIntervalID",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RoleIntervalID_Table,{{"DateTime", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}})
in
    #"Changed Type";

shared Table_Intervals = let
    
    
    
    Source = Excel.Workbook(File.Contents(FilePath &  "\2. Calculations\Intervals.xlsx"), null, true),
    Table_Intervals_Table = Source{[Item="Intervals",Kind="Table"]}[Data],
    #"Sorted Rows" = Table.Sort(Table_Intervals_Table,{{"StartInterval", Order.Ascending}})
in
    #"Sorted Rows";

shared IntervalIDs = let
    Source = Table_IntervalIDList
in
    Source;

shared ShiftINTERVALS = let
    Source = #"IMPORT Table_ShiftDemandHRS",
    #"Grouped Rows" = Table.Group(Source, {"Shift", "Role", "Facility", "StartTime", "EndTime", "Date"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Role", Order.Ascending}, {"StartTime", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Count"}),
    #"Merged DEMAND+STARTINTERVALID" = Table.NestedJoin(#"Removed Columns1", {"StartTime"}, IntervalIDs, {"DateTime"}, "IntervalIDs", JoinKind.LeftOuter),
    #"Expanded IntervalIDs" = Table.ExpandTableColumn(#"Merged DEMAND+STARTINTERVALID", "IntervalIDs", {"DateTime", "DayDate", "IntervalID"}, {"DateTime", "DayDate", "IntervalID"}),
    #"Sorted Rows3" = Table.Sort(#"Expanded IntervalIDs",{{"IntervalID", Order.Ascending}}),
    #"Merged DEMAND+ENDINTERVAL" = Table.NestedJoin(#"Sorted Rows3", {"EndTime"}, IntervalIDs, {"DateTime"}, "IntervalIDs", JoinKind.LeftOuter),
    #"Expanded IntervalIDs1" = Table.ExpandTableColumn(#"Merged DEMAND+ENDINTERVAL", "IntervalIDs", {"DateTime", "IntervalID"}, {"DateTime.1", "IntervalID.1"}),
    #"Sorted Rows4" = Table.Sort(#"Expanded IntervalIDs1",{{"DateTime.1", Order.Ascending}}),
    #"Inserted RANGE" = Table.AddColumn(#"Sorted Rows4", "Range", each if [IntervalID.1] - [IntervalID] <> 0
then [IntervalID.1] - [IntervalID] -1
else 0),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted RANGE",{{"Range", Int64.Type}}),
    #"Sorted Rows2" = Table.Sort(#"Changed Type",{{"Role", Order.Ascending}, {"StartTime", Order.Ascending}}),
    #"Added INTERVALS" = Table.AddColumn(#"Sorted Rows2", "IntervalListx", each if [IntervalID] = null
then null 
else 
List.Numbers([IntervalID],[Range]+1)),
    #"Expanded IntervalListx" = Table.ExpandListColumn(#"Added INTERVALS", "IntervalListx"),
    #"Removed Duplicates" = Table.Distinct(#"Expanded IntervalListx"),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Duplicates",{"IntervalID", "IntervalID.1", "Range"}),
    #"Sorted Rows1" = Table.Sort(#"Removed Columns",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}})
in
    #"Sorted Rows1";

[ Description = "Modified for ANACC demand#(lf)Not master roster demand" ]
shared ShiftDemandUnitINTERVAL = let
    Source = ShiftINTERVALS,
    #"Merged INTERVALS" = Table.NestedJoin(Source, {"IntervalListx", "Role"}, Table_Intervals, {"IntervalID", "Role"}, "Table_Intervals", JoinKind.LeftOuter),
    #"Expanded Table_Intervals" = Table.ExpandTableColumn(#"Merged INTERVALS", "Table_Intervals", {"StartInterval", "EndInterval", "Duration", "ShiftPeriod"}, {"StartInterval", "EndInterval", "Duration", "ShiftPeriod"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded Table_Intervals",{{"StartInterval", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type1",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Sorted Rows1",{{"EndInterval", type datetime}}),
    #"Merged SHIFTUNITDEMAND" = Table.NestedJoin(#"Changed Type2", {"Facility", "Role", "StartTime", "EndTime"}, #"IMPORT Table_ShiftDemandHRS", {"Facility", "Role", "StartTime", "EndTime"}, "Table_ShiftDemandHRS", JoinKind.LeftOuter),
    #"Expanded Table_ShiftDemandHRS" = Table.ExpandTableColumn(#"Merged SHIFTUNITDEMAND", "Table_ShiftDemandHRS", {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}, {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}),
    EFFECTIVEDURATION = Table.AddColumn(#"Expanded Table_ShiftDemandHRS", "EffectiveDuration", each if [ShiftDuration] < MealBreakStart
then [ShiftDuration]
else [ShiftDuration] - 0.5),
    INTEFFECTIVERATIO = Table.AddColumn(EFFECTIVEDURATION, "IntervalEffectiveRatio", each [EffectiveDuration]/[ShiftDuration]),
    //not applicable for ANACC demand - only master roster 
    //#"Added EFFECTIVEDEMAND !!" = Table.AddColumn(INTEFFECTIVERATIO, "EffectiveIntervalAttendance", each [DemandFTE]*[IntervalEffectiveRatio]),
    #"Added EFFECTIVEDEMAND !!" = Table.AddColumn(INTEFFECTIVERATIO, "EffectiveIntervalAttendance", each [DemandFTE]),
    //*[IntervalEffectiveRatio])"
    #"Added DEMAND EFFORT" = Table.AddColumn(#"Added EFFECTIVEDEMAND !!", "DemandEffort", each [Duration] * [EffectiveIntervalAttendance]*24)
in
    #"Added DEMAND EFFORT";

shared ShiftDemandINTERVALLIST = let
    Source = ShiftDemandUnitINTERVAL,
    #"Unpivoted Only Selected Columns" = Table.Unpivot(Source, {"StartInterval", "EndInterval"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Only Selected Columns",{{"Attribute", "PointType"}, {"Value", "TimeDate"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"TimeDate", type datetime}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each true),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}, {"EndTime", Order.Ascending},  {"PointType", Order.Descending}})
in
    #"Sorted Rows";

shared Meals = let
    Source = Excel.Workbook(File.Contents(FilePath & "\2. Calculations\Settings Data.xlsx"), null, true),
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared MealBreakStart = let
    Source = Meals,
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MealBreakTime = let
    Source = Meals,
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared Table_ShiftDemandHRSCheck = let
    Source = #"IMPORT Table_ShiftDemandHRS",
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role"}, {{"DemandHrsRoster", each List.Sum([DemandHRS]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DemandHrsWeek", each [DemandHrsRoster]/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "CareDemandHrsPer WeekApprox", each [DemandHrsWeek] *.936)
in
    #"Added Custom1";

shared ShiftDemandUnitINTERVALCheck = let
    Source = ShiftDemandUnitINTERVAL,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Facility"}, {{"DemandHrsPerRoster", each List.Sum([DemandEffort]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DemandHrsPerWeek", each [DemandHrsPerRoster]/2)
in
    #"Added Custom";
