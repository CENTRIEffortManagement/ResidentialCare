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

shared FilePath = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT Shifts" = let

    Source = Excel.Workbook(File.Contents(FilePath  &  "\Shifts.xlsx"), null, true),
    BUFFER = Table.Buffer(Source)
in
    BUFFER;

shared #"EXTRACT NAMES_INTERVALSLIST" = let

    Source = #"IMPORT Shifts",
    NAMES_INTERVALSLIST_Table = Source{[Item="NAMES_INTERVALSLIST",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(NAMES_INTERVALSLIST_Table,{{"DayDate", type date}, {"Name", type text}, {"TimeType", type text}, {"AllocatedIntervals", Int64.Type}, {"IntervalID.1", Int64.Type}, {"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type datetime}, {"IntervalDurationTemp", type number}, {"Double Shift", type any}, {"RealDuration", type number}, {"Effective Duration", type number}, {"IntervalAssociatedShift", type text}, {"Attribute", type text}, {"ShiftType", type text}, {"ShiftDate", type date}, {"Role1", type text}, {"StaffCount", Int64.Type}})
in
    #"Changed Type";

shared #"EXTRACT ShiftsDayTime" = let
   
    Source = #"IMPORT Shifts",
    ShiftsDayTime_Table = Source{[Item="ShiftsDayTime",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftsDayTime_Table,{{"DateTime", type datetime}, {"Shift", type text}, {"TimeType", type text}})
in
    #"Changed Type";

shared #"IMPORT ShiftDuration" = let
     
    Source = Excel.Workbook(File.Contents(FilePath&"\Settings Data.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared #"IMPORT ShiftPeriod" = let

    
    Source = Excel.Workbook(File.Contents(FilePath&"\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type";

shared AMP = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{0}[StartDay]
in
    StartDay;

shared Roles = let
    Source = #"EXTRACT NAMES_INTERVALSLIST",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Role"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns")
in
    #"Removed Duplicates";

shared RoleShiftsDayTime = let
    Source = #"EXTRACT ShiftsDayTime",
    #"Removed Columns" = Table.RemoveColumns(Source,{"TimeType"}),
    #"Added ROLETABLE" = Table.AddColumn(#"Removed Columns", "Roles", each Roles),
    #"Expanded Roles" = Table.ExpandTableColumn(#"Added ROLETABLE", "Roles", {"Role"}, {"Roles.Role"})
in
    #"Expanded Roles";

shared ResourceShiftAllocation = let
    Source = ResourceIntervalAllocation,
    #"Grouped Rows" = Table.Group(Source, {"ShiftDate", "ShiftPeriod", "IntervalAssociatedShift", "Name", "Role", "TimeDate", "Attribute"}, {{"ResShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}, {"ResShiftEffectiveRatio", each List.Average([#"ResEffectiveRatio-FTE"]), type number}}),
    #"Added RESSHIFTFTE" = Table.AddColumn(#"Grouped Rows", "ResShiftFTE", each [ResShiftEffort]*24/#"IMPORT ShiftDuration"),
    #"Filtered Rows" = Table.SelectRows(#"Added RESSHIFTFTE", each ([Attribute] = "IntervalStart"))
in
    #"Filtered Rows";

shared #"ResourceShiftAllocation - less Intervals" = let
    Source = ResourceShiftAllocation,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"TimeDate", type date}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"ShiftPeriod", "Name", "Role", "TimeDate"}, {{"ResShiftFTE", each List.Sum([ResShiftFTE]), type number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared ResourceIntervalAllocation = let
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
    BUFFER;

shared RoleShiftAllocation = let
    Source = ResourceIntervalAllocation,
    #"Filtered Rows" = Table.SelectRows(Source, each ([StaffCount] = 1)),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"ShiftDate", "ShiftPeriod", "Role", "Attribute"}, {{"RoleShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}}),
    #"Filtered Rows1" = Table.SelectRows(#"Grouped Rows", each ([Attribute] = "IntervalStart")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"ShiftPeriod", "Role"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "ShiftPeriod.1", JoinKind.LeftOuter),
    #"Expanded ShiftPeriod.1" = Table.ExpandTableColumn(#"Merged Queries", "ShiftPeriod.1", {"DurationOfShifts"}, {"DurationOfShifts"}),
    ROLESHIFTFTE = Table.AddColumn(#"Expanded ShiftPeriod.1", "RoleShiftFTE", each [RoleShiftEffort]*24/[DurationOfShifts]),
    #"Sorted Rows2" = Table.Sort(ROLESHIFTFTE,{{"Role", Order.Ascending}, {"ShiftDate", Order.Ascending}})
in
    #"Sorted Rows2";

shared RoleShiftIntervalAllocation = let
    Source = ResourceIntervalAllocation,
    #"Grouped Rows" = Table.Group(Source, {"ShiftDate", "ShiftPeriod", "Role", "Attribute"}, {{"RoleShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Grouped Rows", {"ShiftPeriod"}, #"IMPORT ShiftPeriod", {"ShiftPeriod"}, "ShiftPeriod.1", JoinKind.LeftOuter),
    #"Expanded ShiftPeriod.1" = Table.ExpandTableColumn(#"Merged Queries", "ShiftPeriod.1", {"DurationOfShifts"}, {"DurationOfShifts"}),
    ROLESHIFTFTE = Table.AddColumn(#"Expanded ShiftPeriod.1", "RoleShiftFTE", each [RoleShiftEffort]*24/[DurationOfShifts]),
    #"Sorted Rows" = Table.Sort(ROLESHIFTFTE,{{"ShiftDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared RoleShiftEffortFTE = let
    Source = ResourceIntervalAllocation,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayDate", type date}})
in
    #"Changed Type";