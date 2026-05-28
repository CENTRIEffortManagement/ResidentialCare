// Power Query from: Demand.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Demand.xlsx
// Extracted: 2026-05-18T06:14:16.287Z

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

shared #"Folder-2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"Folder-1Input" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;

shared #"INPUT DateShift" = let

    DateShift1 = 0
    
in
    DateShift1;

shared #"INPUT RosterBudget-OFF" = let
        Source1 = 0
    /*Folder = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder &  "\1. Input\Budget and AN-ACC Comparison.xlsx"), null, true),
    RosterBudget_Table = Source{[Item="RosterBudget",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RosterBudget_Table,{{"Location", type text}, {"Shift", type text}, {"Role", type text}, {"Notes", type text}, {"Column1", type any}, {"Mon", type number}, {"Tue", type number}, {"Wed", type number}, {"Thu", type number}, {"Fri", type number}, {"Sat", type number}, {"Sun", type number}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"Mon", "D1-Mon"}, {"Tue", "D2-Tue"}, {"Wed", "D3-Wed"}, {"Thu", "D4-Thu"}, {"Fri", "D5-Fri"}, {"Sat", "D6-Sat"}, {"Sun", "D7-Sun"}}),
    #"Removed Top Rows" = Table.Skip(#"Renamed Columns1",20),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Top Rows",{"Column1", "Notes"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([Location] <> null)),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each not Text.Contains([Location], "TOTAL")),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows1", each ([Shift] <> null) and ([Role] = "CSE2" or [Role] = "CSE4" or [Role] = "RN")),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Filtered Rows2", {"Location", "Shift", "Role"}, "Attribute", "Value"),
    #"Sorted Rows1" = Table.Sort(#"Unpivoted Columns",{{"Attribute", Order.Ascending}, {"Location", Order.Ascending}, {"Shift", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows1",{{"Attribute", "Day"}, {"Value", "Hrs"}}),
    FTE = Table.AddColumn(#"Renamed Columns", "FTE", each [Hrs]/ShiftDuration),
    #"Removed Columns1" = Table.RemoveColumns(FTE,{"Hrs"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns1",{{"Day", Order.Ascending}})
in
    #"Sorted Rows" */
    in 
    Source1;

shared #"IMPORT ShiftDuration" = let


    Source = Excel.Workbook(File.Contents(#"Folder-2Calculations" &  "\Settings Data.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ShiftDuration_Table,{{"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type1"{0}[ShiftDuration]
in
    ShiftDuration1;

shared #"IMPORT ShiftHours" = let
    Source = Excel.Workbook(File.Contents(#"Folder-2Calculations"&"\Settings Data.xlsx"), null, true),
    ShiftHours_Sheet = Source{[Item="ShiftHours",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ShiftHours_Sheet, [PromoteAllScalars=true])
in
    #"Promoted Headers";

shared #"IMPORT ShiftDemandUnitINTERVAL" = let
    Source = Excel.Workbook(File.Contents(#"Folder-2Calculations"&"\DemandIntervals.xlsx"), null, true),
    Table_ShiftDemandUnitINTERVAL_Table = Source{[Item="Table_ShiftDemandUnitINTERVAL",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_ShiftDemandUnitINTERVAL_Table,{{"Shift", type text}, {"Role", type text}, {"Facility", type text}, {"StartTime", type datetime}, {"EndTime", type datetime}, {"Date", type date}, {"IntervalListx", Int64.Type}, {"StartInterval", type datetime}, {"EndInterval", type datetime}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Unit", type text}, {"ShiftDuration", type number}, {"DemandFTE", Int64.Type}, {"DemandHRS", type number}, {"EffectiveDuration", type number}, {"IntervalEffectiveRatio", type number}, {"EffectiveIntervalAttendance", type number}})
in
    #"Changed Type";

shared #"IMPORT IntervalsList" = let
    Source = Excel.Workbook(File.Contents(#"Folder-2Calculations"&"\Intervals.xlsx"), null, true),
    IntervalsList_Sheet = Source{[Item="IntervalsList",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(IntervalsList_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"IntervalID", Int64.Type}, {"DayDate", type date}, {"TimeDate", type datetime}, {"Time", type number}, {"Duration", type number}, {"ShiftPeriod", type text}, {"Attribute", type text}, {"ShiftDate", type date}})
in
    #"Changed Type";

shared #"IMPORT Table_RosterStart" = let
    Source = Excel.Workbook(File.Contents(#"Folder-1Input" &"\1-AllocationExtracted.xlsx"), null, true),
    Table_RosterStart_Table = Source{[Item="Table_RosterStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_RosterStart_Table,{{"RosterStart", type date}})
in
    #"Changed Type";

shared #"ShiftDurations!!" = let
    Source = #"IMPORT ShiftDemandUnitINTERVAL",
    #"Removed Duplicates1" = Table.Distinct(Source, {"Role", "IntervalListx"}),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates1",{{"Role", Order.Ascending}, {"DateTime", Order.Ascending}}),
    #"Grouped Rows1" = Table.Group(#"Sorted Rows", {"Role", "Shift", "Date"}, {{"Duration", each List.Sum([Duration]), type nullable number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows1",{{"Date", "DateX"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX],#"INPUT DateShift")),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"DateX"})
in
    #"Removed Columns";

shared #"ShiftEffort !!" = let
    Source = #"IMPORT ShiftDemandUnitINTERVAL",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Shift", "Role", "Facility", "Date", "IntervalListx", "Duration", "ShiftPeriod", "Unit", "DemandFTE", "IntervalEffectiveRatio", "EffectiveIntervalAttendance"}),
    #"Inserted Multiplication" = Table.AddColumn(#"Removed Other Columns", "UnitIntervalEffort", each [EffectiveIntervalAttendance] * [Duration], type number),
    #"Grouped Rows" = Table.Group(#"Inserted Multiplication", {"Facility", "Role", "Date", "ShiftPeriod"}, {{"UnitShiftEffort", each List.Sum([UnitIntervalEffort]), type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Date", "DateX"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX],#"INPUT DateShift")),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"DateX"})
in
    #"Removed Columns";

shared ShiftDemandHCAverageM = let
    Source = Table.NestedJoin(#"ShiftEffort !!", {"Role", "Date", "ShiftPeriod"}, #"ShiftDurations!!", {"Role", "Date", "Shift"}, "ShiftDurations", JoinKind.LeftOuter),
    #"Expanded ShiftDurations" = Table.ExpandTableColumn(Source, "ShiftDurations", {"Duration"}, {"ShiftDurations.Duration"}),
    #"Inserted SHIFTDEMANDHCAV." = Table.AddColumn(#"Expanded ShiftDurations", "ShiftDemandHCAverage", each [UnitShiftEffort] / [ShiftDurations.Duration], type number),
    #"Sorted Rows" = Table.Sort(#"Inserted SHIFTDEMANDHCAV.",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"ShiftPeriod", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"Date", type date}})
in
    #"Changed Type";

shared ShiftIntervalDemandHCAveM = let
    Source = Table.NestedJoin(#"IMPORT IntervalsList", {"DayDate", "ShiftPeriod"}, ShiftDemandHCAverageM, {"Date", "ShiftPeriod"}, "ShiftDemandHCAverage", JoinKind.LeftOuter),
    #"Expanded ShiftDemandHCAverage" = Table.ExpandTableColumn(Source, "ShiftDemandHCAverage", {"Facility", "ShiftPeriod", "ShiftDemandHCAverage"}, {"Facility", "ShiftPeriod.1", "ShiftDemandHCAverage.1"})
in
    #"Expanded ShiftDemandHCAverage";

shared DemandStart = let
    Source = #"IMPORT ShiftDemandUnitINTERVAL",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Duplicates",{{"Date", Int64.Type}}),
    #"Calculated Minimum" = List.Min(#"Changed Type"[Date])
in
    #"Calculated Minimum";

shared Table_ShiftDemandUnitINTERVALCheck = let
    Source = #"IMPORT ShiftDemandUnitINTERVAL",
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role"}, {{"CareDemandEffortRoster", each List.Sum([DemandEffort]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "CareDemandHRSWeeklyCheck", each [CareDemandEffortRoster]/2)
in
    #"Added Custom";

shared ShiftDemandHCAverageCheck = let
    Source = ShiftDemandHCAverageM,
    #"Inserted Multiplication" = Table.AddColumn(Source, "DemandHrs", each [ShiftDemandHCAverage] * [ShiftDurations.ShiftDuration]*24, type number),
    #"Grouped Rows" = Table.Group(#"Inserted Multiplication", {"Facility", "Role"}, {{"CareDemandRosterHRS", each List.Sum([DemandHrs]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "CareDemandWkHRS", each [CareDemandRosterHRS]/2)
in
    #"Added Custom";

shared CareDemandHRSWeeklyCheck = let
    Source = ShiftIntervalDemandHCAveM,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Attribute] = "StartInterval")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "CareDemandHrs", each [ShiftDemandHCAverage.1]*[Duration]*24
),
    #"Grouped Rows" = Table.Group(#"Added Custom", {"Facility", "Role"}, {{"CareDemandHRSRoster", each List.Sum([CareDemandHrs]), type nullable number}}),
    #"Added Custom1" = Table.AddColumn(#"Grouped Rows", "CareDemandHRSWeekly", each [CareDemandHRSRoster]/2)
in
    #"Added Custom1";

shared #"IMPORT RoleShiftDayDemandANACC" = let
    Source = Excel.Workbook(File.Contents(#"Folder-1Input"&"\Demand-MasterRoster Manual Read.xlsx"), null, true),
    RoleShiftDemandANACC_Table = Source{[Item="RoleShiftDemandANACC",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftDemandANACC_Table,{{"Role", type text}, {"AM", type number}, {"PM", type number}, {"NIGHT", type number}, {"Total", type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> "EN"))
in
    #"Filtered Rows";

shared RoleShiftDayDemandANACCTABLE = let
    Source = #"IMPORT RoleShiftDayDemandANACC",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Total"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"Role"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Shift"}, {"Value", "Demand"}})
in
    #"Renamed Columns";

shared ShiftDemandHCAverageANACC = let
    Source = Table.NestedJoin(#"ShiftEffort !!", {"Role", "Date", "ShiftPeriod"}, #"ShiftDurations!!", {"Role", "Date", "Shift"}, "ShiftDurations", JoinKind.LeftOuter),
    #"Expanded ShiftDurations" = Table.ExpandTableColumn(Source, "ShiftDurations", {"Duration"}, {"ShiftDurations.Duration"}),
    #"Inserted SHIFTDEMANDHCAV." = Table.AddColumn(#"Expanded ShiftDurations", "ShiftDemandHCAverage", each [UnitShiftEffort] / [ShiftDurations.Duration], type number),
    #"Sorted Rows" = Table.Sort(#"Inserted SHIFTDEMANDHCAV.",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"ShiftPeriod", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"Date", type date}})
in
    #"Changed Type";

shared ShiftIntervalDemandHCAveANACC = let
    Source = Table.NestedJoin(#"IMPORT IntervalsList", {"DayDate", "ShiftPeriod"}, ShiftDemandHCAverageANACC, {"Date", "ShiftPeriod"}, "ShiftDemandHCAverage", JoinKind.LeftOuter),
    #"Expanded ShiftDemandHCAverage" = Table.ExpandTableColumn(Source, "ShiftDemandHCAverage", {"Facility", "ShiftPeriod", "ShiftDemandHCAverage"}, {"Facility", "ShiftPeriod.1", "ShiftDemandHCAverage.1"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ShiftDemandHCAverage",{{"Role", Order.Ascending}, {"IntervalID", Order.Ascending}, {"Attribute", Order.Descending}})
in
    #"Sorted Rows";

shared DaysToRosterStart = let
    Source = #"IMPORT Table_RosterStart",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"RosterStart", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Custom", each [RosterStart]-DemandStart),
    Custom = #"Added Custom"{0}[Custom]
in
    Custom;
