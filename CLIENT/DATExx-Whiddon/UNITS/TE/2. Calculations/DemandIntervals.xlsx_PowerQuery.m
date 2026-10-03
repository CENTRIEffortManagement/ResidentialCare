// Power Query for: DemandIntervals.xlsx
// Pathname: CLIENT/DATExx-Whiddon/UNITS/TE/2. Calculations/DemandIntervals.xlsx
// Copied from: CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/DemandIntervals.xlsx_PowerQuery.m
// Source extracted: 2026-10-01T00:33:02.142Z
// Notes: Copied from the approved BD source; destination workbook queries were not re-extracted or synchronized.

section Section1;

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

// Query: ShiftDemandHours_INPUT_CHECK
// Purpose: Verify unique demand rows, valid hours and duration, and unique shift endpoint mappings before interval expansion.
// Inputs: IMPORT Table_ShiftDemandHRS; IntervalIDs.
// Output: One validation row per existing facility, role and shift-endpoint join key.
shared ShiftDemandHours_INPUT_CHECK = let
    Source = #"IMPORT Table_ShiftDemandHRS",
    RequiredColumns = {"Facility", "Role", "StartTime", "EndTime", "Shift", "Date", "DemandHRS", "ShiftDuration"},
    ValidatedSchema = if Table.HasColumns(Source, RequiredColumns) then Source
        else error "Demand interval input is missing required demand-hour or shift columns.",
    IsFiniteNumber = (value as any) as logical =>
        Value.Is(value, type number) and not Number.IsNaN(value) and Number.Abs(value) <> #infinity,
    ValidDemandRow = (row as record) as logical =>
        try (
            IsFiniteNumber(row[DemandHRS]) and row[DemandHRS] >= 0
            and IsFiniteNumber(row[ShiftDuration]) and row[ShiftDuration] > 0
            and Value.Is(row[StartTime], type datetime) and Value.Is(row[EndTime], type datetime)
            and row[EndTime] > row[StartTime]
            and Number.Abs(Duration.TotalHours(row[EndTime] - row[StartTime]) - row[ShiftDuration]) <= 0.0000001
        ) otherwise false,
    // Use the existing demand join key so duplicate rows cannot multiply interval demand.
    DemandKeys = Table.Group(ValidatedSchema, {"Facility", "Role", "StartTime", "EndTime"}, {
        {"DemandRowCount", each Table.RowCount(_), Int64.Type},
        {"DemandValuesValid", each List.AllTrue(List.Transform(Table.ToRecords(_), ValidDemandRow)), type logical}
    }),
    MatchedStartEndpoints = Table.NestedJoin(DemandKeys, {"StartTime"}, IntervalIDs, {"DateTime"}, "StartMatches", JoinKind.LeftOuter),
    MatchedEndEndpoints = Table.NestedJoin(MatchedStartEndpoints, {"EndTime"}, IntervalIDs, {"DateTime"}, "EndMatches", JoinKind.LeftOuter),
    CountedStartMatches = Table.AddColumn(MatchedEndEndpoints, "StartMatchCount", each Table.RowCount([StartMatches]), Int64.Type),
    CountedEndMatches = Table.AddColumn(CountedStartMatches, "EndMatchCount", each Table.RowCount([EndMatches]), Int64.Type),
    CheckedEndpointRanges = Table.AddColumn(CountedEndMatches, "EndpointRangeValid", each
        if [StartMatchCount] <> 1 or [EndMatchCount] <> 1 then false
        else try (
            IsFiniteNumber([StartMatches]{0}[IntervalID]) and IsFiniteNumber([EndMatches]{0}[IntervalID])
            and Number.RoundDown([StartMatches]{0}[IntervalID]) = [StartMatches]{0}[IntervalID]
            and Number.RoundDown([EndMatches]{0}[IntervalID]) = [EndMatches]{0}[IntervalID]
            and [EndMatches]{0}[IntervalID] > [StartMatches]{0}[IntervalID]
        ) otherwise false,
        type logical),
    CheckedInputs = Table.AddColumn(CheckedEndpointRanges, "Passed", each
        [DemandRowCount] = 1 and [DemandValuesValid] and [EndpointRangeValid], type logical),
    CheckResults = Table.RemoveColumns(CheckedInputs, {"StartMatches", "EndMatches"})
in
    CheckResults;

// Query: ShiftINTERVALS
// Purpose: Expand validated shift boundaries using the existing interval-ID generation rules.
// Inputs: IMPORT Table_ShiftDemandHRS; ShiftDemandHours_INPUT_CHECK; IntervalIDs.
shared ShiftINTERVALS = let
    InputFailures = Table.SelectRows(ShiftDemandHours_INPUT_CHECK, each [Passed] <> true),
    Source = if Table.IsEmpty(InputFailures) then #"IMPORT Table_ShiftDemandHRS"
        else error Error.Record("DemandIntervalInputValidation", "Demand hours, duration, uniqueness or shift endpoint validation failed.", InputFailures),
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

// Query: ShiftDemandIntervalHours_CALCULATED
// Purpose: Distribute authoritative shift hours across the existing intervals without using reporting FTE.
// Inputs: ShiftINTERVALS; Table_Intervals; IMPORT Table_ShiftDemandHRS; MealBreakStart.
// Output: Existing interval output columns plus internal IntervalHours and EndPrecise validation helpers.
shared ShiftDemandIntervalHours_CALCULATED = let
    Source = ShiftINTERVALS,
    MatchedIntervals = Table.NestedJoin(Source, {"IntervalListx", "Role"}, Table_Intervals, {"IntervalID", "Role"}, "Table_Intervals", JoinKind.LeftOuter),
    IntervalMatchFailures = Table.SelectRows(MatchedIntervals, each Table.RowCount([Table_Intervals]) <> 1),
    ValidatedIntervalMatches = if Table.IsEmpty(IntervalMatchFailures) then MatchedIntervals
        else error Error.Record("DemandIntervalMatchValidation", "Every role and interval ID must match exactly one interval.", IntervalMatchFailures),
    ExpandedIntervals = Table.ExpandTableColumn(ValidatedIntervalMatches, "Table_Intervals", {"StartInterval", "EndInterval", "Duration", "ShiftPeriod", "EndPrecise"}, {"StartInterval", "EndInterval", "Duration", "ShiftPeriod", "EndPrecise"}),
    TypedIntervalEndpoints = Table.TransformColumnTypes(ExpandedIntervals, {{"StartInterval", type datetime}, {"EndInterval", type datetime}, {"EndPrecise", type datetime}}),
    SortedIntervals = Table.Sort(TypedIntervalEndpoints, {{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}}),
    MatchedDemand = Table.NestedJoin(SortedIntervals, {"Facility", "Role", "StartTime", "EndTime"}, #"IMPORT Table_ShiftDemandHRS", {"Facility", "Role", "StartTime", "EndTime"}, "Table_ShiftDemandHRS", JoinKind.LeftOuter),
    DemandMatchFailures = Table.SelectRows(MatchedDemand, each Table.RowCount([Table_ShiftDemandHRS]) <> 1),
    ValidatedDemandMatches = if Table.IsEmpty(DemandMatchFailures) then MatchedDemand
        else error Error.Record("DemandIntervalMatchValidation", "Every demand interval must match exactly one source demand row.", DemandMatchFailures),
    ExpandedDemand = Table.ExpandTableColumn(ValidatedDemandMatches, "Table_ShiftDemandHRS", {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}, {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}),
    // Preserve legacy diagnostic columns; the demand-hour calculation applies no further meal deduction.
    AddedEffectiveDuration = Table.AddColumn(ExpandedDemand, "EffectiveDuration", each
        if [ShiftDuration] < MealBreakStart then [ShiftDuration] else [ShiftDuration] - 0.5),
    AddedIntervalEffectiveRatio = Table.AddColumn(AddedEffectiveDuration, "IntervalEffectiveRatio", each [EffectiveDuration] / [ShiftDuration]),
    // Duration is stored in days; ShiftDuration is the actual elapsed shift duration in hours.
    AddedIntervalHours = Table.AddColumn(AddedIntervalEffectiveRatio, "IntervalHours", each [Duration] * 24, type number),
    AddedDemandEffort = Table.AddColumn(AddedIntervalHours, "DemandEffort", each [DemandHRS] * ([IntervalHours] / [ShiftDuration]), type number),
    AddedAverageAttendance = Table.AddColumn(AddedDemandEffort, "EffectiveIntervalAttendance", each [DemandEffort] / [IntervalHours], type number)
in
    AddedAverageAttendance;

// Query: ShiftDemandIntervalHours_CHECK
// Purpose: Reconcile interval coverage and demand hours for every source shift before publication.
// Inputs: ShiftDemandIntervalHours_CALCULATED.
// Output: One check row per facility, role, date, shift and pair of shift endpoints; Passed must be true.
shared ShiftDemandIntervalHours_CHECK = let
    Source = ShiftDemandIntervalHours_CALCULATED,
    HourTolerance = 0.0000001,
    IsFiniteNumber = (value as any) as logical =>
        Value.Is(value, type number) and not Number.IsNaN(value) and Number.Abs(value) <> #infinity,
    CheckShiftCoverage = (rows as table) as record =>
        let
            SortedRows = Table.Sort(rows, {{"StartInterval", Order.Ascending}, {"EndPrecise", Order.Ascending}}),
            IntervalCount = Table.RowCount(SortedRows),
            UniqueIntervals = Table.Distinct(SortedRows, {"IntervalListx"}),
            IntervalRowsValid = List.AllTrue(List.Transform(Table.ToRecords(SortedRows), (row as record) =>
                try (
                    IsFiniteNumber(row[IntervalHours]) and row[IntervalHours] > 0
                    and Value.Is(row[StartInterval], type datetime) and Value.Is(row[EndInterval], type datetime)
                    and Value.Is(row[EndPrecise], type datetime) and row[EndPrecise] > row[StartInterval]
                    // EndInterval is the existing inclusive minute marker; EndPrecise is the elapsed-time boundary.
                    and row[EndInterval] = row[EndPrecise] - #duration(0, 0, 1, 0)
                    and Number.Abs(Duration.TotalHours(row[EndPrecise] - row[StartInterval]) - row[IntervalHours]) <= HourTolerance
                    and IsFiniteNumber(row[DemandEffort]) and row[DemandEffort] >= 0
                ) otherwise false)),
            // Equal total widths alone would miss a gap offset by an overlap; compare each adjacent pair as well.
            AdjacentEndpoints = List.Zip({List.RemoveLastN(SortedRows[EndPrecise], 1), List.Skip(SortedRows[StartInterval], 1)}),
            AdjacentIntervalsMeet = List.AllTrue(List.Transform(AdjacentEndpoints, each _{0} = _{1})),
            IntervalCoverageValid = try (
                IntervalCount > 0 and Table.RowCount(UniqueIntervals) = IntervalCount
                and SortedRows{0}[StartInterval] = SortedRows{0}[StartTime]
                and SortedRows{IntervalCount - 1}[EndPrecise] = SortedRows{0}[EndTime]
                and AdjacentIntervalsMeet
            ) otherwise false,
            IntervalHours = List.Sum(SortedRows[IntervalHours]),
            ShiftDuration = SortedRows{0}[ShiftDuration],
            DemandEffort = List.Sum(SortedRows[DemandEffort]),
            DemandHRS = SortedRows{0}[DemandHRS],
            HoursDifference = IntervalHours - ShiftDuration,
            DemandHoursDifference = DemandEffort - DemandHRS,
            Passed = IntervalRowsValid and IntervalCoverageValid
                and Number.Abs(HoursDifference) <= HourTolerance
                and Number.Abs(DemandHoursDifference) <= HourTolerance
        in
            [IntervalCount = IntervalCount, IntervalRowsValid = IntervalRowsValid, IntervalCoverageValid = IntervalCoverageValid,
             IntervalHours = IntervalHours, ShiftDuration = ShiftDuration, HoursDifference = HoursDifference,
             DemandEffort = DemandEffort, DemandHRS = DemandHRS, DemandHoursDifference = DemandHoursDifference, Passed = Passed],
    GroupedShiftChecks = Table.Group(Source, {"Shift", "Role", "Facility", "StartTime", "EndTime", "Date"}, {{"CoverageCheck", each CheckShiftCoverage(_), type record}}),
    CheckResults = Table.ExpandRecordColumn(GroupedShiftChecks, "CoverageCheck", {
        "IntervalCount", "IntervalRowsValid", "IntervalCoverageValid", "IntervalHours", "ShiftDuration", "HoursDifference",
        "DemandEffort", "DemandHRS", "DemandHoursDifference", "Passed"
    })
in
    CheckResults;

[ Description = "Modified for ANACC demand#(lf)Not master roster demand" ]
// Query: ShiftDemandUnitINTERVAL
// Purpose: Publish the existing interval demand interface only after coverage and hour conservation pass.
// Inputs: ShiftDemandIntervalHours_CALCULATED; ShiftDemandIntervalHours_CHECK.
shared ShiftDemandUnitINTERVAL = let
    CheckFailures = Table.SelectRows(ShiftDemandIntervalHours_CHECK, each [Passed] <> true),
    ValidatedDemand = if Table.IsEmpty(CheckFailures) then ShiftDemandIntervalHours_CALCULATED
        else error Error.Record("DemandIntervalHoursValidation", "Interval coverage or demand-hour reconciliation failed.", CheckFailures),
    RemovedValidationHelper = Table.RemoveColumns(ValidatedDemand, {"IntervalHours", "EndPrecise"}),
    // Keep the workbook-facing column order and names, including its reporting FTE and legacy meal diagnostics.
    PublishedColumns = Table.ReorderColumns(RemovedValidationHelper, {
        "Shift", "Role", "Facility", "StartTime", "EndTime", "Date", "DateTime", "DayDate", "DateTime.1", "IntervalListx",
        "StartInterval", "EndInterval", "Duration", "ShiftPeriod", "Unit", "DemandFTE", "DemandHRS", "ShiftDuration",
        "EffectiveDuration", "IntervalEffectiveRatio", "EffectiveIntervalAttendance", "DemandEffort"
    })
in
    PublishedColumns;

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
