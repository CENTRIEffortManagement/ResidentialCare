// Power Query from: AllocationByShiftAverage.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\AllocationByShiftAverage.xlsx
// Extracted: 2026-09-24T05:33:16.188Z

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

// Query: RoleShiftAllocation_StartEffort
// Purpose: Aggregate each resource interval once at date, role and shift grain.
// Inputs: ResourceIntervalAllocation; effort is measured in days.
// Output: One row per ShiftDate, ShiftPeriod and Role with Attribute = IntervalStart.
shared RoleShiftAllocation_StartEffort = let
    // Shifts exposes both endpoints and a Role1 expansion. Count only the matching role at the start.
    AllocationStarts = Table.SelectRows(ResourceIntervalAllocation, each [StaffCount] = 1 and [Attribute] = "IntervalStart"),
    EffortByRoleShift = Table.Group(AllocationStarts, {"ShiftDate", "ShiftPeriod", "Role"}, {{"RoleShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}}),
    WithEndpoint = Table.AddColumn(EffortByRoleShift, "Attribute", each "IntervalStart", type text)
in
    WithEndpoint;

// Query: RoleShiftAllocation_Calculated
// Purpose: Match each role shift to its own duration and calculate shift-equivalent FTE.
// Output: One row per date, role and shift, including DurationMatchCount for validation.
shared RoleShiftAllocation_Calculated = let
    // Role is part of the duration key: matching ShiftPeriod alone fans out across all roles.
    MatchedDurations = Table.NestedJoin(RoleShiftAllocation_StartEffort, {"Role", "ShiftPeriod"}, #"IMPORT ShiftPeriod", {"Role", "ShiftPeriod"}, "ShiftDurationMatches", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(MatchedDurations, "DurationMatchCount", each Table.RowCount([ShiftDurationMatches]), Int64.Type),
    // Keep ambiguous/missing matches visible to the check query without multiplying output rows.
    WithDuration = Table.AddColumn(WithMatchCount, "DurationOfShifts", each if [DurationMatchCount] = 1 then [ShiftDurationMatches]{0}[DurationOfShifts] else null, type nullable number),
    WithoutNestedMatches = Table.RemoveColumns(WithDuration, {"ShiftDurationMatches"}),
    // Effective effort is in days; convert to hours, then divide by the role's shift hours.
    WithFTE = Table.AddColumn(WithoutNestedMatches, "RoleShiftFTE", each
        if [DurationOfShifts] <> null and [DurationOfShifts] > 0 and [DurationOfShifts] < #infinity
        then [RoleShiftEffort] * 24 / [DurationOfShifts]
        else null, type nullable number)
in
    WithFTE;

// Query: RoleShiftAllocation
// Purpose: Publish validated allocation at date, role and shift grain.
// Notes: Preserve the existing seven-column interface; required check failures block publication.
shared RoleShiftAllocation = let
    FailedChecks = Table.SelectRows(RoleShiftAllocation_CHECK, each not [Passed]),
    ValidatedAllocation = if Table.IsEmpty(FailedChecks) then RoleShiftAllocation_Calculated
        else error Error.Record("RoleShiftAllocation.Validation", "Allocation validation failed. Inspect RoleShiftAllocation_CHECK.", FailedChecks),
    OutputColumns = Table.SelectColumns(ValidatedAllocation, {"ShiftDate", "ShiftPeriod", "Role", "Attribute", "RoleShiftEffort", "DurationOfShifts", "RoleShiftFTE"}),
    SortedAllocation = Table.Sort(OutputColumns, {{"Role", Order.Ascending}, {"ShiftDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    SortedAllocation;

// Query: RoleShiftIntervalAllocation
// Purpose: Supply the existing Table_RoleShiftIntervalAllocation load with validated shift totals.
// Notes: Retain this query/table binding for Allocation.xlsx; endpoints are not separate allocations.
shared RoleShiftIntervalAllocation = let
    SortedAllocation = Table.Sort(RoleShiftAllocation, {{"ShiftDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}, {"Role", Order.Ascending}})
in
    SortedAllocation;

// Query: RoleShiftAllocation_CHECK
// Purpose: Expose the key, duration, completeness and FTE invariants required by both outputs.
// Inputs: Staging queries only, so output validation does not create a dependency cycle.
// Output: One row per check with a failure count and Passed flag; keep connection-only when syncing.
shared RoleShiftAllocation_CHECK = let
    AllocationRows = RoleShiftAllocation_Calculated,
    IsNonnegativeFinite = (Value as any) as logical =>
        if Value = null then false else (try Value >= 0 and Value < #infinity otherwise false),
    KeyCounts = Table.Group(AllocationRows, {"ShiftDate", "ShiftPeriod", "Role"}, {{"RowsPerKey", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.RowCount(Table.SelectRows(KeyCounts, each [RowsPerKey] <> 1)),
    InvalidKeys = Table.RowCount(Table.SelectRows(AllocationRows, each [ShiftDate] = null or [Role] = null or [ShiftPeriod] = null or Text.Trim([Role]) = "" or Text.Trim([ShiftPeriod]) = "")),
    InvalidEndpoints = Table.RowCount(Table.SelectRows(AllocationRows, each [Attribute] <> "IntervalStart")),
    InvalidDurations = Table.RowCount(Table.SelectRows(AllocationRows, each [DurationMatchCount] <> 1 or not IsNonnegativeFinite([DurationOfShifts]) or [DurationOfShifts] = 0)),
    InvalidMeasures = Table.RowCount(Table.SelectRows(AllocationRows, each not IsNonnegativeFinite([RoleShiftEffort]) or not IsNonnegativeFinite([RoleShiftFTE]))),
    // Reconcile in hours with a relative tolerance for floating-point arithmetic.
    UnreconciledRows = Table.RowCount(Table.SelectRows(AllocationRows, each
        if not IsNonnegativeFinite([RoleShiftEffort]) or not IsNonnegativeFinite([RoleShiftFTE]) or not IsNonnegativeFinite([DurationOfShifts]) then true
        else Number.Abs([RoleShiftFTE] * [DurationOfShifts] - [RoleShiftEffort] * 24) > 0.000000001 * List.Max({1, Number.Abs([RoleShiftEffort] * 24)}))),
    MissingStartKeys = Table.RowCount(Table.NestedJoin(RoleShiftAllocation_StartEffort, {"ShiftDate", "ShiftPeriod", "Role"}, AllocationRows, {"ShiftDate", "ShiftPeriod", "Role"}, "AllocationMatch", JoinKind.LeftAnti)),
    CheckResults = #table(type table [Check = text, Failures = Int64.Type], {
        {"Unique date-role-shift keys", DuplicateKeys},
        {"Required date-role-shift keys", InvalidKeys},
        {"Interval starts only", InvalidEndpoints},
        {"One positive finite duration for each role-shift", InvalidDurations},
        {"Nonnegative finite effort and FTE", InvalidMeasures},
        {"FTE reconciles to effective hours", UnreconciledRows},
        {"All start-effort keys retained", MissingStartKeys}
    }),
    WithStatus = Table.AddColumn(CheckResults, "Passed", each [Failures] = 0, type logical)
in
    WithStatus;


shared RoleShiftEffortFTE = let
    Source = ResourceIntervalAllocation,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayDate", type date}})
in
    #"Changed Type";
