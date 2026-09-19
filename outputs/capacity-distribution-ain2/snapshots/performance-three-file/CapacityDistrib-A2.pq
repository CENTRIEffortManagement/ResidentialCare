section Section1;

// Query: fnAddIndexedRunningTotal
// Purpose: Calculate inclusive running totals in one pass through an existing 2-based indexed priority order.
// Inputs: Ordered rows, a value column, the global index, group-start index and per-row prefix count.
// Output: The input rows with one running-total column; invalid ordering fails rather than changing allocation semantics.
shared fnAddIndexedRunningTotal = (
    orderedInput as table,
    valueColumn as text,
    indexColumn as text,
    startColumn as text,
    countColumn as text,
    outputColumn as text
) as table =>
let
    RequiredColumns = List.Distinct({valueColumn, indexColumn, startColumn, countColumn}),
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(orderedInput)),
    CheckedSchema =
        if not List.IsEmpty(MissingColumns) then
            error Error.Record("CapacityDistribA2.RunningTotalSchema", "Required running-total columns are missing.", MissingColumns)
        else if Table.HasColumns(orderedInput, {outputColumn}) then
            error Error.Record("CapacityDistribA2.RunningTotalSchema", "The running-total output column already exists.", outputColumn)
        else orderedInput,
    // Buffer only the narrow ordered calculation input and its scalar lists for this invocation.
    StableInput = Table.Buffer(CheckedSchema),
    Values = List.Buffer(Table.Column(StableInput, valueColumn)),
    Indexes = List.Buffer(Table.Column(StableInput, indexColumn)),
    GroupStarts = List.Buffer(Table.Column(StableInput, startColumn)),
    PrefixCounts = List.Buffer(Table.Column(StableInput, countColumn)),
    RowCount = Table.RowCount(StableInput),
    IsValidControl = (position as number) as logical =>
        let
            RowIndex = Indexes{position},
            GroupStart = GroupStarts{position},
            PrefixCount = PrefixCounts{position}
        in
            try (
                Value.Is(RowIndex, type number) and RowIndex = position + 2
                and Value.Is(GroupStart, type number)
                and GroupStart >= 2 and GroupStart <= RowIndex and Number.RoundDown(GroupStart) = GroupStart
                and Value.Is(PrefixCount, type number) and PrefixCount = RowIndex - GroupStart + 1
                and (if position = 0 then GroupStart = RowIndex
                     else GroupStart = GroupStarts{position - 1} or GroupStart = RowIndex)
            ) otherwise false,
    FirstInvalidPosition = List.First(List.Select(List.Numbers(0, RowCount), each not IsValidControl(_)), null),
    ValidatedInput =
        if FirstInvalidPosition = null then StableInput
        else error Error.Record(
            "CapacityDistribA2.RunningTotalOrder",
            "Running-total indexes or group boundaries are invalid. Check joins and deterministic priority order.",
            [RowPosition = FirstInvalidPosition + 1, IndexColumn = indexColumn, StartColumn = startColumn, CountColumn = countColumn]
        ),
    RunningTotals = List.Buffer(List.Generate(
        () => [Position = 0, Total = if RowCount = 0 then null else Values{0}],
        (state as record) as logical => state[Position] < RowCount,
        (state as record) as record =>
            let NextPosition = state[Position] + 1
            in [
                Position = NextPosition,
                Total = if NextPosition >= RowCount then null
                        else if GroupStarts{NextPosition} <> GroupStarts{state[Position]} then Values{NextPosition}
                        else List.Sum({state[Total], Values{NextPosition}})
            ],
        (state as record) => state[Total]
    )),
    WithRunningTotal = Table.AddColumn(
        ValidatedInput,
        outputColumn,
        (row as record) => RunningTotals{Record.Field(row, indexColumn) - 2},
        type any
    )
in
    WithRunningTotal;

[ Description = "BUFFER" ]
// Query: ResCapPrioritisedTABLE
// Purpose: Establish deterministic per-resource removal priority before the resource running total.
shared ResCapPrioritisedTABLE = let
    Source = #"Period Running Total",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "PRCell", "Reduction", "ResAv-C'", "ClusterAllocation", "NWDPriority", "NWDType", "MaxPeriodReduction", "PeriodRunningTotal", "ReducePeriod"}),
    // Period and PRCell break otherwise-equal priority rows without changing the business priority fields.
    #"Sorted RES,NWDPRI,MAXPRT,PRT" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"MaxPeriodReduction", Order.Descending}, {"PeriodRunningTotal", Order.Ascending}, {"Period", Order.Ascending}, {"PRCell", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES,NWDPRI,MAXPRT,PRT", "ResIndex", 2, 1, Int64.Type)
in
    #"Added Index";

// Query: ResCapIndexLimits
// Purpose: Capture each resource's 2-based start index and cap controls for the linear running total.
shared ResCapIndexLimits = let
    Source = ResCapPrioritisedTABLE,
    // Use the actual 2-based ResIndex; mixing a separate 1-based index caused fragile positional ranges.
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"StartResIndex", each List.Min([ResIndex]), Int64.Type}})
in
    #"Grouped Rows";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
// Query: Resource Running total
// Purpose: Apply the resource cap with one linear pass over the deterministic priority order.
shared #"Resource Running total" = let
    Source = Table.NestedJoin(ResCapPrioritisedTABLE, {"Resource"}, ResCapIndexLimits, {"Resource"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartResIndex"}, {"StartResIndex"}),
    #"Sorted Rows" = Table.Sort(#"Expanded REsLimits",{{"ResIndex", Order.Ascending}}),
    #"Added Prefix Count" = Table.AddColumn(#"Sorted Rows", "ResourcePrefixCount", each [ResIndex] - [StartResIndex] + 1, Int64.Type),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Added Prefix Count", "Reduction", "ResIndex", "StartResIndex", "ResourcePrefixCount", "ResRunningTotal"),
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "MaxResReduction", each List.Min({ [#"ResAv-C'"]})),
    #"Added KEEP" = Table.AddColumn(#"Inserted MAXRESTOTAL", "KeepRunningTotal", each if [ResRunningTotal] > [MaxResReduction] then "Remove for Res" else if [PeriodRunningTotal] > [MaxPeriodReduction] then "Remove for Perod" else if [ResRunningTotal] <= [MaxResReduction] then "Keep" else "Split"),
    #"Filtered KEEP" = Table.SelectRows(#"Added KEEP", each ([KeepRunningTotal] = "Keep")),
    #"Multiplied Column" = Table.TransformColumns(#"Filtered KEEP", {{"Reduction", each _ * -1, type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Multiplied Column",{"Resource", "Period", "PRCell", "Reduction", "ResRunningTotal"})
in
    #"Removed Other Columns";

shared CheckCap = let
    Source = #"Resource Running total",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"MAXResRunningTotal", each List.Max([ResRunningTotal]), type number}, {"AvailabilityReduction", each List.Sum([Reduction]), type number}})
in
    #"Grouped Rows";

// Query: ResPeriodCappedAvailability(C#)TABLE
// Purpose: Apply accepted cap reductions to the authoritative Resource x Period skeleton.
// Notes: Role comes from the skeleton; row order must never determine Role assignment.
shared #"ResPeriodCappedAvailability(C#)TABLE" = let
    Source = #"IMPORT ResPeriodAvailabilityTABLE",
    
    #"Merged PREMPTY" = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"IMPORT ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}, "ResourcePeriodTABLE_empty", JoinKind.RightOuter),
    #"Removed Source Keys" = Table.RemoveColumns(#"Merged PREMPTY",{"Role", "Resource", "Period"}),
    #"Expanded Authoritative Skeleton" = Table.ExpandTableColumn(#"Removed Source Keys", "ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}, {"Role", "Resource", "Period"}),
    
    #"Merged Queries" = Table.NestedJoin(#"Expanded Authoritative Skeleton", {"Resource", "Period"}, #"Resource Running total", {"Resource", "Period"}, "ReCapAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReCapAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReCapAvailabilityTABLE", {"Reduction"}, {"Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ReCapAvailabilityTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Resource", "Period"}, #"IMPORT ResPeriodWDTABLE", {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.FullOuter),
    #"Expanded ResPeriodWDTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodWDTABLE", {"RosteredPeriodStatus"}, {"RosteredPeriodStatus"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ResPeriodWDTABLE",null,0,Replacer.ReplaceValue,{"Reduction"}),
    
    #"Merged Queries2" = Table.NestedJoin(#"Replaced Value", {"Resource", "Period"}, #"IMPORT MultiPeriod - Remove", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-Remove" = Table.ExpandTableColumn(#"Merged Queries2", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    #"Inserted REDUCTION" = Table.AddColumn(#"Expanded MultiDayPeriod-Remove", "AvailabilityCapped", each if (
        [RosteredPeriodStatus] = "RosteredDay" 
        or
        [RosteredPeriodStatus] = "UnavailableAllocatedPeriod" 
        or 
        [Availability] = null
        or 
        [NoAllocationKeep] = false 
        ) 
    then 0
    else 
        [Availability] + [Reduction]),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted REDUCTION",{"Availability", "Reduction", "NoAllocationKeep"})
    in
        #"Removed Columns";

// Query: A2_COMPLETE_GRID_CHECK
// Purpose: Reconcile the capped output keys to the authoritative Resource x Period skeleton.
// Output: Connection-only diagnostic rows; every failure blocks the published C# table.
shared A2_COMPLETE_GRID_CHECK = let
    SkeletonKeys = Table.Buffer(Table.Distinct(Table.SelectColumns(#"IMPORT ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}))),
    ResultKeys = Table.Buffer(Table.SelectColumns(#"ResPeriodCappedAvailability(C#)TABLE", {"Role", "Resource", "Period"})),
    DistinctResultKeys = Table.Distinct(ResultKeys),
    MissingKeys = Table.RowCount(Table.NestedJoin(SkeletonKeys, {"Role", "Resource", "Period"}, DistinctResultKeys, {"Role", "Resource", "Period"}, "Result", JoinKind.LeftAnti)),
    ExtraKeys = Table.RowCount(Table.NestedJoin(DistinctResultKeys, {"Role", "Resource", "Period"}, SkeletonKeys, {"Role", "Resource", "Period"}, "Skeleton", JoinKind.LeftAnti)),
    DuplicateResultKeys = Table.RowCount(ResultKeys) - Table.RowCount(DistinctResultKeys),
    UnexpectedRoleRows = Table.RowCount(Table.SelectRows(ResultKeys, each [Role] <> OutputRole or [Role] = null)),
    Checks = #table(
        type table [Check = text, Status = text, Failures = Int64.Type],
        {
            {"Missing Resource x Period keys", if MissingKeys = 0 then "Pass" else "Fail", MissingKeys},
            {"Extra Resource x Period keys", if ExtraKeys = 0 then "Pass" else "Fail", ExtraKeys},
            {"Duplicate Resource x Period keys", if DuplicateResultKeys = 0 then "Pass" else "Fail", DuplicateResultKeys},
            {"Published Role equals OutputRole", if UnexpectedRoleRows = 0 then "Pass" else "Fail", UnexpectedRoleRows}
        }
    )
in
    Checks;

[ Description = "BUFFER-All cells capped availability" ]
// Query: ResPeriodAvailabilityCapped(C#)TABLE
// Purpose: Publish the compatibility C# interface only after role and complete-grid validation.
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    RoleFailures = Table.SelectRows(RoleContext_CHECK, each [Status] <> "Pass"),
    GridFailures = Table.SelectRows(A2_COMPLETE_GRID_CHECK, each [Status] <> "Pass"),
    Source = if Table.IsEmpty(RoleFailures) and Table.IsEmpty(GridFailures) then #"ResPeriodCappedAvailability(C#)TABLE"
        else error Error.Record(
            "CapacityDistribA2.Validation",
            "Role context or Resource x Period completeness failed. Review RoleContext_CHECK and A2_COMPLETE_GRID_CHECK.",
            [RoleFailures = RoleFailures, GridFailures = GridFailures]
        )
in
    Source;

shared #"ResPeriodAvailabilityCapped(C#)SUM" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    AvailabilityCapped = Source[AvailabilityCapped],
    #"Calculated Sum" = List.Sum(AvailabilityCapped)
in
    #"Calculated Sum";

shared ResPeriodAvailabilityCappedMATRIX = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Removed Columns" = Table.RemoveColumns(Source,{"RosteredPeriodStatus"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "AvailabilityCapped", List.Sum)
in
    #"Pivoted Column";

// Query: PeriodCapPrioritisedTABLE
// Purpose: Establish deterministic per-period removal priority before the period running total.
shared PeriodCapPrioritisedTABLE = let
    Source = #"IMPORT ResPeriodCapPrioritised",
    // Resource and PRCell break otherwise-equal priority rows without changing the business priority fields.
    #"Sorted NWD,C-D,PA-Cn,PERIOD" = Table.Sort(Source,{{"Period", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"ClusterAllocation", Order.Descending}, {"C'-D", Order.Ascending}, {"PeriodA-CNeg", Order.Ascending}, {"Resource", Order.Ascending}, {"PRCell", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted NWD,C-D,PA-Cn,PERIOD", "PeriodIndex", 2, 1, Int64.Type)
in
    #"Added Index";

// Query: PeriodCapIndexLimits
// Purpose: Capture each period's 2-based start index and cap controls for the linear running total.
shared PeriodCapIndexLimits = let
    Source = PeriodCapPrioritisedTABLE,
    // Use the actual 2-based PeriodIndex rather than the former parallel 1-based index.
    #"Grouped INDEX+COUNT" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([PeriodIndex]), Int64.Type}})
in
    #"Grouped INDEX+COUNT";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
// Query: Period Running Total
// Purpose: Apply the period cap with one linear pass over the deterministic priority order.
shared #"Period Running Total" = let
    Source = Table.NestedJoin(PeriodCapPrioritisedTABLE, {"Period"}, PeriodCapIndexLimits, {"Period"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded PeriodLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartPeriodIndex"}, {"StartPeriodIndex"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PeriodLimits",{{"PeriodIndex", Order.Ascending}}),
    // Min reduction: PeriodA-C, PeriodC-D
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Sorted Rows", "MaxPeriodReduction", each List.Min({-1*[#"PeriodA-CNeg"], [#"C'-D"]})),
    #"Added Prefix Count" = Table.AddColumn(#"Inserted MAXRESTOTAL", "PeriodPrefixCount", each [PeriodIndex] - [StartPeriodIndex] + 1, Int64.Type),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Added Prefix Count", "UnassignedAvail", "PeriodIndex", "StartPeriodIndex", "PeriodPrefixCount", "PeriodRunningTotal"),
    #"Renamed Columns" = Table.RenameColumns(#"Added RUNNINGRESTOTAL",{{"UnassignedAvail", "Reduction"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ReducePeriod", each if [PeriodRunningTotal] <= [MaxPeriodReduction] then "RemovableByPeriod" else "UnremovableByPeriod")
in
    #"Added Conditional Column";

shared ResPeriodAvailabilityCappedMATRIXSUM = let
    Source = ResPeriodAvailabilityCappedMATRIX,
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Resource","Role"}, "Attribute", "Value"),
    Value = #"Unpivoted Other Columns"[Value],
    #"Calculated Sum" = List.Sum(Value)
in
    #"Calculated Sum";

// Query: RolePathTABLE
// Purpose: Resolve this workbook's folder and dynamic role through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value rows used by RolePath and the A.2 imports.
shared RolePathTABLE =
let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    FilePathUrl =
    let
        // Accept the legacy FilePAthUrl casing without masking a missing or malformed input.
        PathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
        Source = if Table.RowCount(PathInputs) = 1 then PathInputs{0}[Content]
            else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
        // A single named cell is exposed by Excel as Column1.
        PathColumn = if Table.HasColumns(Source, {"FilePath"}) then Table.SelectColumns(Source, {"FilePath"})
            else if Table.ColumnNames(Source) = {"Column1"} then Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must contain a FilePath column or be a single named cell.",
        ValidatedRows = if Table.RowCount(PathColumn) = 1 then PathColumn
            else error "FilePathUrl must contain exactly one data row.",
        TypedPath = Table.TransformColumnTypes(ValidatedRows, {{"FilePath", type text}}),
        BufferedTable = Table.Buffer(TypedPath)
    in
        BufferedTable,

    RawFilePath = NormalizePath(FilePathUrl{0}[FilePath]),
    NonBlankFilePath = if RawFilePath = null or RawFilePath = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else RawFilePath,
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; retain only the workbook path.
    WorkbookPath = if Text.Contains(NonBlankFilePath, "[") then
        Text.BeforeDelimiter(NonBlankFilePath, "[") & Text.BetweenDelimiters(NonBlankFilePath, "[", "]")
        else NonBlankFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "CapacityDistrib(A.2)-shifts.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate CapacityDistrib(A.2)-shifts.xlsx.",
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    // Match full roots for both SharePoint and local paths; no Site-name join is required.
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
            // A root must end at a folder boundary, not part-way through a different folder name.
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    LongestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0}
        else error "FilePathUrl has no usable CentriSyncPaths mapping. Check SharepointRootUrl and SyncedFolderRootPath for: " & FilePath,
    // Duplicate matching rows may agree, but conflicting destinations must not depend on row order.
    BestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    BestMatch = if Table.IsEmpty(SortedMatches) then LongestMatch
        else if List.Count(List.Distinct(BestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase)) = 1 then LongestMatch
        else error "CentriSyncPaths contains conflicting local folders for: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    FolderSegments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    UnitsIndex = List.PositionOf(FolderSegments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    CalcIndex = List.PositionOf(FolderSegments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    // Imports require UNITS/<unit>/2. Calculations/<role>; unit and role names remain dynamic.
    RootPath = if UnitsIndex >= 0
        and CalcIndex = UnitsIndex + 2
        and List.Count(FolderSegments) = CalcIndex + 2
        and Text.Trim(FolderSegments{UnitsIndex + 1}) <> ""
        and Text.Trim(FolderSegments{CalcIndex + 1}) <> "" then WorkbookFolder
        else error "FilePathUrl must identify this workbook under UNITS\<unit>\2. Calculations\<role>. Save and recalculate its path input.",
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    // Preserve the role folder name so copies into another role folder use that role automatically.
    Role = Segments{CalcIndex + 1},
    FileName = InputFileName,
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"Role", Role},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

// Query: RoleFolder
// Purpose: Expose the physical calculation-folder name resolved from this workbook's own FilePathUrl.
shared RoleFolder = let
    MatchingRows = Table.SelectRows(RolePathTABLE, each [Variable Name] = "Role"),
    Value = if Table.RowCount(MatchingRows) = 1 then MatchingRows{0}[Value]
        else error "RolePathTABLE must contain exactly one physical Role row."
in
    Value;

// Query: SourceRole
// Purpose: Keep AIN as the upstream business role while the isolated workbook physically runs in AIN2.
shared SourceRole =
    if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0 or Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN"
    else error Error.Record("CapacityDistribA2.UnsupportedRoleFolder", "This AIN build must run only from physical folder AIN or AIN2.", RoleFolder);

// Query: OutputRole
// Purpose: Publish AIN2 labels during development and AIN labels automatically after the approved folder rename.
shared OutputRole =
    if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN2"
    else if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0 then "AIN"
    else error Error.Record("CapacityDistribA2.UnsupportedRoleFolder", "This AIN build must run only from physical folder AIN or AIN2.", RoleFolder);

// Query: Role
// Purpose: Preserve the existing compatibility scalar while publishing the explicit output role.
shared Role = OutputRole;

// Query: RoleContext_CHECK
// Purpose: Validate the only two supported physical/source/output tuples for this AIN build.
// Output: Connection-only check rows forced by the published C# interface.
shared RoleContext_CHECK = let
    ExpectedSourceRole = "AIN",
    ExpectedOutputRole = if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN2" else "AIN",
    SupportedFolder = Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 or Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0,
    SourceAttempt = try SourceRole,
    OutputAttempt = try OutputRole,
    ValidTuple = SupportedFolder
        and not SourceAttempt[HasError]
        and not OutputAttempt[HasError]
        and SourceAttempt[Value] = ExpectedSourceRole
        and OutputAttempt[Value] = ExpectedOutputRole,
    Checks = #table(
        type table [Check = text, Status = text, Actual = nullable text, Expected = nullable text],
        {
            {"Supported RoleFolder", if SupportedFolder then "Pass" else "Fail", RoleFolder, "AIN2 or AIN"},
            {"SourceRole", if SourceRole = ExpectedSourceRole then "Pass" else "Fail", SourceRole, ExpectedSourceRole},
            {"OutputRole", if OutputRole = ExpectedOutputRole then "Pass" else "Fail", OutputRole, ExpectedOutputRole}
        }
    ),
    Checked = if ValidTuple then Checks
        else error Error.Record(
            "CapacityDistribA2.RoleContext",
            "The physical/source/output role tuple is unsupported.",
            [RoleFolder = RoleFolder, ExpectedSourceRole = ExpectedSourceRole, ExpectedOutputRole = ExpectedOutputRole]
        )
in
    Checked;

// Query: RolePath
// Purpose: Preserve the physical workbook folder for same-lane A.1 imports.
shared RolePath = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    Folder = #"Renamed Columns"{0}[Folder]
in
    Folder;

// Query: IMPORTSource A1
// Purpose: Navigate the physical-lane A.1 workbook once for all five A.2 imports.
// Notes: The binary/navigation buffers are local to an evaluation and avoid five independent File.Contents definitions.
shared #"IMPORTSource A1" = let
    WorkbookPath = RolePath & "\CapacityDistrib(A.1)-shifts.xlsx",
    CandidateFolder = Text.BeforeDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    CheckedPath = if Comparer.OrdinalIgnoreCase(CandidateFolder, RolePath) = 0 then WorkbookPath
        else error Error.Record("CapacityDistribA2.IntraLaneEscape", "The A.1 import resolved outside the physical role folder.", WorkbookPath),
    WorkbookBinary = Binary.Buffer(File.Contents(CheckedPath)),
    WorkbookNavigation = Table.Buffer(Excel.Workbook(WorkbookBinary, null, true))
in
    WorkbookNavigation;

// Query: IMPORT ResPeriodCapPrioritised
// Purpose: Read the narrow A.1 period-cap priority interface used by A.2.
shared #"IMPORT ResPeriodCapPrioritised" = let
    Source = #"IMPORTSource A1",
    ResPeriodCapPrioritised_Table = Source{[Item="ResPeriodCapPrioritised",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodCapPrioritised_Table, {"Resource", "Period", "Allocation", "PotentialAvailability", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "C'/D", "ClusterAllocation", "NWDPriority", "NWDType"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Allocation", type any}, {"PotentialAvailability", type any}, {"PRCell", type text}, {"UnassignedAvail", Int64.Type}, {"ResAv-C'", Int64.Type}, {"PeriodA-CNeg", type number}, {"C'-D", Int64.Type}, {"C'/D", type number}, {"ClusterAllocation", type number}, {"NWDPriority", Int64.Type}, {"NWDType", type text}})
in
    #"Changed Type";

// Query: IMPORT MultiPeriod - Remove
// Purpose: Read only the A.1 multi-period keep flag used by the final capped-availability rule.
shared #"IMPORT MultiPeriod - Remove" = let
    Source = #"IMPORTSource A1",
    MultiDayPeriod_Remove_Table = Source{[Item="MultiDayPeriod_Remove",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(MultiDayPeriod_Remove_Table, {"Resource", "AvailablePeriod", "NoAllocationKeep"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"NoAllocationKeep", type logical}})
in
    #"Changed Type";

// Query: IMPORT ResPeriodWDTABLE
// Purpose: Read only roster status and availability at Resource x Period grain from A.1.
shared #"IMPORT ResPeriodWDTABLE" = let
    Source = #"IMPORTSource A1",
    ResPeriodWDTABLE_Table = Source{[Item="ResPeriodWDTABLE",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodWDTABLE_Table, {"Resource", "Period", "Availability", "RosteredPeriodStatus"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type";

// Query: IMPORT ResPeriodAvailabilityTABLE
// Purpose: Read A.1 availability at Role x Resource x Period grain.
shared #"IMPORT ResPeriodAvailabilityTABLE" = let
    Source = #"IMPORTSource A1",
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodAvailabilityTABLE_Table, {"Role", "Resource", "Period", "Availability"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORT ResourcePeriodTABLE_empty
// Purpose: Read the authoritative narrow Role x Resource x Period skeleton from A.1.
shared #"IMPORT ResourcePeriodTABLE_empty" = let
    Source = #"IMPORTSource A1",
    ResourcePeriodTABLE_empty_Table = Source{[Item="ResourcePeriodTABLE_empty",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResourcePeriodTABLE_empty_Table, {"Role", "Resource", "Period"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";
