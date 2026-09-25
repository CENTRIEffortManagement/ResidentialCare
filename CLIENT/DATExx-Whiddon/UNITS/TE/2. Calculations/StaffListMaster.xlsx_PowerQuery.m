// Power Query from: StaffListMaster.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\StaffListMaster.xlsx
// Extracted: 2026-09-24T22:56:49.656Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve the current workbook to its local Unit1 Calculations folder using FilePathUrl and CentriSyncPaths.
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

// Query: FilePath-2Calculations
// Purpose: Expose the resolved Unit1 Calculations folder for all local workbook imports.
shared #"FilePath-2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: StaffListText
// Purpose: Normalize imported identifiers and descriptive text without converting text identifiers to numbers.
shared StaffListText = (value as any) as nullable text =>
    let Clean = try Text.Clean(Text.Trim(Text.From(value))) otherwise null
    in if Clean = "" then null else Clean;

// Query: StaffListCheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared StaffListCheckResult = (name as text, evaluate as function) as record =>
    let Result = try evaluate()
    in [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
        Failures = if Result[HasError] then null else Result[Value],
        Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

// Query: Masterlist Join
// Purpose: Combine availability staff first with allocation-only staff while retaining Step 1 contract fields.
shared #"Masterlist Join" = let
    Source = AvailableStaffList,
    #"Appended Query" = Table.Buffer(Table.Combine({Source, AllocationTable_StaffList})),
    // Availability rows are first so an existing worker keeps employee and contract fields when also allocated.
    #"Removed Duplicates" = Table.Distinct(#"Appended Query", {"Name"})
in
    #"Removed Duplicates";

// Query: Masterlist
// Purpose: Publish the validated Table_Masterlist interface with employee contracts and effective shift caps.
// Output: One row per Resource; this query remains the existing loaded Table_Masterlist source.
shared Masterlist = let
    Failures = Table.SelectRows(ResourceContract_CHECK, each [Status] = "Fail"),
    Checked = if Table.IsEmpty(Failures) then ResourceContract
        else error Error.Record("Resource contract validation", "Required resource contract checks failed.", Failures),
    Output = Table.SelectColumns(Checked,
        {"Name", "Role", "Resource", "Misalignment", "Source", "EmployeeID", "Facility-Abbrev", "EmploymentType", "PreferredRole",
         "Worker Record Status", "Worker Contract Issue",
         "Contracted FN Hours", "Roster Start", "Roster End", "Roster Fortnights", "Contracted Roster Hours",
         "Contracted Shift Equivalent", "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"})
in Table.Sort(Output, {{"Resource", Order.Ascending}});

// Query: Masterlist_Base
// Purpose: Preserve the existing resource assignment and name-misalignment result before contract calculations.
shared Masterlist_Base = let
    Source = #"Masterlist Join",
    #"Added RESOURCE INDEX" = Table.AddIndexColumn(Source, "Resource", 1, 1, Int64.Type),
    #"Merged Queries" = Table.NestedJoin(#"Added RESOURCE INDEX", {"Name", "Role"}, #"Misaligned Join", {"Misaligned-Name", "Misaligned-Role"}, "Misaligned Join", JoinKind.FullOuter),
    #"Expanded Misaligned Join" = Table.ExpandTableColumn(#"Merged Queries", "Misaligned Join", {"Misalignment"}, {"Misalignment"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Misaligned Join",{{"Misalignment", Order.Ascending}})
in
    #"Sorted Rows";

// Query: ResourceContract
// Purpose: Calculate roster contract hours and the effective availability cap at one row per Resource.
// Notes: Contracted shifts are floored only for capacity. Exact roster hours and shift equivalents remain available for reporting.
shared ResourceContract = let
    Source = Masterlist_Base,
    RosterDates = #"EXTRACT StaffList RosterDates",
    RosterDateCount = Table.RowCount(RosterDates),
    RosterStart = List.Min(RosterDates[Date]),
    RosterEnd = List.Max(RosterDates[Date]),
    RosterFortnights = Number.From(RosterDateCount) / 14,
    ShiftDuration = #"EXTRACT StaffList ShiftDuration",
    SettingsMaximum = #"EXTRACT StaffList MaxAvailability",
    WithRosterStart = Table.AddColumn(Source, "Roster Start", each RosterStart, type date),
    WithRosterEnd = Table.AddColumn(WithRosterStart, "Roster End", each RosterEnd, type date),
    WithRosterFortnights = Table.AddColumn(WithRosterEnd, "Roster Fortnights", each RosterFortnights, type number),
    WithContractRosterHours = Table.AddColumn(WithRosterFortnights, "Contracted Roster Hours", each
        if [Contracted FN Hours] = null then null else [Contracted FN Hours] * RosterFortnights, type nullable number),
    WithShiftEquivalent = Table.AddColumn(WithContractRosterHours, "Contracted Shift Equivalent", each
        if [Contracted Roster Hours] = null then null else [Contracted Roster Hours] / ShiftDuration, type nullable number),
    WithContractedShifts = Table.AddColumn(WithShiftEquivalent, "Contracted Shifts", each
        if [Contracted Shift Equivalent] = null then null else Int64.From(Number.RoundDown([Contracted Shift Equivalent])), Int64.Type),
    WithSettingsMaximum = Table.AddColumn(WithContractedShifts, "Settings Max Availability", each SettingsMaximum, Int64.Type),
    WithCasualFallback = Table.AddColumn(WithSettingsMaximum, "Is Casual Fallback", each
        [Source] = "Availability"
            and [EmploymentType] <> null
            and Text.Contains(Text.Upper([EmploymentType]), "CASUAL")
            and ([Contracted FN Hours] = null or [Contracted FN Hours] = 0), type logical),
    WithMissingWorkerFallback = Table.AddColumn(WithCasualFallback, "Is Missing Worker Fallback", each
        [Source] = "Allocation" and [Worker Record Status] = "Not found", type logical),
    WithEffectiveCap = Table.AddColumn(WithMissingWorkerFallback, "Effective Shift Cap", each
        if [Is Missing Worker Fallback] then [Settings Max Availability]
        else if [Source] <> "Availability" then null
        else if [Is Casual Fallback] then [Settings Max Availability]
        else if [Contracted Shifts] = null or [Contracted Shifts] <= 0 then null
        else List.Min({[Settings Max Availability], [Contracted Shifts]}), type nullable number),
    WithLimitBasis = Table.AddColumn(WithEffectiveCap, "Limit Basis", each
        if [Is Missing Worker Fallback] then "Worker record missing - Settings maximum fallback"
        else if [Source] <> "Availability" then "Allocation-only resource"
        else if [Is Casual Fallback] then "Casual fallback - Settings maximum"
        else if [Contracted Shifts] = null or [Contracted Shifts] <= 0 then "Invalid or missing contract"
        else if [Contracted Shifts] < [Settings Max Availability] then "Contracted shifts"
        else "Settings maximum", type text)
in Table.Buffer(WithLimitBasis);

// Query: ResourceContract_CHECK
// Purpose: Validate settings, employee coverage, contract policy and one-row-per-Resource output.
shared ResourceContract_CHECK = let
    Contracts = ResourceContract,
    AvailabilityInput = #"EXTRACT Availability Staff List",
    AvailabilityResources = Table.SelectRows(Contracts, each [Source] = "Availability"),
    AllocationResources = Table.SelectRows(Contracts, each [Source] = "Allocation"),
    MissingWorkerRows = Table.SelectRows(AllocationResources, each [Worker Record Status] = "Not found"),
    InvalidAllocationWorkerRows = Table.SelectRows(AllocationResources, each
        [Worker Record Status] <> null and Text.StartsWith([Worker Record Status], "Found -", Comparer.OrdinalIgnoreCase)),
    BlockingChecks = {
        StaffListCheckResult("Valid roster and cap settings", () =>
            if Table.RowCount(#"EXTRACT StaffList RosterDates") > 0
                and #"EXTRACT StaffList ShiftDuration" > 0
                and #"EXTRACT StaffList MaxAvailability" > 0 then 0 else 1),
        StaffListCheckResult("Unique Resource rows", () => Table.RowCount(Contracts) -
            Table.RowCount(Table.Distinct(Contracts, {"Resource"}))),
        StaffListCheckResult("Populated Resource identifiers", () => Table.RowCount(Table.SelectRows(Contracts, each [Resource] = null))),
        StaffListCheckResult("Unique availability employee and facility input", () => Table.RowCount(AvailabilityInput) -
            Table.RowCount(Table.Distinct(AvailabilityInput, {"EmployeeID", "Facility-Abbrev"}))),
        StaffListCheckResult("Unique availability names", () => Table.RowCount(AvailabilityInput) -
            Table.RowCount(Table.Distinct(AvailabilityInput, {"Name"}))),
        StaffListCheckResult("Unique preferred employee and facility resources", () => Table.RowCount(AvailabilityResources) -
            Table.RowCount(Table.Distinct(AvailabilityResources, {"EmployeeID", "Facility-Abbrev"}))),
        StaffListCheckResult("Complete availability employee mapping", () => Table.RowCount(Table.SelectRows(AvailabilityResources, each
            [EmployeeID] = null or [#"Facility-Abbrev"] = null or [EmploymentType] = null
            or [PreferredRole] = null or [Role] <> [PreferredRole]))),
        StaffListCheckResult("Valid availability contracts", () => Table.RowCount(Table.SelectRows(AvailabilityResources, each
            let IsCasual = [EmploymentType] <> null and Text.Contains(Text.Upper([EmploymentType]), "CASUAL")
            in ([Contracted FN Hours] <> null and [Contracted FN Hours] < 0)
                or (not IsCasual and ([Contracted FN Hours] = null or [Contracted FN Hours] <= 0))))),
        StaffListCheckResult("Calculated resource contract caps", () => Table.RowCount(Table.SelectRows(AvailabilityResources, each
            let
                ExpectedRosterHours = if [Contracted FN Hours] = null then null else [Contracted FN Hours] * [Roster Fortnights],
                ExpectedEquivalent = if ExpectedRosterHours = null then null else ExpectedRosterHours / #"EXTRACT StaffList ShiftDuration",
                ExpectedShifts = if ExpectedEquivalent = null then null else Number.RoundDown(ExpectedEquivalent),
                ExpectedCap = if [Is Casual Fallback] then [Settings Max Availability]
                    else if ExpectedShifts = null or ExpectedShifts <= 0 then null
                    else List.Min({[Settings Max Availability], ExpectedShifts})
            in ([Contracted Roster Hours] <> null and ExpectedRosterHours <> null and Number.Abs([Contracted Roster Hours] - ExpectedRosterHours) > 0.00000001)
                or ([Contracted Shift Equivalent] <> null and ExpectedEquivalent <> null and Number.Abs([Contracted Shift Equivalent] - ExpectedEquivalent) > 0.00000001)
                or [Contracted Shifts] <> ExpectedShifts
                or [Effective Shift Cap] <> ExpectedCap
                or [Effective Shift Cap] = null)))
    },
    // Allocation-only worker-record gaps remain visible without blocking publication of the fallback cap.
    DiagnosticChecks = {
        [Check = "Allocation Resources missing Worker's Report", Status = if Table.IsEmpty(MissingWorkerRows) then "Pass" else "Warning",
            Failures = Table.RowCount(MissingWorkerRows), Details = "Filter Table_Masterlist where Worker Record Status = Not found."],
        [Check = "Allocation Resources with unusable Worker contract", Status = if Table.IsEmpty(InvalidAllocationWorkerRows) then "Pass" else "Warning",
            Failures = Table.RowCount(InvalidAllocationWorkerRows), Details = "Filter Table_Masterlist where Worker Record Status begins Found -."],
        [Check = "Missing-worker fallback uses Settings maximum", Status = if Table.IsEmpty(Table.SelectRows(MissingWorkerRows, each
                [Effective Shift Cap] <> [Settings Max Availability]
                or [Contracted FN Hours] <> null
                or [Contracted Roster Hours] <> null
                or [Contracted Shifts] <> null)) then "Pass" else "Fail",
            Failures = Table.RowCount(Table.SelectRows(MissingWorkerRows, each
                [Effective Shift Cap] <> [Settings Max Availability]
                or [Contracted FN Hours] <> null
                or [Contracted Roster Hours] <> null
                or [Contracted Shifts] <> null)),
            Details = "Fallback affects capacity only and must not fabricate contract hours."]
    },
    Checks = BlockingChecks & DiagnosticChecks
in Table.Buffer(Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]));

// Query: AllocationTable_StaffList
// Purpose: Prepare the allocation workbook's staff identity list for the existing master-list union.
shared AllocationTable_StaffList = let
    Source = #"EXTRACT Allocation Staff With Code",
    JoinedWorker = Table.NestedJoin(Source, {"EmployeeID"}, #"StaffList WorkerContracts", {"EmployeeID"}, "WorkerContract", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(JoinedWorker, "Worker Record Matches", each Table.RowCount([WorkerContract]), Int64.Type),
    ExpandedWorker = Table.ExpandTableColumn(WithMatchCount, "WorkerContract",
        {"Contracted FN Hours", "Worker Contract Issue"}, {"Contracted FN Hours", "Worker Contract Issue"}),
    WithWorkerStatus = Table.AddColumn(ExpandedWorker, "Worker Record Status", each
        if [Worker Record Matches] = 0 then "Not found"
        else if [Worker Contract Issue] <> null then "Found - " & [Worker Contract Issue]
        else "Found", type text),
    RemovedMatchCount = Table.RemoveColumns(WithWorkerStatus, {"Worker Record Matches"}),
    WithSource = Table.AddColumn(RemovedMatchCount, "Source", each "Allocation", type text),
    Sorted = Table.Sort(WithSource, {{"Name", Order.Ascending}, {"Role", Order.Ascending}})
in
    Sorted;

// Query: AvailableStaffListInitial
// Purpose: Retain Step 1 employee, preferred-role and contract fields before applying the legacy display-name cleanup.
shared AvailableStaffListInitial = let
    Source = #"EXTRACT Availability Staff List",
    #"Trimmed Text" = Table.TransformColumns(Source,{{"Role", Text.Trim, type text}, {"Name", Text.Trim, type text}}),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"Role", Text.Clean, type text}, {"Name", Text.Clean, type text}})
in
    #"Cleaned Text";

// Query: AvailableStaffList
// Purpose: Preserve the existing display-name cleanup while carrying the employee contract columns unchanged.
shared AvailableStaffList = let
     Source1 = AvailableStaffListInitial,
    #"Added Custom" = Table.AddColumn(Source1, "Source", each "Availability"),
    #"Added Worker Record Status" = Table.AddColumn(#"Added Custom", "Worker Record Status", each "Found", type text),
    #"Added Worker Contract Issue" = Table.AddColumn(#"Added Worker Record Status", "Worker Contract Issue", each null, type nullable text),
    #"Duplicated Column" = Table.DuplicateColumn(#"Added Worker Contract Issue", "Name", "Name.A"),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Duplicated Column", "Name", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Name.1", "Name.2", "Name.3", "Name.4", "Name.5", "Name.6"}),
    #"Removed Columns" = Table.RemoveColumns(#"Split Column by Delimiter",{"Name.4", "Name.5", "Name.6"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns",null,"",Replacer.ReplaceValue,{"Name.2", "Name.3"}),
    #"Replaced NANDEEENI" = Table.ReplaceValue(#"Replaced Value","Muni","",Replacer.ReplaceText,{"Name.2"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced NANDEEENI",{"Name.2", "Name.3"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Merged"),
    #"Trimmed Text" = Table.TransformColumns(#"Merged Columns",{{"Merged", Text.Trim, type text}}),
    #"Merged Columns1" = Table.CombineColumns(#"Trimmed Text",{"Name.1", "Merged"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Name"),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Merged Columns1", "Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Name.1", "Name.2"}),
    #"Renamed Columns" = Table.RenameColumns(#"Split Column by Delimiter1",{{"Name.A", "Name"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Name", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Name.1", "Name.2"})
in
    #"Removed Columns1";

// Query: IMPORT StaffList Worker's Report
// Purpose: Read Unit1 employee contract rows once for allocation-only worker-record coverage.
// Inputs: Table1 in 1. Input/Worker's Report.xlsx.
shared #"IMPORT StaffList Worker's Report" = let
    Navigation = Excel.Workbook(File.Contents(#"FilePath-1Input" & "\Worker's Report.xlsx"), null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "Table1" and [Kind] = "Table"),
    ContractTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("StaffList worker contract import", "Expected exactly one Table1 table in Worker's Report.xlsx.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Employee Code", "Contracted FN Hours"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(ContractTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(ContractTable, RequiredColumns)
        else error Error.Record("StaffList worker contract import", "Worker's Report.xlsx/Table1 is missing required columns.", [MissingColumns = MissingColumns])
in Table.Buffer(Selected);

// Query: StaffList WorkerContracts
// Purpose: Resolve Worker's Report to one auditable contract row per employee for allocation-only matching.
// Notes: Identical duplicate values are allowed; conflicting, negative, non-numeric and non-finite values remain explicit issues.
shared #"StaffList WorkerContracts" = let
    NormalizedEmployee = Table.TransformColumns(#"IMPORT StaffList Worker's Report",
        {{"Employee Code", StaffListText, type nullable text}}),
    ParsedContract = Table.AddColumn(NormalizedEmployee, "ContractParse", each
        let
            RawAttempt = try [Contracted FN Hours],
            RawValue = if RawAttempt[HasError] then null else RawAttempt[Value],
            IsBlank = not RawAttempt[HasError] and StaffListText(RawValue) = null,
            Attempt = if RawAttempt[HasError] or IsBlank then null else try Number.From(RawValue),
            ParseError = RawAttempt[HasError] or (not IsBlank and Attempt[HasError]),
            ContractValue = if ParseError or IsBlank then null else Attempt[Value]
        in [ContractValue = ContractValue, IsBlank = IsBlank, ParseError = ParseError],
        type [ContractValue = nullable number, IsBlank = logical, ParseError = logical]),
    ExpandedContract = Table.ExpandRecordColumn(ParsedContract, "ContractParse", {"ContractValue", "IsBlank", "ParseError"}),
    IdentifiedRows = Table.SelectRows(ExpandedContract, each [Employee Code] <> null),
    RenamedEmployee = Table.RenameColumns(IdentifiedRows, {{"Employee Code", "EmployeeID"}}),
    Grouped = Table.Group(RenamedEmployee, {"EmployeeID"}, {
        {"ContractValues", each List.Distinct(List.RemoveNulls([ContractValue])), type list},
        {"BlankRows", each List.Count(List.Select([IsBlank], each _ = true)), Int64.Type},
        {"ParseErrors", each List.Count(List.Select([ParseError], each _ = true)), Int64.Type}
    }),
    Resolved = Table.AddColumn(Grouped, "ContractResolution", each
        let
            Values = [ContractValues],
            HasNonFinite = List.AnyTrue(List.Transform(Values, each Number.IsNaN(_) or Number.Abs(_) = #infinity)),
            HasNegative = List.AnyTrue(List.Transform(Values, each _ < 0)),
            HasConflict = List.Count(Values) > 1 or ([BlankRows] > 0 and List.Count(Values) > 0),
            Issue = if [ParseErrors] > 0 then "Non-numeric Contracted FN Hours"
                else if HasNonFinite then "Non-finite Contracted FN Hours"
                else if HasNegative then "Negative Contracted FN Hours"
                else if HasConflict then "Conflicting Contracted FN Hours"
                else if List.Count(Values) = 0 then "Blank Contracted FN Hours"
                else null,
            ContractHours = if List.Count(Values) = 1 and Issue = null then Values{0} else null
        in [Contracted FN Hours = ContractHours, Worker Contract Issue = Issue],
        type [Contracted FN Hours = nullable number, Worker Contract Issue = nullable text]),
    ExpandedResolution = Table.ExpandRecordColumn(Resolved, "ContractResolution", {"Contracted FN Hours", "Worker Contract Issue"})
in Table.Buffer(Table.SelectColumns(ExpandedResolution, {"EmployeeID", "Contracted FN Hours", "Worker Contract Issue"}));

// Query: IMPORT Allocation Extracted
// Purpose: Import the Unit1 allocation workbook once for its allocated staff list.
shared #"IMPORT Allocation Extracted" = let
    WorkbookBinary = Binary.Buffer(File.Contents(#"FilePath-1Input" & "\1-AllocationExtracted.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in Table.Buffer(WorkbookNavigation);

// Query: IMPORT Table_AllocatedStaffList
// Purpose: Extract the allocated staff list from the consolidated allocation workbook import.
shared #"IMPORT Table_AllocatedStaffList" = let
    Source = #"IMPORT Allocation Extracted",
    AllocatedStaffList_Table = Source{[Item="AllocatedStaffList",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(AllocatedStaffList_Table,{{"Name", type text}, {"Role", type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Name", type text}, {"Role", type text}})
in
    #"Changed Type";

// Query: EXTRACT Allocation Staff With Code
// Purpose: Retain allocation employee codes so allocation-only Resources can be checked against Worker's Report.
// Output: Distinct allocation employee/role/name rows; employee codes remain text identifiers.
shared #"EXTRACT Allocation Staff With Code" = let
    Matches = Table.SelectRows(#"IMPORT Allocation Extracted", each [Item] = "AllocationExtracted" and [Kind] = "Table"),
    AllocationTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Allocation staff import", "Expected exactly one AllocationExtracted table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Code", "Name", "Role"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(AllocationTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(AllocationTable, RequiredColumns)
        else error Error.Record("Allocation staff import", "AllocationExtracted is missing required employee columns.", [MissingColumns = MissingColumns]),
    Normalized = Table.TransformColumns(Selected, {
        {"Code", StaffListText, type nullable text},
        {"Name", StaffListText, type nullable text},
        {"Role", StaffListText, type nullable text}
    }),
    Identified = Table.SelectRows(Normalized, each [Code] <> null and [Name] <> null and [Role] <> null),
    Renamed = Table.RenameColumns(Identified, {{"Code", "EmployeeID"}})
in Table.Buffer(Table.Distinct(Renamed));

// Query: IMPORT Capacity Shift Availability
// Purpose: Import the Capacity-ShiftAvailability workbook once for the staff-list extraction.
shared #"IMPORT Capacity Shift Availability" = let
    WorkbookBinary = Binary.Buffer(File.Contents(#"FilePath-2Calculations" & "\Capacity-ShiftAvailability.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in Table.Buffer(WorkbookNavigation);

// Query: EXTRACT Availability Staff List
// Purpose: Select and validate the employee-grain Availability-StaffList output produced in Step 1.
// Output: One preferred-role row per reconciled employee and facility.
shared #"EXTRACT Availability Staff List" = let
    Matches = Table.SelectRows(#"IMPORT Capacity Shift Availability", each
        [Item] = "Availability-StaffList" and [Kind] = "Sheet"),
    RawSheet = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Availability staff import", "Expected exactly one Availability-StaffList sheet.", [Matches = Table.RowCount(Matches)]),
    PromotedHeaders = Table.PromoteHeaders(RawSheet, [PromoteAllScalars = true]),
    RequiredColumns = {"EmployeeID", "Facility-Abbrev", "Name", "Role", "PreferredRole", "EmploymentType", "Contracted FN Hours"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(PromotedHeaders)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(PromotedHeaders, RequiredColumns)
        else error Error.Record("Availability staff import", "Availability-StaffList is missing required Step 1 columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected,
        {{"EmployeeID", type text}, {"Facility-Abbrev", type text}, {"Name", type text}, {"Role", type text},
         {"PreferredRole", type text}, {"EmploymentType", type text}, {"Contracted FN Hours", type nullable number}}),
    Normalized = Table.TransformColumns(Typed, List.Transform(
        {"EmployeeID", "Facility-Abbrev", "Name", "Role", "PreferredRole", "EmploymentType"},
        each {_, StaffListText, type nullable text})),
    Identified = Table.SelectRows(Normalized, each [EmployeeID] <> null or [Name] <> null)
in Table.Buffer(Identified);

// Query: IMPORT StaffList Settings
// Purpose: Import Settings Data once for roster dates, shift duration and the Settings availability cap.
shared #"IMPORT StaffList Settings" = let
    WorkbookBinary = Binary.Buffer(File.Contents(#"FilePath-2Calculations" & "\Settings Data.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in Table.Buffer(WorkbookNavigation);

// Query: EXTRACT StaffList ShiftDuration
// Purpose: Validate the standard shift duration in hours from Settings Data.
shared #"EXTRACT StaffList ShiftDuration" = let
    Matches = Table.SelectRows(#"IMPORT StaffList Settings", each
        [Item] = "ShiftDuration" and List.Contains({"Table", "DefinedName"}, [Kind])),
    SettingsTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Shift duration settings", "Expected exactly one ShiftDuration table or defined name.", [Matches = Table.RowCount(Matches)]),
    WithHeaders = if Table.HasColumns(SettingsTable, "ShiftDuration") then SettingsTable
        else Table.PromoteHeaders(SettingsTable, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "ShiftDuration"),
    ShiftDuration = if List.Count(Values) = 1 then Number.From(Values{0})
        else error "ShiftDuration must contain exactly one value.",
    Validated = if ShiftDuration = null or Number.IsNaN(ShiftDuration) or Number.Abs(ShiftDuration) = #infinity or ShiftDuration <= 0 then
        error "ShiftDuration must be one positive finite number in hours." else ShiftDuration
in Validated;

// Query: EXTRACT StaffList RosterDates
// Purpose: Validate and expose the distinct contiguous roster dates from Settings PermutationDimensions.
shared #"EXTRACT StaffList RosterDates" = let
    Matches = Table.SelectRows(#"IMPORT StaffList Settings", each [Item] = "PermutationDimensions" and [Kind] = "Table"),
    SettingsTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Roster settings", "Expected exactly one PermutationDimensions table.", [Matches = Table.RowCount(Matches)]),
    Required = if Table.HasColumns(SettingsTable, "Date") then Table.SelectColumns(SettingsTable, {"Date"})
        else error "PermutationDimensions is missing Date.",
    Typed = Table.TransformColumnTypes(Required, {{"Date", type date}}),
    MissingDates = Table.RowCount(Table.SelectRows(Typed, each [Date] = null)),
    DistinctDates = Table.Sort(Table.Distinct(Table.SelectRows(Typed, each [Date] <> null)), {{"Date", Order.Ascending}}),
    DateCount = Table.RowCount(DistinctDates),
    SpanDays = if DateCount = 0 then 0 else Duration.Days(List.Max(DistinctDates[Date]) - List.Min(DistinctDates[Date])) + 1,
    Validated = if MissingDates > 0 then error "PermutationDimensions contains blank roster dates."
        else if DateCount = 0 then error "PermutationDimensions contains no roster dates."
        else if DateCount <> SpanDays then error "Settings roster dates must be contiguous."
        else if Number.Mod(DateCount, 14) <> 0 then error "Settings roster date count must be a whole number of fortnights."
        else DistinctDates
in Table.Buffer(Validated);

// Query: EXTRACT StaffList MaxAvailability
// Purpose: Validate the Settings maximum number of shifts available to one resource for the roster.
shared #"EXTRACT StaffList MaxAvailability" = let
    Matches = Table.SelectRows(#"IMPORT StaffList Settings", each
        [Item] = "MaxAvailability" and List.Contains({"Table", "DefinedName"}, [Kind])),
    SettingsTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Maximum availability settings", "Expected exactly one MaxAvailability table or defined name.", [Matches = Table.RowCount(Matches)]),
    WithHeaders = if Table.HasColumns(SettingsTable, "MaxAvailability") then SettingsTable
        else Table.PromoteHeaders(SettingsTable, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "MaxAvailability"),
    Maximum = if List.Count(Values) = 1 then Number.From(Values{0})
        else error "MaxAvailability must contain exactly one value.",
    Validated = if Maximum = null or Number.IsNaN(Maximum) or Number.Abs(Maximum) = #infinity or Maximum <= 0 or Maximum <> Number.RoundDown(Maximum) then
        error "MaxAvailability must be one positive whole number." else Int64.From(Maximum)
in Validated;

shared #"Misaligned Join" = let
    Source = Table.NestedJoin(AllocationTable_StaffList, {"EmployeeID"}, AvailableStaffList, {"EmployeeID"}, "Availability-StaffList", JoinKind.FullOuter),
    #"Expanded Availability-StaffList" = Table.ExpandTableColumn(Source, "Availability-StaffList", {"Name", "Role"}, {"Availability-StaffList.Name", "Availability-StaffList.Role"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Availability-StaffList",{{"Name", Order.Ascending}}),
    #"Added Conditional Column" = Table.AddColumn(#"Sorted Rows", "Custom", each if [Name] <> [#"Availability-StaffList.Name"] then "Misaligned" else null),
    #"Sorted Rows1" = Table.Sort(#"Added Conditional Column",{{"Availability-StaffList.Name", Order.Ascending}}),
    #"Filtered Rows1" = Table.SelectRows(#"Sorted Rows1", each ([Custom] = "Misaligned")),
    #"Inserted Merged Column" = Table.AddColumn(#"Filtered Rows1", "Misaligned-Name", each Text.Combine({[Name], [#"Availability-StaffList.Name"]}, ""), type text),
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Misaligned-Role", each Text.Combine({[Role], [#"Availability-StaffList.Role"]}, ""), type text),
    #"Added Conditional Column1" = Table.AddColumn(#"Inserted Merged Column1", "Misalignment", each if [Name] = null then "Available Only" else if [#"Availability-StaffList.Name"] = null then "Allocated Only" else "ERROR"),
    #"Sorted Rows2" = Table.Sort(#"Added Conditional Column1",{{"Name", Order.Ascending}, {"Availability-StaffList.Name", Order.Ascending}})
in
    #"Sorted Rows2";

shared #"Misaligned Allocation Availability Names" = let
    Source = #"Misaligned Join",
    #"Merged Queries" = Table.NestedJoin(Source, {"Misaligned-Name", "Misaligned-Role"}, Masterlist, {"Name", "Role"}, "Masterlist", JoinKind.LeftOuter),
    #"Expanded Masterlist" = Table.ExpandTableColumn(#"Merged Queries", "Masterlist", {"Resource"}, {"Resource"}),
    #"Sorted Rows2" = Table.Sort(#"Expanded Masterlist",{{"Misaligned-Role", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows2",{"Custom", "Misaligned-Name", "Misaligned-Role", "Misalignment", "Resource"})
in
    #"Removed Other Columns";

// Query: FilePath-1Input
// Purpose: Derive the Unit1 Input folder from the resolved Calculations folder using a boundary-safe suffix replacement.
shared #"FilePath-1Input" = let
    CalculationsFolder = #"FilePath-2Calculations",
    Suffix = "\2. Calculations",
    UnitFolder = if Text.EndsWith(CalculationsFolder, Suffix, Comparer.OrdinalIgnoreCase) then
        Text.Start(CalculationsFolder, Text.Length(CalculationsFolder) - Text.Length(Suffix))
        else error "Resolved StaffListMaster folder is not the expected Unit1 2. Calculations folder.",
    Value = UnitFolder & "\1. Input"
in
    Value;