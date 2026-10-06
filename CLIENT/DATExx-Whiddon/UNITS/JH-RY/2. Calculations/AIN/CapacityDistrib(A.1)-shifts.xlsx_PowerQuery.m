// Power Query from: CapacityDistrib(A.1)-shifts.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\AIN\CapacityDistrib(A.1)-shifts.xlsx
// Extracted: 2026-10-05T07:47:05.241Z

section Section1;

// Query: IMPORT Effort Management Parameter Lists
// Purpose: Read the existing WhiddonCENTRI site's list navigation without business transformations.
// Output: The connector's complete list navigation, including each list's Items table.
shared #"IMPORT Effort Management Parameter Lists" = let
    ListNavigation = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI",
        [Implementation = "2.0", ViewMode = "All"]),
    BufferedNavigation = Table.Buffer(ListNavigation)
in
    BufferedNavigation;

// Query: WorkingDayAllocationThreshold_Status
// Purpose: Read one numeric working-day threshold and report missing, duplicate or invalid settings without stopping diagnostics.
// Notes: Failed reads return null; no default threshold is substituted.
shared WorkingDayAllocationThreshold_Status = let
    SettingAttempt = try
        let
            Lists = #"IMPORT Effort Management Parameter Lists",
            ListRows = Table.SelectRows(Lists, each [Title] = "Effort Management Parameters"),
            ListCount = Table.RowCount(ListRows),
            ParameterList = if ListCount = 1 then ListRows{0}[Items]
                else error Error.Record("A1.ParameterValidation", "Expected one Effort Management Parameters list; found " & Text.From(ListCount) & ".", null),
            ParameterRows = Table.SelectRows(ParameterList, each [Parameter] = "WorkingDayAllocationThreshold"),
            ParameterCount = Table.RowCount(ParameterRows),
            SettingValue = if ParameterCount = 1 then ParameterRows{0}[Value]
                else error Error.Record("A1.ParameterValidation", "Expected one WorkingDayAllocationThreshold row; found " & Text.From(ParameterCount) & ".",
                    null),
            IsFiniteNumber = if not Value.Is(SettingValue, type number) then false
                else not Number.IsNaN(SettingValue) and Number.Abs(SettingValue) <> #infinity,
            ValidatedValue = if IsFiniteNumber then SettingValue
                else error Error.Record("A1.ParameterValidation", "WorkingDayAllocationThreshold must contain a finite numeric Value.", null)
        in
            ValidatedValue,
    Failed = SettingAttempt[HasError],
    Output = [
        Status = if not Failed then "Pass" else if SettingAttempt[Error][Reason] = "A1.ParameterValidation" then "Fail" else "Error",
        Value = if Failed then null else SettingAttempt[Value],
        Details = if Failed then SettingAttempt[Error][Message] else null
    ]
in
    Output;

// Query: WorkingDayAllocationThreshold
// Purpose: Expose the validated SharePoint number; null represents an unavailable setting, never a fallback threshold.
shared WorkingDayAllocationThreshold = WorkingDayAllocationThreshold_Status[Value];

// Query: IMPORT Demand
// Purpose: Open Demand.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT Demand" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Demand.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: Demand Source Prepare
// Purpose: Select and type the role-specific shift-demand table used by A.1.
shared #"Demand Source Prepare" = let
    Navigation = #"IMPORT Demand",
    Matches = Table.SelectRows(Navigation, each [Item] = "ShiftDemandHCAverageANACC" and [Kind] = "Table"),
    DemandTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 demand import", "Expected exactly one ShiftDemandHCAverageANACC table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Facility", "Role", "Date", "ShiftPeriod", "UnitShiftEffort", "ShiftDurations.Duration", "ShiftDemandHCAverage"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(DemandTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then DemandTable
        else error Error.Record("A.1 demand import", "ShiftDemandHCAverageANACC is missing required columns.", [MissingColumns = MissingColumns]),
    FilteredRole = Table.SelectRows(ValidatedTable, each [Role] = Role),
    RenamedShiftDuration = Table.RenameColumns(FilteredRole, {{"ShiftDurations.Duration", "ShiftDurations.ShiftDuration"}}),
    Typed = Table.TransformColumnTypes(RenamedShiftDuration, {{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"ShiftPeriod", type text}, {"UnitShiftEffort", type number}, {"ShiftDurations.ShiftDuration", type number}, {"ShiftDemandHCAverage", type number}})
in
    Typed;

// Query: IMPORT ShiftUnitDemandHRS !!
// Purpose: Preserve the existing prepared-demand interface for downstream queries.
shared #"IMPORT ShiftUnitDemandHRS !!" = #"Demand Source Prepare";

// Query: IMPORT StaffListMaster
// Purpose: Open StaffListMaster.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT StaffListMaster" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\StaffListMaster.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: Masterlist Prepare
// Purpose: Validate and prepare the role-specific employee contract used by A.1.
// Output: One row per source master-list record before Resource-grain validation.
shared #"Masterlist Prepare" = let
    Navigation = #"IMPORT StaffListMaster",
    Matches = Table.SelectRows(Navigation, each [Item] = "Table_Masterlist" and [Kind] = "Table"),
    MasterlistTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 master-list import", "Expected exactly one Table_Masterlist table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Name", "Role", "Resource", "EmployeeID", "Facility-Abbrev", "PreferredRole", "Contracted FN Hours",
        "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MasterlistTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(MasterlistTable, RequiredColumns)
        else error Error.Record("A.1 master-list import", "Table_Masterlist is missing required contract columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected,
        {{"Name", type text}, {"Role", type text}, {"Resource", Int64.Type}, {"EmployeeID", type text},
         {"Facility-Abbrev", type text}, {"PreferredRole", type text}, {"Contracted FN Hours", type nullable number},
         {"Contracted Shifts", type nullable number}, {"Settings Max Availability", type nullable number},
         {"Effective Shift Cap", type nullable number}, {"Limit Basis", type text}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT Masterlist !!
// Purpose: Preserve the existing prepared-master-list interface for downstream queries.
shared #"IMPORT Masterlist !!" = #"Masterlist Prepare";

// Query: A1CheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared A1CheckResult = (name as text, evaluate as function) as record =>
    let
        Result = try evaluate()
    in
        [Check = name, Status = if Result[HasError] then "Error" else if Result[Value] = 0 then "Pass" else "Fail",
         Failures = if Result[HasError] then null else Result[Value],
         Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

// Query: IMPORT Capacity ShiftAvailability
// Purpose: Open Capacity-ShiftAvailability.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT Capacity ShiftAvailability" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Capacity-ShiftAvailability.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: ResDayShift Prepare
// Purpose: Validate and type only the availability fields consumed by A.1, then apply the dynamic role scope.
shared #"ResDayShift Prepare" = let
    Navigation = #"IMPORT Capacity ShiftAvailability",
    Matches = Table.SelectRows(Navigation, each [Item] = "ResDayShift" and [Kind] = "Table"),
    ResDayShiftTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 availability import", "Expected exactly one ResDayShift table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Name", "Role", "Week", "Day", "Shift", "EffectiveShiftHrs", "Date"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(ResDayShiftTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then ResDayShiftTable
        else error Error.Record("A.1 availability import", "ResDayShift is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"Name", type text}, {"Role", type text}, {"Week", Int64.Type}, {"Day", type text}, {"Shift", type text}, {"EffectiveShiftHrs", type number}, {"Date", type date}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT ResDayShift !!
// Purpose: Preserve the existing prepared-availability interface for downstream queries.
shared #"IMPORT ResDayShift !!" = #"ResDayShift Prepare";

// Query: IMPORT AllocationByShiftAverage
// Purpose: Open AllocationByShiftAverage.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT AllocationByShiftAverage" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\AllocationByShiftAverage.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: ResourceShiftAllocation Prepare
// Purpose: Validate and prepare the role-specific resource-shift allocations used by A.1.
shared #"ResourceShiftAllocation Prepare" = let
    Navigation = #"IMPORT AllocationByShiftAverage",
    Matches = Table.SelectRows(Navigation, each [Item] = "ResourceShiftAllocation" and [Kind] = "Table"),
    AllocationTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 allocation import", "Expected exactly one ResourceShiftAllocation table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"ShiftDate", "ShiftPeriod", "IntervalAssociatedShift", "Name", "Role", "ResShiftEffort", "ResShiftEffectiveRatio", "ResShiftFTE"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(AllocationTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then AllocationTable
        else error Error.Record("A.1 allocation import", "ResourceShiftAllocation is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"IntervalAssociatedShift", type text}, {"Name", type text}, {"Role", type text}, {"ResShiftEffort", type number}, {"ResShiftEffectiveRatio", type number}, {"ResShiftFTE", type number}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT ResourceShiftAllocation - Role!!
// Purpose: Preserve the existing prepared-allocation interface for downstream queries.
shared #"IMPORT ResourceShiftAllocation - Role!!" = #"ResourceShiftAllocation Prepare";

shared MaxShiftCluster = 5 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

[ Description = "Use to filter out shift creep into adjacent shifts" ]
shared NotShiftThreshold = 0.2 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared #"SET NWDPriority" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("JYopEsAgEAS/Qq1ew+b28aGCpBA5QFFx/D/L4Lp7JgSyxHSmtz7XXRJFDiRa7G6OnKGDqhc2TqCjqhP2FjZhBM4YgEur/bC22nFT/GopFOMP", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [NWDPriority = _t, #"NWD Type" = _t])
in
    Source;

[ Description = "BUFFER" ]
// Query: PeriodShiftDay B
// Purpose: Validate, type and role-scope the Settings calendar, preserving the existing period interface.
// Inputs: EXTRACT PermutationDimensions and the complete configured Role value.
// Output: Date, Day, Shifts, Period and RolesList in the original Settings row order.
shared #"PeriodShiftDay B" = let
    Source = #"EXTRACT PermutationDimensions",
    RequiredColumns = {"Date", "Day", "Shifts", "Period", "RolesList"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Source)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then Source
        else error Error.Record("A.1 Settings import", "PermutationDimensions is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    FilteredRole = Table.SelectRows(Typed, each [RolesList] = Role),
    #"Removed Other Columns" = Table.SelectColumns(FilteredRole,{"Date", "Day", "Shifts", "Period", "RolesList"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

// Query: Resources
// Purpose: Retain only the stable fields required to map role data to Resource grain.
shared Resources = let
    Source = #"IMPORT Masterlist !!",
    Selected = Table.SelectColumns(Source, {"Name", "Role", "Resource"}),
    // Bad identity cells remain visible as missing keys in diagnostics, never arbitrary matches.
    ReadableKeys = Table.ReplaceErrorValues(Selected, {{"Name", null}, {"Role", null}, {"Resource", null}}),
    Buffered = Table.Buffer(ReadableKeys)
in
    Buffered;

// Query: A1MapResourceIdentity
// Purpose: Keep the source row and every identity candidate, assigning a Resource only for one usable match.
// Output: Source columns plus ResourceMatches, MatchCount and nullable Resource.
shared A1MapResourceIdentity = (source as table, resourceMap as table) as table =>
let
    ResourceKeyCounts = Table.Group(resourceMap, {"Resource"}, {{"IdentityCount", each Table.RowCount(_), Int64.Type}}),
    UniqueResourceRows = Table.SelectRows(ResourceKeyCounts, each [Resource] <> null and [IdentityCount] = 1),
    UniqueResourceKeys = List.Buffer(UniqueResourceRows[Resource]),
    JoinedCandidates = Table.NestedJoin(source, {"Name", "Role"}, resourceMap,
        {"Name", "Role"}, "ResourceMatches", JoinKind.LeftOuter),
    CountedCandidates = Table.AddColumn(JoinedCandidates, "MatchCount", each Table.RowCount([ResourceMatches]), Int64.Type),
    AssignedResource = Table.AddColumn(CountedCandidates, "Resource", each
        let
            UsableName = try [Name] <> null and Text.Trim([Name]) <> "" otherwise false,
            UsableRole = try [Role] <> null and Text.Trim([Role]) <> "" otherwise false
        in
            if [MatchCount] = 1 and UsableName and UsableRole
                and List.Contains(UniqueResourceKeys, [ResourceMatches]{0}[Resource])
            then [ResourceMatches]{0}[Resource] else null, type nullable number)
in
    AssignedResource;

// Query: A1AvailabilityIdentity_Prepare
// Purpose: Assess availability identities without expanding ambiguous joins.
shared A1AvailabilityIdentity_Prepare =
    A1MapResourceIdentity(#"IMPORT ResDayShift !!", Resources);

// Query: A1AllocationIdentity_Prepare
// Purpose: Assess allocation identities within the existing date scope and retain unmatched evidence.
shared A1AllocationIdentity_Prepare = let
    Source = #"IMPORT ResourceShiftAllocation - Role!!",
    InDateScope = Table.SelectRows(Source, each [ShiftDate] >= #"EXTRACT Date_From"),
    Mapped = A1MapResourceIdentity(InDateScope, Resources)
in
    Mapped;

// Query: A1AllocationIdentity_EXCEPTIONS
// Purpose: Publish skipped allocation rows with their original quantities, errors and identity candidates.
// Notes: These rows are excluded from allocation calculations, not discarded from the audit evidence.
shared A1AllocationIdentity_EXCEPTIONS = let
    Source = A1AllocationIdentity_Prepare,
    Unresolved = Table.SelectRows(Source, each [MatchCount] <> 1 or [Resource] = null
        or not (try A1AllocationValueIsUsable([ResShiftFTE]) otherwise false)),
    CandidateIdentifiers = Table.AddColumn(Unresolved, "CandidateResources", each
        Text.Combine(List.Transform(List.Distinct(List.RemoveNulls([ResourceMatches][Resource])), Text.From), ", "), type text),
    WithReason = Table.AddColumn(CandidateIdentifiers, "Reason", each
        if [MatchCount] = 0 then "No matching Resource"
        else if [MatchCount] > 1 then "Ambiguous Resource identity"
        else if [Resource] = null then "Matched identity has a missing, duplicated or invalid key"
        else "Allocation amount is missing, invalid or contains an error", type text),
    WithAction = Table.AddColumn(WithReason, "Action", each "Allocation skipped; evidence retained", type text),
    Output = Table.RemoveColumns(WithAction, {"ResourceMatches"})
in
    Output;

// Query: A1AllocationValueIsUsable
// Purpose: Require a readable finite allocation amount before arithmetic; no replacement amount is invented.
shared A1AllocationValueIsUsable = (value as any) as logical =>
    if not Value.Is(value, type number) then false
    else not Number.IsNaN(value) and Number.Abs(value) <> #infinity;

// Query: A1AllocationValueIssues_Prepare
// Purpose: Localise malformed allocation amounts to known Resources while retaining the raw exception evidence.
shared A1AllocationValueIssues_Prepare = let
    Identifiable = Table.SelectRows(A1AllocationIdentity_Prepare, each [MatchCount] = 1 and [Resource] <> null),
    InvalidAmounts = Table.SelectRows(Identifiable, each not (try A1AllocationValueIsUsable([ResShiftFTE]) otherwise false)),
    WithDate = Table.RenameColumns(InvalidAmounts, {{"ShiftDate", "Date"}}),
    Output = A1IssueDetails(WithDate, "Allocation amount", "Allocation amount is missing, invalid or contains an error",
        "Allocation skipped; Resource spare excluded; evidence retained")
in
    Output;

// Query: A1IssueDetails
// Purpose: Give a set of affected rows the common published diagnostic detail interface.
shared A1IssueDetails = (rows as table, check as text, reason as text, action as text) as table =>
let
    SelectedEvidence = Table.SelectColumns(rows, {"Resource", "Role", "Name", "Date", "Period"}, MissingField.UseNull),
    WithStage = Table.AddColumn(SelectedEvidence, "Stage", each "A.1", type text),
    WithCheck = Table.AddColumn(WithStage, "Check", each check, type text),
    WithReason = Table.AddColumn(WithCheck, "Reason", each reason, type text),
    WithAction = Table.AddColumn(WithReason, "Action", each action, type text),
    Output = Table.ReorderColumns(WithAction, {"Stage", "Check", "Resource", "Role", "Name", "Date", "Period", "Reason", "Action"})
in
    Output;

// Query: A1IdentityIssues_Prepare
// Purpose: Identify individual Resources affected by malformed or ambiguous identities, without a whole-run stop.
shared A1IdentityIssues_Prepare = let
    ResourceMap = Resources,
    InvalidMasterKeys = Table.SelectRows(ResourceMap, each [Resource] = null or [Name] = null or Text.Trim([Name]) = ""),
    MasterKeyCounts = Table.Group(ResourceMap, {"Resource"}, {{"Rows", each _, type table}, {"Count", each Table.RowCount(_), Int64.Type}}),
    DuplicateMasterKeys = Table.SelectRows(MasterKeyCounts, each [Resource] <> null and [Count] > 1),
    DuplicateRows = if Table.IsEmpty(DuplicateMasterKeys) then Table.FirstN(ResourceMap, 0)
        else Table.Combine(DuplicateMasterKeys[Rows]),
    AmbiguousAvailability = Table.SelectRows(A1AvailabilityIdentity_Prepare, each [MatchCount] <> 1 or [Resource] = null),
    AvailabilityEvidence = Table.TransformColumns(AmbiguousAvailability, {{"ResourceMatches", each
        if Table.IsEmpty(_) then #table(type table [Resource = nullable number], {{null}}) else _, type table}}),
    AvailabilityCandidates = Table.ExpandTableColumn(Table.RemoveColumns(AvailabilityEvidence, {"Resource"}),
        "ResourceMatches", {"Resource"}, {"Resource"}),
    AmbiguousAllocations = Table.SelectRows(A1AllocationIdentity_Prepare, each [MatchCount] <> 1 or [Resource] = null),
    AllocationDates = Table.RenameColumns(AmbiguousAllocations, {{"ShiftDate", "Date"}}),
    AllocationEvidence = Table.TransformColumns(AllocationDates, {{"ResourceMatches", each
        if Table.IsEmpty(_) then #table(type table [Resource = nullable number], {{null}}) else _, type table}}),
    AllocationCandidates = Table.ExpandTableColumn(Table.RemoveColumns(AllocationEvidence, {"Resource"}),
        "ResourceMatches", {"Resource"}, {"Resource"}),
    Details = Table.Combine({
        A1IssueDetails(InvalidMasterKeys, "Resource identity", "Missing Resource identifier or name", "Resource spare excluded where identifiable"),
        A1IssueDetails(DuplicateRows, "Resource identity", "Duplicate Resource identifier", "Resource spare excluded"),
        A1IssueDetails(AvailabilityCandidates, "Availability identity", "Availability does not have one usable Resource match", "Availability row skipped; candidate Resource spare excluded"),
        A1IssueDetails(AllocationCandidates, "Allocation identity", "Allocation does not have one usable Resource match", "Allocation skipped; candidate Resource spare excluded")
    })
in
    Table.Buffer(Details);

// Query: A1ResourceGrid_Prepare
// Purpose: Build one calculation-grid row per usable Resource while retaining conflicting names in diagnostics.
shared A1ResourceGrid_Prepare = let
    UsableKeys = Table.SelectRows(Resources, each [Resource] <> null),
    ResourceRows = Table.Group(UsableKeys, {"Role", "Resource"},
        {{"Name", each Text.Combine(List.Sort(List.Distinct(List.RemoveNulls([Name]))), "; "), type text}})
in
    ResourceRows;

// Query: A1AvailabilityPeriods_Prepare
// Purpose: Map uniquely identified availability to the existing Settings periods before spare exclusions.
// Notes: This ungated preparation is also used by allocation-code diagnostics to avoid a dependency cycle.
shared A1AvailabilityPeriods_Prepare = let
    UniqueIdentities = Table.SelectRows(A1AvailabilityIdentity_Prepare, each [MatchCount] = 1 and [Resource] <> null),
    Selected = Table.SelectColumns(UniqueIdentities, {"Name", "Role", "Week", "Shift", "Date", "Resource"}),
    JoinedPeriods = Table.NestedJoin(Selected, {"Date", "Shift", "Role"}, #"PeriodShiftDay B",
        {"Date", "Shifts", "RolesList"}, "PeriodLookup", JoinKind.LeftOuter),
    ExpandedPeriods = Table.ExpandTableColumn(JoinedPeriods, "PeriodLookup", {"Period"}, {"Period"}),
    AddedAvailability = Table.AddColumn(ExpandedPeriods, "Availability", each 1, Int64.Type)
in
    Table.Buffer(AddedAvailability);

// Query: A1AllocationEvidence_Prepare
// Purpose: Preserve any positive mapped allocation as evidence, including values below the existing 0.2/0.7 thresholds.
// Output: One evidence flag per Role, Resource and Period; it does not change allocation quantities.
shared A1AllocationEvidence_Prepare = let
    UniqueIdentities = Table.SelectRows(A1AllocationIdentity_Prepare, each [MatchCount] = 1 and [Resource] <> null),
    PositiveWork = Table.SelectRows(UniqueIdentities, each
        (try A1AllocationValueIsUsable([ResShiftFTE]) and [ResShiftFTE] > 0 otherwise false)),
    JoinedPeriods = Table.NestedJoin(PositiveWork, {"ShiftDate", "IntervalAssociatedShift", "Role"}, #"PeriodShiftDay B",
        {"Date", "Shifts", "RolesList"}, "PeriodLookup", JoinKind.LeftOuter),
    ExpandedPeriods = Table.ExpandTableColumn(JoinedPeriods, "PeriodLookup", {"Period"}, {"Period"}),
    MatchedPeriods = Table.SelectRows(ExpandedPeriods, each [Period] <> null),
    Keys = Table.Distinct(Table.SelectColumns(MatchedPeriods, {"Role", "Resource", "Period"})),
    EvidenceFlag = Table.AddColumn(Keys, "HasAllocationEvidence", each true, type logical)
in
    Table.Buffer(EvidenceFlag);

// Query: A1ResourceIdentity_CHECK
// Purpose: Publish truthful identity checks for manual review without stopping unaffected Resources.
// Notes: Unresolved allocation rows remain in A1AllocationIdentity_EXCEPTIONS.
shared A1ResourceIdentity_CHECK = let
    ResourceMap = Resources,
    AvailabilityMatchCounts = Table.Distinct(Table.SelectColumns(A1AvailabilityIdentity_Prepare, {"Name", "Role", "MatchCount", "Resource"})),
    AllocationMatchCounts = Table.Distinct(Table.SelectColumns(A1AllocationIdentity_Prepare, {"Name", "Role", "MatchCount", "Resource"})),
    AllocationMappingFailureCount = Table.RowCount(Table.SelectRows(AllocationMatchCounts, each [MatchCount] <> 1 or [Resource] = null)),
    Checks = {
        A1CheckResult("Resource identifiers are populated", () => Table.RowCount(Table.SelectRows(ResourceMap, each [Resource] = null))),
        A1CheckResult("Resource identifiers are unique", () => Table.RowCount(ResourceMap) - Table.RowCount(Table.Distinct(ResourceMap, {"Resource"}))),
        A1CheckResult("Resource names are populated", () => Table.RowCount(Table.SelectRows(ResourceMap, each [Name] = null or Text.Trim([Name]) = ""))),
        A1CheckResult("Resource names are unique within role", () => Table.RowCount(ResourceMap) - Table.RowCount(Table.Distinct(ResourceMap, {"Name", "Role"}))),
        A1CheckResult("Availability names map to exactly one Resource", () => Table.RowCount(Table.SelectRows(AvailabilityMatchCounts, each [MatchCount] <> 1 or [Resource] = null))),
        [Check = "Allocation names map to exactly one Resource", Status = if AllocationMappingFailureCount = 0 then "Pass" else "Fail",
         Failures = AllocationMappingFailureCount,
         Details = if AllocationMappingFailureCount = 0 then null else "Unresolved allocations are skipped and retained in A1AllocationIdentity_EXCEPTIONS; other Resources continue."]
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Buffered = Table.Buffer(CheckTable)
in
    Buffered;

// Query: A1ResourceContract
// Purpose: Expose the role-filtered master-list contract at one row per Resource for the A.1 cap join.
shared A1ResourceContract = let
    Source = #"IMPORT Masterlist !!",
    Selected = Table.SelectColumns(Source,
        {"Resource", "Role", "EmployeeID", "Facility-Abbrev", "PreferredRole", "Contracted FN Hours",
         "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"})
in Table.Buffer(Selected);

// Query: A1ContractIssueReason
// Purpose: Assess one Resource's contract without letting another Resource's error stop the calculation.
shared A1ContractIssueReason = (contracts as table, expectedRole as text) as nullable text =>
let
    MatchCount = Table.RowCount(contracts),
    Reason = if MatchCount = 0 then "Missing Resource contract"
        else if MatchCount > 1 then "Duplicate Resource contracts"
        else
            let
                Contract = contracts{0},
                CapAttempt = try Contract[Effective Shift Cap],
                Cap = if CapAttempt[HasError] then null else CapAttempt[Value],
                IsNumeric = Value.Is(Cap, type number),
                IsUsableCap = if not IsNumeric then false else
                    not Number.IsNaN(Cap) and Number.Abs(Cap) <> #infinity and Cap > 0 and Cap = Number.RoundDown(Cap),
                ContractRole = try Contract[Role] otherwise null,
                PreferredRole = try Contract[PreferredRole] otherwise null,
                Employee = try Text.Trim(Contract[EmployeeID]) otherwise null,
                LimitBasis = try Text.Trim(Contract[Limit Basis]) otherwise null,
                Reasons = List.RemoveNulls({
                    if IsUsableCap then null else "Invalid effective shift cap",
                    if ContractRole = expectedRole and PreferredRole = expectedRole then null else "Preferred role does not match contract",
                    if Employee = null or Employee = "" then "Missing Employee identifier" else null,
                    if LimitBasis = null or LimitBasis = "" then "Missing limit basis" else null
                })
            in
                if List.IsEmpty(Reasons) then null else Text.Combine(Reasons, "; ")
in
    Reason;

// Query: A1ResourceContract_Usable
// Purpose: Supply only unique, readable contracts to optional spare arithmetic; failures remain diagnostic evidence.
shared A1ResourceContract_Usable = let
    ReadableKeys = Table.ReplaceErrorValues(A1ResourceContract, {{"Resource", null}}),
    Identifiable = Table.SelectRows(ReadableKeys, each [Resource] <> null),
    Grouped = Table.Group(Identifiable, {"Resource"}, {{"Rows", each _, type table}}),
    WithReason = Table.AddColumn(Grouped, "Issue", each A1ContractIssueReason([Rows], Role), type nullable text),
    Usable = Table.SelectRows(WithReason, each [Issue] = null),
    Output = if Table.IsEmpty(Usable) then Table.FirstN(ReadableKeys, 0) else Table.Combine(Usable[Rows])
in
    Table.Buffer(Output);

// Query: A1ContractIssues_Prepare
// Purpose: Identify affected Resources from contracts independently of prepared capacity and its checks.
shared A1ContractIssues_Prepare = let
    ContractRows = Table.ReplaceErrorValues(A1ResourceContract, {{"Resource", null}}),
    JoinedContracts = Table.NestedJoin(A1ResourceGrid_Prepare, {"Resource"}, ContractRows,
        {"Resource"}, "ContractRows", JoinKind.LeftOuter),
    WithReason = Table.AddColumn(JoinedContracts, "ContractIssue", each A1ContractIssueReason([ContractRows], Role), type nullable text),
    Affected = Table.SelectRows(WithReason, each [ContractIssue] <> null),
    WithStage = Table.AddColumn(Affected, "Stage", each "A.1", type text),
    WithCheck = Table.AddColumn(WithStage, "Check", each "Resource contract", type text),
    WithAction = Table.AddColumn(WithCheck, "Action", each "Resource spare excluded; allocation evidence retained", type text),
    RenamedReason = Table.RenameColumns(WithAction, {{"ContractIssue", "Reason"}}),
    Output = Table.SelectColumns(RenamedReason, {"Stage", "Check", "Resource", "Role", "Name", "Date", "Period", "Reason", "Action"}, MissingField.UseNull)
in
    Table.Buffer(Output);

// Query: A1AvailabilityPeriodIssues_Prepare
// Purpose: Flag missing or duplicate availability-period keys without multiplying the published grid.
shared A1AvailabilityPeriodIssues_Prepare = let
    Source = A1AvailabilityPeriods_Prepare,
    MissingPeriod = Table.SelectRows(Source, each [Period] = null),
    KeyCounts = Table.Group(Source, {"Resource", "Period"}, {{"Rows", each _, type table}, {"Count", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.SelectRows(KeyCounts, each [Period] <> null and [Count] > 1),
    DuplicateRows = if Table.IsEmpty(DuplicateKeys) then Table.FirstN(Source, 0) else Table.Combine(DuplicateKeys[Rows]),
    Details = Table.Combine({
        A1IssueDetails(MissingPeriod, "Availability period", "Availability date/shift has no Settings period", "Resource spare excluded"),
        A1IssueDetails(DuplicateRows, "Availability period", "Availability Resource/Period key is duplicated", "Resource spare excluded")
    })
in
    Details;

// Query: A1ResourceSpareIssues_Prepare
// Purpose: Combine input-scoped issues that exclude Resource spare without depending on published capacity.
// Notes: A shared parameter issue affects every dependent Resource; an unmatched allocation does not.
shared A1ResourceSpareIssues_Prepare = let
    ParameterIssues = if WorkingDayAllocationThreshold_Status[Status] = "Pass" then
            Table.FirstN(A1IdentityIssues_Prepare, 0)
        else A1IssueDetails(A1ResourceGrid_Prepare, "Working-day threshold", WorkingDayAllocationThreshold_Status[Details],
            "Resource spare not evaluated; allocation evidence retained"),
    Details = Table.Combine({A1IdentityIssues_Prepare, A1AllocationValueIssues_Prepare, A1ContractIssues_Prepare,
        A1AvailabilityPeriodIssues_Prepare, A1ClusterIssues_Prepare, ParameterIssues}),
    Identifiable = Table.SelectRows(Details, each [Resource] <> null),
    ResourceReasons = Table.Group(Identifiable, {"Resource"},
        {{"SpareIssueReason", each Text.Combine(List.Sort(List.Distinct([Reason])), "; "), type text}})
in
    Table.Buffer(ResourceReasons);

// Query: A1ResourceSpareStatus_Prepare
// Purpose: Expose one explicit spare-calculation decision per Resource for the existing workbook handoffs.
shared A1ResourceSpareStatus_Prepare = let
    JoinedIssues = Table.NestedJoin(A1ResourceGrid_Prepare, {"Resource"}, A1ResourceSpareIssues_Prepare,
        {"Resource"}, "SpareIssues", JoinKind.LeftOuter),
    ExpandedReason = Table.ExpandTableColumn(JoinedIssues, "SpareIssues", {"SpareIssueReason"}, {"SpareIssueReason"}),
    WithDecision = Table.AddColumn(ExpandedReason, "SpareCalculationAllowed", each [SpareIssueReason] = null, type logical)
in
    WithDecision;

// Query: A1ResourcePeriods_Prepare
// Purpose: Build the ungated Resource-period skeleton used to calculate and diagnose allocation clusters.
shared A1ResourcePeriods_Prepare = let
    Source = A1ResourceGrid_Prepare,
    #"Added Custom" = Table.AddColumn(Source, "Periods", each #"PeriodShiftDay B"),
    // Settings exposes Shifts; downstream A.1 calculations use the singular Shift interface.
    #"Expanded Periods" = Table.ExpandTableColumn(#"Added Custom", "Periods", {"Period", "Shifts", "Day"}, {"Period", "Shift", "Day"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Periods", each ([Period] <> "W")),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Resource", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";

// Query: ResourcePeriodTABLE-empty
// Purpose: Carry Resource spare decisions and allocation evidence through the existing A.2 skeleton interface.
shared #"ResourcePeriodTABLE-empty" = let
    JoinedStatus = Table.NestedJoin(A1ResourcePeriods_Prepare, {"Resource"}, A1ResourceSpareStatus_Prepare,
        {"Resource"}, "SpareStatus", JoinKind.LeftOuter),
    ExpandedStatus = Table.ExpandTableColumn(JoinedStatus, "SpareStatus",
        {"SpareCalculationAllowed", "SpareIssueReason"}, {"SpareCalculationAllowed", "SpareIssueReason"}),
    JoinedEvidence = Table.NestedJoin(ExpandedStatus, {"Role", "Resource", "Period"}, A1AllocationEvidence_Prepare,
        {"Role", "Resource", "Period"}, "AllocationEvidence", JoinKind.LeftOuter),
    ExpandedEvidence = Table.ExpandTableColumn(JoinedEvidence, "AllocationEvidence", {"HasAllocationEvidence"}, {"HasAllocationEvidence"}),
    CompletedEvidence = Table.ReplaceValue(ExpandedEvidence, null, false, Replacer.ReplaceValue, {"HasAllocationEvidence"}),
    AuditAvailability = Table.Group(A1AvailabilityPeriods_Prepare, {"Role", "Resource", "Period"},
        {{"OriginalAvailability", each List.Sum([Availability]), type number}}),
    JoinedAudit = Table.NestedJoin(CompletedEvidence, {"Role", "Resource", "Period"}, AuditAvailability,
        {"Role", "Resource", "Period"}, "AuditAvailability", JoinKind.LeftOuter),
    Output = Table.ExpandTableColumn(JoinedAudit, "AuditAvailability", {"OriginalAvailability"}, {"OriginalAvailability"})
in
    Output;

[ Description = "BUFFER" ]
// Query: ResPeriodAvailabilityTABLE !!
// Purpose: Publish mapped availability with Resource spare exclusions and an unchanged audit value.
shared #"ResPeriodAvailabilityTABLE !!" = let
    Source = Table.SelectRows(A1AvailabilityPeriods_Prepare, each [Period] <> null),
    // Duplicate input cells are diagnosed; the calculation handoff must still have one row per key.
    CellAvailability = Table.Group(Source, {"Role", "Resource", "Period"},
        {{"OriginalAvailability", each List.Sum([Availability]), type number},
         {"AvailabilityBeforeExclusion", each List.Max([Availability]), type number}}),
    JoinedStatus = Table.NestedJoin(CellAvailability, {"Resource"}, A1ResourceSpareStatus_Prepare,
        {"Resource"}, "SpareStatus", JoinKind.LeftOuter),
    ExpandedStatus = Table.ExpandTableColumn(JoinedStatus, "SpareStatus",
        {"SpareCalculationAllowed", "SpareIssueReason"}, {"SpareCalculationAllowed", "SpareIssueReason"}),
    JoinedEvidence = Table.NestedJoin(ExpandedStatus, {"Role", "Resource", "Period"}, A1AllocationEvidence_Prepare,
        {"Role", "Resource", "Period"}, "AllocationEvidence", JoinKind.LeftOuter),
    ExpandedEvidence = Table.ExpandTableColumn(JoinedEvidence, "AllocationEvidence", {"HasAllocationEvidence"}, {"HasAllocationEvidence"}),
    CompletedEvidence = Table.ReplaceValue(ExpandedEvidence, null, false, Replacer.ReplaceValue, {"HasAllocationEvidence"}),
    PublishedAvailability = Table.AddColumn(CompletedEvidence, "Availability", each
        if [SpareCalculationAllowed] = true or [HasAllocationEvidence] = true then [AvailabilityBeforeExclusion] else 0, type number),
    SelectedOutput = Table.RemoveColumns(PublishedAvailability, {"AvailabilityBeforeExclusion"}),
    BUFFER = Table.Buffer(SelectedOutput)
in
    BUFFER;

shared ResDayPeriodAvailabilityTABLE = let
    Source = #"ResPeriodAvailabilityTABLE !!",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period", "Role"}, #"PeriodShiftDay B", {"Period", "RolesList"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Day"}, {"Day"}),
    BUFFER = Table.Buffer(#"Expanded PeriodShiftDay")
in
    BUFFER;

shared ResDayAvailabilityTABLE = let
    Source = ResDayPeriodAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Resource", "Day"}, {{"DayAvailbility", each List.Sum([Availability]), type number}, {"DayShiftCount", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared ResPeriodAvailabilitySUM = let
    Source = List.Sum(#"ResPeriodAvailabilityTABLE !!"[Availability])
in
    Source;

// Query: ResPeriodAllocationCHECK
// Purpose: Calculate only uniquely mapped allocations; skipped rows remain in the exception output.
shared ResPeriodAllocationCHECK = let
    Source = Table.SelectRows(A1AllocationIdentity_Prepare, each [MatchCount] = 1 and [Resource] <> null
        and (try A1AllocationValueIsUsable([ResShiftFTE]) otherwise false)),
    #"Removed Columns" = Table.RemoveColumns(Source,{"ResShiftEffort", "ResShiftEffectiveRatio", "ResourceMatches", "MatchCount"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"ResShiftFTE", "Allocation"}})
in
    #"Renamed Columns";

[ Description = "BUFFER" ]
shared #"ResPeriodAllocation-B" = let
    Source = ResPeriodAllocationCHECK,
    #"Filtered Rows NULL" = Table.SelectRows(Source, each ([Resource] <> null)),
    #"Added SHIFTNUMBER" = Table.AddColumn(#"Filtered Rows NULL", "ShiftNumber", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else if [IntervalAssociatedShift] = "NIGHT" then 3 else null),
    BUFFER = Table.Buffer(#"Added SHIFTNUMBER")
in
    BUFFER;

shared #"ResPeriodAllocation-ShiftBias" = let
    Source = #"ResPeriodAllocation-B",
    #"Grouped Rows1" = Table.Group(Source, {"IntervalAssociatedShift", "Resource", "ShiftDate"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Allocation] > NotShiftThreshold)),
    #"Added SHIFTINDEX" = Table.AddColumn(#"Filtered Rows", "ShiftIndex", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else 3),
    #"Grouped SHIFTBIAS" = Table.Group(#"Added SHIFTINDEX", {"Resource"}, {{"ShiftBias", each List.Average([ShiftIndex]), type number}}),
    BUFFER = Table.Buffer(#"Grouped SHIFTBIAS")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodAllocationTABLE = let
    Source = #"ResPeriodAllocation-B",
    #"Merged Queries1" = Table.NestedJoin(Source, {"ShiftDate", "IntervalAssociatedShift"}, #"PeriodShiftDay B", {"Date", "Shifts"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Grouped ALLOCATION" = Table.Group(#"Expanded PeriodShiftDay", {"ShiftDate", "IntervalAssociatedShift", "Role", "Resource", "Period", "ShiftNumber"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered NOTFULLSHIFT" = Table.SelectRows(#"Grouped ALLOCATION", each ([Allocation] > NotShiftThreshold)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NOTFULLSHIFT",{"Role", "Resource", "Period", "ShiftNumber", "Allocation"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

shared ResDayAllocationTABLE = let
    Source = Table.NestedJoin(ResPeriodAllocationTABLE, {"Period"}, #"PeriodShiftDay B", {"Period"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay1" = Table.ExpandTableColumn(Source, "PeriodShiftDay", {"Day"}, {"Day"}),
    #"Group RESDAYALLOCATION" = Table.Group(#"Expanded PeriodShiftDay1", {"Resource", "Day"}, {{"ResDayAllocation", each List.Sum([Allocation]), type nullable number}, {"ShiftAverage", each List.Average([ShiftNumber]), type number}}),
    BUFFER = Table.Buffer(#"Group RESDAYALLOCATION")
in
    BUFFER;

shared PeriodAllocationTABLE = let
    Source = ResPeriodAllocationTABLE,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Period", "Role"}, {{"A", each List.Sum([Allocation]), type nullable number}})
in
    #"Grouped Rows";

shared #"PeriodA-C' Neg(Surfit)" = let
    Source = #"PeriodC'",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, PeriodAllocationTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodALLOCTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"A"}, {"A"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded PeriodALLOCTABLE",null,0,Replacer.ReplaceValue,{"A"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Replaced Value", "PeriodA-CNeg", each [A] - [#"PeriodC'"], type number),
    #"Reordered A-C'" = Table.ReorderColumns(#"Inserted Subtraction",{"Period", "A", "PeriodC'", "PeriodA-CNeg"}),
    #"Filtered Rows" = Table.SelectRows(#"Reordered A-C'", each ([#"PeriodA-CNeg"] < 0)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared #"PeriodC'" = let
    Source = Table.NestedJoin(#"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, #"ResourcePeriodTABLE-empty", {"Resource", "Period"}, "ResourcePeriodTABLE-empty", JoinKind.LeftOuter),
    #"Expanded DAY" = Table.ExpandTableColumn(Source, "ResourcePeriodTABLE-empty", {"Day"}, {"Day"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded DAY", {"Resource", "Day"}, ResDayNWDTABLE, {"Resource", "Day"}, "ResDayNWDTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayNWDTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayNWDTABLE", {"PotentialAvailability"}, {"PotentialAvailability"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDayNWDTABLE", each ([PotentialAvailability] = null)),
    #"Grouped C'" = Table.Group(#"Filtered Rows", {"Period"}, {{"PeriodC'", each List.Sum([Availability]), type number}})
in
    #"Grouped C'";

// Query: PeriodC'-D
// Purpose: Compare potential capacity with demand and expose a nullable diagnostic capacity-to-demand ratio.
shared #"PeriodC'-D" = let
    Source = #"PeriodC'",
    Custom1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(Custom1, {"Period"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C'-D", each [#"PeriodC'"] - [D], type number),
    // The ratio is diagnostic only; null preserves undefined or invalid demand without propagating an error to A.2.
    #"Inserted Safe Division" = Table.AddColumn(#"Inserted Subtraction", "C'/D", each
        let
            Demand = try Number.From([D]) otherwise null
        in
            if Demand = null or Demand = 0 then null else [#"PeriodC'"] / Demand,
        type nullable number
    )
in
    #"Inserted Safe Division";

shared #"PeriodC'-DPos (ExcessPot)" = let
    Source = #"PeriodC'-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C'-D"] > 0)
in
    #"Filtered Rows";

shared #"ResC'" = let
    Source = #"C' ResSingleShiftDayTABLE",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([Availability]), type number}})
in
    #"Grouped Rows";

// Query: A1ResourceContract_CHECK
// Purpose: Report preferred-role contract validity and coverage without a publication stop.
shared A1ResourceContract_CHECK = let
    Contracts = A1ResourceContract,
    // Inspect raw mapped availability so Resource exclusion cannot hide a contract failure.
    AvailabilityResources = Table.Distinct(Table.SelectColumns(A1AvailabilityPeriods_Prepare, {"Resource"})),
    Coverage = Table.NestedJoin(AvailabilityResources, {"Resource"}, Contracts, {"Resource"}, "Contract", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(Coverage, "ContractMatchCount", each Table.RowCount([Contract]), Int64.Type),
    SingleMatches = Table.SelectRows(WithMatchCount, each [ContractMatchCount] = 1),
    ExpandedMatches = Table.ExpandTableColumn(SingleMatches, "Contract",
        {"Role", "EmployeeID", "PreferredRole", "Effective Shift Cap", "Limit Basis"},
        {"Contract Role", "EmployeeID", "PreferredRole", "Effective Shift Cap", "Limit Basis"}),
    Checks = {
        A1CheckResult("Contract Resources are populated", () => Table.RowCount(Table.SelectRows(Contracts, each [Resource] = null))),
        A1CheckResult("Unique Resource contracts", () => Table.RowCount(Contracts) -
            Table.RowCount(Table.Distinct(Contracts, {"Resource"}))),
        A1CheckResult("Complete availability contract coverage", () =>
            Table.RowCount(Table.SelectRows(WithMatchCount, each [ContractMatchCount] <> 1))),
        A1CheckResult("Valid effective shift caps", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Effective Shift Cap] = null or Number.IsNaN([Effective Shift Cap])
            or Number.Abs([Effective Shift Cap]) = #infinity or [Effective Shift Cap] <= 0
            or [Effective Shift Cap] <> Number.RoundDown([Effective Shift Cap])))),
        A1CheckResult("Preferred role matches contract", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Contract Role] <> Role or [PreferredRole] <> Role))),
        A1CheckResult("Employee identifiers are populated", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [EmployeeID] = null or (try Text.Trim([EmployeeID]) otherwise "") = ""))),
        A1CheckResult("Limit basis is populated", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Limit Basis] = null or (try Text.Trim([Limit Basis]) otherwise "") = "")))
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Buffered = Table.Buffer(CheckTable)
in
    Buffered;

// Query: ResAv-C'
// Purpose: Apply contract caps only to Resources whose spare calculations can be evaluated.
shared #"ResAv-C'" = let
    EligibleResources = Table.SelectRows(A1ResourceSpareStatus_Prepare, each [SpareCalculationAllowed] = true),
    JoinedStatus = Table.NestedJoin(#"ResC'", {"Resource"}, EligibleResources, {"Resource"}, "SpareStatus", JoinKind.Inner),
    Source = Table.RemoveColumns(JoinedStatus, {"SpareStatus"}),
    // Join once at Resource grain; never repeat roster contract totals across Resource-period rows.
    JoinedContract = Table.NestedJoin(Source, {"Resource"}, A1ResourceContract_Usable, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    ExpandedContract = Table.ExpandTableColumn(JoinedContract, "ResourceContract",
        {"Effective Shift Cap", "Limit Basis"}, {"Effective Shift Cap", "Limit Basis"}),
    #"Added AVAILCAP" = Table.AddColumn(ExpandedContract, "RosterAvailabilityCAPPED", each
        if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap] else [RosterAvailability], type number),
    #"Inserted REDUCTION" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted REDUCTION",{{"AvailabilityReduction", "ResAv-C'"}})
in
    #"Renamed Columns";

[ Description = "BUFFER   Tag RP cell with R that need to reduced becuase over allocated" ]
shared #"ResPeriod(ExcessPot)SpareAvailabilityTABLE" = let
    Source = #"C' ResSingleShiftDayTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, ResPeriodWDTABLE, {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodWDTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodWDTABLE", {"PotentialAvailability"}, {"PotentialAvailability"}),
    #"Filtered SPARE-ALLOCATIONNULL" = Table.SelectRows(#"Expanded ResPeriodWDTABLE", each ([Allocation] = null) and ([PotentialAvailability] = null)),
    #"Merged RESAv-C'" = Table.NestedJoin(#"Filtered SPARE-ALLOCATIONNULL", {"Resource"}, #"ResAv-C'", {"Resource"}, "ResRosterAvailability", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailability" = Table.ExpandTableColumn(#"Merged RESAv-C'", "ResRosterAvailability", {"ResAv-C'"}, {"ResAv-C'"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResRosterAvailability", {"Period"}, #"PeriodC'-DPos (ExcessPot)", {"Period"}, "PeriodD-A", JoinKind.LeftOuter),
    #"Expanded SURPLUS-PERIODC'-D.POS" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodD-A", {"C'-D", "C'/D"}, {"C'-D", "C'/D"}),
    #"Merged Queries3" = Table.NestedJoin(#"Expanded SURPLUS-PERIODC'-D.POS", {"Resource", "Period"}, #"MultiDayPeriod-Remove", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded MULTIDAYPERIOD -REMOVE" = Table.ExpandTableColumn(#"Merged Queries3", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    #"Filtered MULTISHIFTDAY-LOWERPRIORITY" = Table.SelectRows(#"Expanded MULTIDAYPERIOD -REMOVE", each ([NoAllocationKeep] <> false       )),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered MULTISHIFTDAY-LOWERPRIORITY",null,0,Replacer.ReplaceValue,{"C'-D"}),
    #"Added UNASSIGNEDAVAIL" = Table.AddColumn(#"Replaced Value", "UnassignedAvail", each if [#"C'-D"] > 0 and [#"ResAv-C'"] >0 
then [Availability]
else 0),
    #"Filtered UNASSIGNEDAVAIL" = Table.SelectRows(#"Added UNASSIGNEDAVAIL", each ([UnassignedAvail] <>0)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered UNASSIGNEDAVAIL",{"Availability", "C'-D", "C'/D", "NoAllocationKeep"}),
    #"Added COMBINE P+R" = Table.AddColumn(#"Removed Columns", "PRCell", each "P" & Text.From([Period]) & "-R" & Text.From([Resource]))
in
    #"Added COMBINE P+R";

[ Description = "BUFFER" ]
shared ResPeriodCapPrioritised = let
    Source = #"ResPeriod(ExcessPot)SpareAvailabilityTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([UnassignedAvail] > 0)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period"}, #"PeriodC'-DPos (ExcessPot)", {"Period"}, "C-DPos (ExcessPot)", JoinKind.LeftOuter),
    #"Expanded zPeriodC-DPos (ExcessPot)" = Table.ExpandTableColumn(#"Merged Queries", "C-DPos (ExcessPot)", {"C'-D", "C'/D"}, {"C'-D", "C'/D"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded zPeriodC-DPos (ExcessPot)", {"Period"}, #"PeriodA-C' Neg(Surfit)", {"Period"}, "PeriodA-C", JoinKind.LeftOuter),
    #"Expanded PeriodA-C" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodA-C", {"PeriodA-CNeg"}, {"PeriodA-CNeg"}),
    #"Filtered A-C C-D" = Table.SelectRows(#"Expanded PeriodA-C", each [#"PeriodA-CNeg"] <=0 and [#"C'-D"] > 0),
    #"Reordered Columns1" = Table.ReorderColumns(#"Filtered A-C C-D",{"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "C'/D"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Resource", "Period"}, ResPeriodShiftNWDTABLE, {"Resource", "Period"}, "ResPeriodAvailPrioritiesTABLE", JoinKind.LeftOuter),
    #"Expanded NWDPRIORITIES" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailPrioritiesTABLE", {"ClusterAllocation", "NWDPriority", "NWDType"}, {"ClusterAllocation", "NWDPriority", "NWDType"})
in
    #"Expanded NWDPRIORITIES";

[ Description = "BUFFER" ]
// Query: ResourcePeriodTABLEIndex
// Purpose: Index the ungated Resource-period grid used by allocation-code diagnostics.
shared ResourcePeriodTABLEIndex = let
    Source = A1ResourcePeriods_Prepare,
    #"Added Index" = Table.AddIndexColumn(Source, "ResIndex", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

[ Description = "BUFFER" ]
// Query: AppliedAllocation+ActiveDays
// Purpose: Combine allocation evidence with ungated availability for independent allocation-code diagnostics.
shared #"AppliedAllocation+ActiveDays" = let
    Source = Table.NestedJoin(ResourcePeriodTABLEIndex, {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(Source, "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, A1AvailabilityPeriods_Prepare, {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Added INVALIDALLOCATION" = Table.AddColumn(#"Expanded ResPeriodAvailabilityTABLE", "InvalidAllocation", each if [Allocation] <> null 
and [Availability] = null 
then true
else null),
    #"Renamed Columns" = Table.RenameColumns(#"Added INVALIDALLOCATION",{{"InvalidAllocation", "InValidAllocation"}})
in
    #"Renamed Columns";

// Query: ResDaysAllocated
// Purpose: Collapse shift-level allocations to daily flags and calculate deterministic five-day allocation codes per Resource.
shared ResDaysAllocated = let
    // 1) Start from your per‑shift allocations
    Raw = #"AppliedAllocation+ActiveDays",

    // 2) Keep only Resource, Day and the raw Allocation
    Base = Table.SelectColumns(Raw,{"Resource", "Day", "Allocation"}),
    Threshold = ValidAllocation,

    // 3) Turn each shift into a 1/0 flag
    ShiftFlagged = Table.TransformColumns(
        Base,
        {{"Allocation", each if _ <> null and Threshold <> null and _ > Threshold then 1 else 0, Int64.Type}}
    ),

    // 4) Collapse to one row per Resource/Day (1 if any shift had allocation)
    Daily = Table.Group(ShiftFlagged, {"Resource", "Day"}, {{"AllocFlag", each List.Max([Allocation]), type number}}),

    // 5) Enforce numeric chronological order before buffering and sequence calculations.
    TypedDay = Table.TransformColumnTypes(Daily, {{"Day", Int64.Type}}),
    LeanDailyBUFFER = Table.Buffer(TypedDay),

    // 6) Sort each nested Resource table explicitly; Table.Group does not guarantee preservation of its input row order.
    Sorted = Table.Sort(LeanDailyBUFFER, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    Grouped = Table.Group(
        Sorted,
        {"Resource"},
        {{"AllData", each Table.AddIndexColumn(
            Table.Sort(_, {{"Day", Order.Ascending}}),
            "Idx", 0, 1, Int64.Type
        ), type table}}
    ),

    // 7) **Note the double‐braces** here—this makes your transformOperations a list of one operation
    WithCodes = Table.TransformColumns(
        Grouped,
        {
            {
                "AllData",
                (tbl) =>
                    let
                        flags    = Table.Column(tbl, "AllocFlag"),
                        n        = List.Count(flags),
                        getD     = (i, o) => if i+o < 0 or i+o >= n then "0" else Text.From(flags{i+o}),
                        enriched = Table.AddColumn(
                                     tbl,
                                     "Code",
                                     each Text.Combine(List.Transform({-2..2}, (off) => getD([Idx], off)), "")
                                   )
                    in
                        enriched
            }
        }
    ),

    // 8) Expand only the new fields (Day, AllocFlag, Idx, Code)
    Expanded = Table.ExpandTableColumn(
        WithCodes,
        "AllData",
        {"Day", "AllocFlag", "Idx", "Code"},
        {"Day", "AllocFlag", "Idx", "Code"}
    ),
    #"Renamed Columns" = Table.RenameColumns(Expanded,{{"Idx", "ResIndex"}})
in
    #"Renamed Columns"
;

[ Description = "BUFFER" ]
// Query: AllocationCode
// Purpose: Identify cluster boundaries from deterministic Resource/day allocation sequences.
shared AllocationCode = let
    // 1) Base daily codes
    Source = ResDaysAllocated,
    // Cluster neighbour offsets must use numeric day order, never lexical text order.
    SourceDaily = Table.TransformColumnTypes(Source, {{"Day", Int64.Type}}),

    // 2) Group by Resource → sort by Day → add a zero‑based index
    Grouped = Table.Group(
      SourceDaily,
      {"Resource"},
      {{"Group",
        each Table.AddIndexColumn(
               Table.Sort(_, {{"Day", Order.Ascending}}),
               "Idx", 0, 1, Int64.Type
             ),
        type table
      }}
    ),

    // 3) Inside each Resource‑group, peek at Resource IDs ±2 days
    WithNeighbors = Table.TransformColumns(
      Grouped,
      {"Group",(tbl) =>
        let
          resList = Table.Column(tbl, "Resource"),
          n       = List.Count(resList),
          getRes  = (i,off) => let p = i+off in if p<0 or p>=n then null else resList{p},
          step1   = Table.AddColumn(tbl, "Yesterday.Resource2", each getRes([Idx], -2), Int64.Type),
          step2   = Table.AddColumn(step1,   "Yesterday.Resource1", each getRes([Idx], -1), Int64.Type),
          step3   = Table.AddColumn(step2,   "Tomorrow.Resource1",  each getRes([Idx],  1), Int64.Type),
          step4   = Table.AddColumn(step3,   "Tomorrow.Resource2",  each getRes([Idx],  2), Int64.Type)
        in
          step4
      }
    ),

    // 4) Flatten back to one table (bringing in the new neighbor columns)
    Expanded = Table.ExpandTableColumn(WithNeighbors, "Group", {"Day", "AllocFlag", "ResIndex", "Code", "Yesterday.Resource2", "Yesterday.Resource1", "Tomorrow.Resource1", "Tomorrow.Resource2"}, {"Day", "AllocFlag", "ResIndex", "Code", "Yesterday.Resource2", "Yesterday.Resource1", "Tomorrow.Resource1", "Tomorrow.Resource2"}),
    SortedExpanded = Table.Sort(Expanded, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),

    // 5) Original “Added RESOURCEEND” logic, unchanged
    AddedResourceEnd = Table.AddColumn(
      SortedExpanded,
      "ResourceStartEnd",
      each 
        if      [Resource] <> [Tomorrow.Resource1]   then "ResourceEnd1"
        else if [Resource] <> [Yesterday.Resource1]  then "ResourceStart1"
        else if [Resource] <> [Tomorrow.Resource2]   then "ResourceEnd2"
        else if [Resource] <> [Yesterday.Resource2]  then "ResourceStart2"
        else null,
      type text
    ),





 //#"Removed Other Columns1" = Table.SelectColumns(#"Added RESOURCEEND",{"Resource", "Day", "Allocation", "ResIndex", "ResDayAllocated", "Tomorrow.Resource1", "Code", "ResourceStartEnd"}),
    // A one-day allocation is both a start and an end, including when it occurs on a Resource boundary.
    #"Added CLUSTERPOINTS" = Table.AddColumn(AddedResourceEnd, "ClusterPoints", each if [Code] = "00100" then "Allocated-SingleStart/End"
        else if (
            //general start
            (  [Code] = "00111" 
                or [Code] ="00101" 
                or [Code] ="00110" 
            )
            //Resouce start
            or (
                (   [ResourceStartEnd] = "ResourceStart1"
                    or
                    [ResourceStartEnd] = "ResourceStart2"
                )
                and (
                    [Code] = "11111"
                    or 
                    [Code] = "10111"

                    )
                )           
           )     
        then 
        "Allocated-ClusterStart"

        else 
            if 
            //General end
            [Code] = "11100" 
            or [Code] ="10100" 
            or [Code] ="01100"
            //Change of resource
            or  (
                [Tomorrow.Resource1] <> [Resource] 
                and [AllocFlag] <> 0
                )
           
        then "Allocated-ClusterEnd" 

        else null),
    BUFFER = Table.Buffer(#"Added CLUSTERPOINTS")
in
    BUFFER;

// Query: ClusterIDStart
// Purpose: Assign stable sequence numbers to every cluster start, including single-day clusters.
shared ClusterIDStart = let
    Source = Table.Sort(#"AllocationCode", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterStart" or [ClusterPoints] = "Allocated-ClusterSINGLE" or [ClusterPoints] = "Allocated-SingleStart/End"),
    #"Added Index1" = Table.AddIndexColumn(#"Filtered Rows", "ClusterIndex", 1, 1, Int64.Type)
in
    #"Added Index1";

// Query: ClusterIDEnd
// Purpose: Pair each Resource's chronological cluster ends with that Resource's start identifiers.
// Notes: Start identifiers remain globally unique. An unpaired end has a null identifier and is diagnosed without shifting other Resources.
shared ClusterIDEnd = let
    Source = Table.Sort(#"AllocationCode", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterEnd" or [ClusterPoints] = "Allocated-ClusterSINGLE" or [ClusterPoints] = "Allocated-SingleStart/End"),
    EndResourceGroups = Table.Group(#"Filtered Rows", {"Resource"}, {{"Ends", each
        let
            OrderedEnds = Table.Sort(_, {{"Day", Order.Ascending}}),
            NumberedResourceEnds = Table.AddIndexColumn(OrderedEnds, "BoundaryOrdinal", 1, 1, Int64.Type)
        in
            NumberedResourceEnds, type table}}),
    NumberedEnds = Table.ExpandTableColumn(EndResourceGroups, "Ends",
        {"Day", "ClusterPoints", "BoundaryOrdinal"}, {"Day", "ClusterPoints", "BoundaryOrdinal"}),
    StartResourceGroups = Table.Group(ClusterIDStart, {"Resource"}, {{"Starts", each
        let
            OrderedStarts = Table.Sort(_, {{"Day", Order.Ascending}, {"ClusterIndex", Order.Ascending}}),
            NumberedResourceStarts = Table.AddIndexColumn(OrderedStarts, "BoundaryOrdinal", 1, 1, Int64.Type)
        in
            NumberedResourceStarts, type table}}),
    NumberedStarts = Table.ExpandTableColumn(StartResourceGroups, "Starts",
        {"BoundaryOrdinal", "ClusterIndex"}, {"BoundaryOrdinal", "ClusterIndex"}),
    JoinedStarts = Table.NestedJoin(NumberedEnds, {"Resource", "BoundaryOrdinal"}, NumberedStarts,
        {"Resource", "BoundaryOrdinal"}, "PairedStart", JoinKind.LeftOuter),
    ExpandedStartIdentifier = Table.ExpandTableColumn(JoinedStarts, "PairedStart", {"ClusterIndex"}, {"ClusterIndex"}),
    PublicEndColumns = Table.SelectColumns(ExpandedStartIdentifier, {"Resource", "Day", "ClusterPoints", "ClusterIndex"}),
    TypedIdentifier = Table.TransformColumnTypes(PublicEndColumns, {{"ClusterIndex", Int64.Type}}),
    SortedEnds = Table.Sort(TypedIdentifier, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}})
in
    SortedEnds;

shared #"ClusterIDStart+End" = let
    Source = Table.Combine({ClusterIDEnd, ClusterIDStart}),
    #"Removed Duplicates" = Table.Distinct(Source)
in
    #"Removed Duplicates";

// Query: A1NumberClusterBoundaries
// Purpose: Give each Resource's existing start or end boundaries a chronological ordinal for validation.
// Notes: These validation ordinals do not replace the calculation's existing ClusterIndex values.
shared A1NumberClusterBoundaries = (boundaries as table) as table =>
let
    BoundaryColumns = Table.SelectColumns(boundaries, {"Resource", "Day", "ClusterIndex"}),
    ResourceGroups = Table.Group(BoundaryColumns, {"Resource"}, {{"Boundaries", each
        let
            OrderedBoundaries = Table.Sort(_, {{"Day", Order.Ascending}, {"ClusterIndex", Order.Ascending}}),
            NumberedResourceBoundaries = Table.AddIndexColumn(OrderedBoundaries, "BoundaryOrdinal", 1, 1, Int64.Type)
        in
            NumberedResourceBoundaries, type table}}),
    NumberedBoundaries = Table.ExpandTableColumn(ResourceGroups, "Boundaries",
        {"Day", "ClusterIndex", "BoundaryOrdinal"}, {"Day", "ClusterIndex", "BoundaryOrdinal"})
in
    NumberedBoundaries;

// Query: A1ClusterBoundaryPairs_Prepare
// Purpose: Pair existing starts and ends by Resource and chronological ordinal without changing cluster assignment.
// Output: One row per Resource/boundary ordinal, including unmatched boundaries and the preceding end day.
shared A1ClusterBoundaryPairs_Prepare = let
    NumberedStarts = A1NumberClusterBoundaries(ClusterIDStart),
    StartBoundaries = Table.RenameColumns(NumberedStarts, {{"Day", "StartDay"}, {"ClusterIndex", "StartClusterIndex"}}),
    NumberedEnds = A1NumberClusterBoundaries(ClusterIDEnd),
    EndBoundaries = Table.RenameColumns(NumberedEnds, {{"Day", "EndDay"}, {"ClusterIndex", "EndClusterIndex"}}),
    JoinedBoundaries = Table.NestedJoin(StartBoundaries, {"Resource", "BoundaryOrdinal"}, EndBoundaries,
        {"Resource", "BoundaryOrdinal"}, "Ends", JoinKind.FullOuter),
    ExpandedEnds = Table.ExpandTableColumn(JoinedBoundaries, "Ends",
        {"Resource", "BoundaryOrdinal", "EndDay", "EndClusterIndex"},
        {"EndResource", "EndBoundaryOrdinal", "EndDay", "EndClusterIndex"}),
    WithResource = Table.AddColumn(ExpandedEnds, "PairedResource", each
        if [Resource] = null then [EndResource] else [Resource], type nullable number),
    WithOrdinal = Table.AddColumn(WithResource, "PairedOrdinal", each
        if [BoundaryOrdinal] = null then [EndBoundaryOrdinal] else [BoundaryOrdinal], Int64.Type),
    RemovedJoinKeys = Table.RemoveColumns(WithOrdinal, {"Resource", "BoundaryOrdinal", "EndResource", "EndBoundaryOrdinal"}),
    RenamedPairKeys = Table.RenameColumns(RemovedJoinKeys, {{"PairedResource", "Resource"}, {"PairedOrdinal", "BoundaryOrdinal"}}),
    ResourceGroups = Table.Group(RenamedPairKeys, {"Resource"}, {{"Pairs", each
        let
            OrderedPairs = Table.Sort(_, {{"BoundaryOrdinal", Order.Ascending}}),
            EndDays = List.Buffer(OrderedPairs[EndDay]),
            WithPreviousEnd = Table.AddColumn(OrderedPairs, "PreviousEndDay", each
                if [BoundaryOrdinal] = 1 then null else EndDays{[BoundaryOrdinal] - 2}, type nullable number)
        in
            WithPreviousEnd, type table}}),
    ExpandedPairs = Table.ExpandTableColumn(ResourceGroups, "Pairs",
        {"BoundaryOrdinal", "StartDay", "EndDay", "StartClusterIndex", "EndClusterIndex", "PreviousEndDay"},
        {"BoundaryOrdinal", "StartDay", "EndDay", "StartClusterIndex", "EndClusterIndex", "PreviousEndDay"}),
    SortedPairs = Table.Sort(ExpandedPairs, {{"Resource", Order.Ascending}, {"BoundaryOrdinal", Order.Ascending}})
in
    Table.Buffer(SortedPairs);

// Query: A1ClusterBoundaryCoverage_Prepare
// Purpose: Count existing cluster intervals covering each qualifying allocated Resource/day.
// Notes: Coverage uses AllocFlag = 1 under the unchanged working-day threshold, never interval length as worked-day count.
shared A1ClusterBoundaryCoverage_Prepare = let
    AllocatedDays = Table.SelectRows(AllocationCode, each [AllocFlag] = 1),
    DailyKeys = Table.SelectColumns(AllocatedDays, {"Resource", "Day"}),
    JoinedPairs = Table.NestedJoin(DailyKeys, {"Resource"}, A1ClusterBoundaryPairs_Prepare,
        {"Resource"}, "BoundaryPairs", JoinKind.LeftOuter),
    WithCoverageCount = Table.AddColumn(JoinedPairs, "BoundaryCoverageCount", each
        let
            AllocatedDay = [Day],
            CoveringPairs = Table.SelectRows([BoundaryPairs], each
                [StartDay] <> null and [EndDay] <> null
                    and [StartDay] <= AllocatedDay and AllocatedDay <= [EndDay])
        in
            Table.RowCount(CoveringPairs), Int64.Type),
    Output = Table.RemoveColumns(WithCoverageCount, {"BoundaryPairs"})
in
    Table.Buffer(Output);

// Query: A1ClusterBoundaryIssues_Prepare
// Purpose: Publish missing, misordered or misidentified boundaries and uncovered qualifying work for Resource exclusions.
// Notes: Single-day intervals and internal one-day gaps retain the existing business definition; no cluster limit is changed here.
shared A1ClusterBoundaryIssues_Prepare = let
    Pairs = A1ClusterBoundaryPairs_Prepare,
    UnpairedDates = Table.SelectRows(Pairs, each [StartDay] = null or [EndDay] = null),
    MisalignedIdentifiers = Table.SelectRows(Pairs, each [StartClusterIndex] = null or [EndClusterIndex] = null
        or [StartClusterIndex] <> [EndClusterIndex]),
    MisorderedIntervals = Table.SelectRows(Pairs, each
        ([StartDay] <> null and [EndDay] <> null and [StartDay] > [EndDay])
            or ([StartDay] <> null and [PreviousEndDay] <> null and [StartDay] <= [PreviousEndDay])),
    AllocatedDays = Table.SelectRows(AllocationCode, each [AllocFlag] = 1),
    AllocatedDayKeys = Table.SelectColumns(AllocatedDays, {"Resource", "Day"}),
    JoinedStartWork = Table.NestedJoin(Pairs, {"Resource", "StartDay"}, AllocatedDayKeys,
        {"Resource", "Day"}, "StartWork", JoinKind.LeftOuter),
    JoinedEndWork = Table.NestedJoin(JoinedStartWork, {"Resource", "EndDay"}, AllocatedDayKeys,
        {"Resource", "Day"}, "EndWork", JoinKind.LeftOuter),
    NonworkingBoundaries = Table.SelectRows(JoinedEndWork, each
        ([StartDay] <> null and Table.IsEmpty([StartWork])) or ([EndDay] <> null and Table.IsEmpty([EndWork]))),
    IncorrectCoverage = Table.SelectRows(A1ClusterBoundaryCoverage_Prepare, each [BoundaryCoverageCount] <> 1),
    Details = Table.Combine({
        A1IssueDetails(UnpairedDates, "Cluster boundary pairing", "Cluster start/end dates are not paired within the Resource", "Resource spare excluded; allocation evidence retained"),
        A1IssueDetails(MisalignedIdentifiers, "Cluster boundary identifiers", "Paired start/end cluster identifiers do not align within the Resource", "Resource spare excluded; allocation evidence retained"),
        A1IssueDetails(MisorderedIntervals, "Cluster boundary order", "Cluster intervals are reversed or overlap within the Resource", "Resource spare excluded; allocation evidence retained"),
        A1IssueDetails(NonworkingBoundaries, "Cluster boundary endpoints", "Cluster boundary is not a qualifying allocated day", "Resource spare excluded; allocation evidence retained"),
        A1IssueDetails(IncorrectCoverage, "Cluster boundary work coverage", "Qualifying allocated day does not have exactly one cluster interval", "Resource spare excluded; allocation evidence retained")
    })
in
    Table.Buffer(Details);

// Query: A1ClusterIssues_Prepare
// Purpose: Identify Resources with cluster-key, boundary-order or allocated-day coverage issues without blocking others.
shared A1ClusterIssues_Prepare = let
    Source = AllocationCode,
    MissingKeys = Table.SelectRows(Source, each [Resource] = null or [ResIndex] = null or [Day] = null),
    LookupCounts = Table.Group(Source, {"Resource", "ResIndex"}, {{"Rows", each _, type table}, {"Count", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.SelectRows(LookupCounts, each [Count] > 1),
    DuplicateRows = if Table.IsEmpty(DuplicateKeys) then Table.FirstN(Source, 0) else Table.Combine(DuplicateKeys[Rows]),
    ResourceKeys = Table.Distinct(Table.SelectColumns(Source, {"Resource"})),
    StartCounts = Table.Group(ClusterIDStart, {"Resource"}, {{"StartCount", each Table.RowCount(_), Int64.Type}}),
    EndCounts = Table.Group(ClusterIDEnd, {"Resource"}, {{"EndCount", each Table.RowCount(_), Int64.Type}}),
    JoinedStarts = Table.NestedJoin(ResourceKeys, {"Resource"}, StartCounts, {"Resource"}, "Starts", JoinKind.LeftOuter),
    ExpandedStarts = Table.ExpandTableColumn(JoinedStarts, "Starts", {"StartCount"}, {"StartCount"}),
    JoinedEnds = Table.NestedJoin(ExpandedStarts, {"Resource"}, EndCounts, {"Resource"}, "Ends", JoinKind.LeftOuter),
    ExpandedEnds = Table.ExpandTableColumn(JoinedEnds, "Ends", {"EndCount"}, {"EndCount"}),
    CompleteCounts = Table.ReplaceValue(ExpandedEnds, null, 0, Replacer.ReplaceValue, {"StartCount", "EndCount"}),
    Unbalanced = Table.SelectRows(CompleteCounts, each [StartCount] <> [EndCount]),
    Details = Table.Combine({
        A1IssueDetails(MissingKeys, "Cluster lookup", "Missing Resource/day/index cluster key", "Resource spare excluded"),
        A1IssueDetails(DuplicateRows, "Cluster lookup", "Duplicate Resource/index cluster key", "Resource spare excluded"),
        A1IssueDetails(Unbalanced, "Cluster boundaries", "Start/end counts are unbalanced", "Resource spare excluded"),
        A1ClusterBoundaryIssues_Prepare
    })
in
    Table.Buffer(Details);

// Query: A1ClusterLookup_CHECK
// Purpose: Validate the Resource-scoped lookup grain and the paired cluster-boundary sequence.
shared A1ClusterLookup_CHECK = let
    LookupKeys = Table.SelectColumns(#"AllocationCode", {"Resource", "ResIndex", "Day"}),
    StartCounts = Table.Group(ClusterIDStart, {"Resource"}, {{"StartCount", each Table.RowCount(_), Int64.Type}}),
    EndCounts = Table.Group(ClusterIDEnd, {"Resource"}, {{"EndCount", each Table.RowCount(_), Int64.Type}}),
    BoundaryCounts = Table.NestedJoin(StartCounts, {"Resource"}, EndCounts, {"Resource"}, "EndCounts", JoinKind.FullOuter),
    ExpandedBoundaryCounts = Table.ExpandTableColumn(BoundaryCounts, "EndCounts", {"EndCount"}, {"EndCount"}),
    NormalizedBoundaryCounts = Table.ReplaceValue(ExpandedBoundaryCounts, null, 0, Replacer.ReplaceValue, {"StartCount", "EndCount"}),
    Checks = {
        A1CheckResult("Cluster lookup Resources are populated", () => Table.RowCount(Table.SelectRows(LookupKeys, each [Resource] = null or [ResIndex] = null))),
        A1CheckResult("Cluster lookup days are populated", () => Table.RowCount(Table.SelectRows(LookupKeys, each [Day] = null))),
        A1CheckResult("Cluster lookup keys are unique", () => Table.RowCount(LookupKeys) - Table.RowCount(Table.Distinct(LookupKeys, {"Resource", "ResIndex"}))),
        A1CheckResult("Cluster boundaries are balanced within each Resource", () => Table.RowCount(Table.SelectRows(NormalizedBoundaryCounts, each [StartCount] <> [EndCount]))),
        A1CheckResult("Cluster start/end dates are paired within each Resource", () => Table.RowCount(Table.SelectRows(A1ClusterBoundaryIssues_Prepare, each [Check] = "Cluster boundary pairing"))),
        A1CheckResult("Cluster boundary identifiers align within each Resource", () => Table.RowCount(Table.SelectRows(A1ClusterBoundaryIssues_Prepare, each [Check] = "Cluster boundary identifiers"))),
        A1CheckResult("Cluster intervals are ordered and non-overlapping", () => Table.RowCount(Table.SelectRows(A1ClusterBoundaryIssues_Prepare, each [Check] = "Cluster boundary order"))),
        A1CheckResult("Cluster boundaries fall on qualifying allocated days", () => Table.RowCount(Table.SelectRows(A1ClusterBoundaryIssues_Prepare, each [Check] = "Cluster boundary endpoints"))),
        A1CheckResult("Qualifying allocated days have exactly one cluster interval", () => Table.RowCount(Table.SelectRows(A1ClusterBoundaryIssues_Prepare, each [Check] = "Cluster boundary work coverage")))
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    // Cluster workday classification cannot be checked without the shared threshold.
    ParameterAwareChecks = if WorkingDayAllocationThreshold_Status[Status] = "Pass" then CheckTable
        else #table(type table [Check = text, Status = text, Failures = nullable number, Details = nullable text],
            List.Transform(List.Transform(Checks, each [Check]),
                each {_, "NotEvaluated", null, "WorkingDayAllocationThreshold is unavailable; Resource spare excluded"})),
    Buffered = Table.Buffer(ParameterAwareChecks)
in
    Buffered;

[ Description = "BUFFER" ]
// Query: Clusters
// Purpose: Assign cluster numbers for evaluable Resources while publishing exceptions for the others.
shared Clusters = let
    AffectedResources = Table.Distinct(Table.SelectColumns(A1ClusterIssues_Prepare, {"Resource"})),
    EvaluableRows = Table.NestedJoin(AllocationCode, {"Resource"}, AffectedResources, {"Resource"}, "ClusterIssue", JoinKind.LeftAnti),
    Source = Table.Sort(Table.RemoveColumns(EvaluableRows, {"ClusterIssue"}), {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Day", "ClusterPoints"}, #"ClusterIDStart+End", {"Resource", "Day", "ClusterPoints"}, "ClusterIDStart+End", JoinKind.LeftOuter),
    #"Expanded ClusterIDStart+End" = Table.ExpandTableColumn(#"Merged Queries", "ClusterIDStart+End", {"ClusterIndex"}, {"ClusterIndex"}),
    // Fill only in numeric day order inside each Resource; grouping does not guarantee row order.
    #"Grouped for Resource Fill Down" = Table.Group(#"Expanded ClusterIDStart+End", {"Resource"}, {{"Rows", each
        let
            OrderedResourceDays = Table.Sort(_, {{"Day", Order.Ascending}}),
            FilledResourceDays = Table.FillDown(OrderedResourceDays, {"ClusterIndex"})
        in
            FilledResourceDays}}),
    #"Filled Down" = if Table.IsEmpty(#"Expanded ClusterIDStart+End") then #"Expanded ClusterIDStart+End" else Table.Combine(#"Grouped for Resource Fill Down"[Rows]),
    #"Renamed Columns" = Table.RenameColumns(#"Filled Down",{{"ClusterIndex", "ClusterIndexX"}, {"AllocFlag", "Allocation"}}),
    #"Added NO ALLOCATION NULL" = Table.AddColumn(#"Renamed Columns", "ClusterIndex", each if [Allocation] = 0 then null else [ClusterIndexX]),
    #"Grouped for Resource Fill Up" = Table.Group(#"Added NO ALLOCATION NULL", {"Resource"}, {{"Rows", each
        let
            OrderedResourceDays = Table.Sort(_, {{"Day", Order.Ascending}}),
            FilledResourceDays = Table.FillUp(OrderedResourceDays, {"ClusterIndex"})
        in
            FilledResourceDays}}),
    #"Filled Up" = if Table.IsEmpty(#"Added NO ALLOCATION NULL") then #"Added NO ALLOCATION NULL" else Table.Combine(#"Grouped for Resource Fill Up"[Rows]),
    #"Sorted Filled Rows" = Table.Sort(#"Filled Up", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Filled Rows",{"ClusterIndexX", "ClusterPoints"}),

    // Preserve the established Code offsets while representing "00000" as no lookup.
    #"Added Lookup Offset" = Table.AddColumn(#"Removed Columns", "LookupOffset", each
        if [Code] = "00000" then null else
        if List.Contains({"11000", "11011", "01010", "01001", "01000", "11001", "11010"}, [Code]) then -1 else
        if List.Contains({"10000", "10001", "10010"}, [Code]) then -2 else
        if List.Contains({"00011", "00010"}, [Code]) then 1 else
        if List.Contains({"00001", "10011"}, [Code]) then 2 else
        0,
        Int64.Type
    ),
    #"Added Lookup ResIndex" = Table.AddColumn(#"Added Lookup Offset", "LookupResIndex", each
        if [LookupOffset] = null then null else [ResIndex] + [LookupOffset],
        Int64.Type
    ),
    #"Selected Cluster Lookup Columns" = Table.Buffer(Table.SelectColumns(#"Removed Columns", {"Resource", "ResIndex", "ClusterIndex"})),
    #"Joined Resource Cluster" = Table.NestedJoin(
        #"Added Lookup ResIndex",
        {"Resource", "LookupResIndex"},
        #"Selected Cluster Lookup Columns",
        {"Resource", "ResIndex"},
        "ClusterMatch",
        JoinKind.LeftOuter
    ),
    #"Added ClusterNumber" = Table.AddColumn(#"Joined Resource Cluster", "ClusterNumber", each
        if [LookupOffset] = null or Table.IsEmpty([ClusterMatch]) then "X"
        else if Table.RowCount([ClusterMatch]) = 1 then [ClusterMatch]{0}[ClusterIndex]
        else "X"
    ),
    #"Removed Lookup Helpers" = Table.RemoveColumns(#"Added ClusterNumber", {"LookupOffset", "LookupResIndex", "ClusterMatch"}),
    BUFFER = Table.Buffer(#"Removed Lookup Helpers")
in
    BUFFER;

shared ClusterAllocation = let
    Source = Clusters,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Allocation", Int64.Type}}),
    #"Grouped CLUSTERSIZE" = Table.Group(#"Changed Type", {"ClusterIndex"}, {{"ClusterAllocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped CLUSTERSIZE", each ([ClusterIndex] <> null)),
    BUFFER = Table.Buffer(#"Filtered Rows")
in
    BUFFER;

[ Description = "BUFFER #(lf)S2, P2 dev debt #(lf)Opportunity to discriminate with small Cluster Allocation#(lf)i.e. change the cluster it is associated with" ]
shared #"WD Type" = let
    Source = Clusters,
    Custom1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(Custom1, {"ClusterNumber"}, ClusterAllocation, {"ClusterIndex"}, "ClusterAllocation", JoinKind.LeftOuter),
    #"Expanded ClusterAllocation" = Table.ExpandTableColumn(#"Merged Queries", "ClusterAllocation", {"ClusterAllocation"}, {"ClusterAllocation"}),
    #"BUFFER LIST" = List.Buffer(#"Expanded ClusterAllocation"[Code]),
    // S2, P2 dev debt 
    // Opportunity to discriminate with small Cluster Allocation
    #"Added NWD TYPE" = Table.AddColumn(#"Expanded ClusterAllocation", "NWDType", each if [Allocation] = 1  then "Allocated" 

        else if [Code] = "00000" then "Reducable"

    //S1
        else if [Code] = "01000" 
            and [ResourceStartEnd] <> "ResourceStart1"  
        then "S1"

        else if [Code] = "11000" 
            and [ResourceStartEnd]<>"ResourceStart1" 
        then "S1"


     //S2       
        else if [Code] = "10000" 
            and [ResourceStartEnd] <> "ResourceStart2"  
            and [ResourceStartEnd] <> "ResourceStart1"         
        then "S2"

    //S2 P2

        else if [Code] = "10001" 
            and [ResourceStartEnd]<>"ResourceStart2" 
            and [ResourceStartEnd]<>"ResourceEnd2"         
        then "S2, P2"


    //P2
        else if (
                    [Code] = "00001" 
                    and [ResourceStartEnd] <> "ResourceEnd2" 
                    and [ResourceStartEnd] <> "ResourceEnd1"
                )
                or 
                (
                    [Code] = "11001" 
                    and [ResourceStartEnd] = "ResourceEnd2" 
                )
        then "P2"


    //P1
     
        else if (
                    ([Code] = "00010" or [Code] = "00011" )
                    and [ResourceStartEnd] <> "ResourceEnd1"    
                )  
                or 
                (
                    [Code] = "11011"
                    and [ResourceStartEnd] = "ResourceEnd1"
                )
        then "P1"
        

    //2d Off

        else if [Code] = "01001" 
            then
                if [ResourceStartEnd]="ResourceStart1"  
                then "P2"
            else "P2,S1"

        else if [Code] = "10010" 
            then 
                if [ResourceStartEnd]="ResourceEnd1"  
                then  "S2"
            else "P2,S1"
        
        else if [Code] = "10011"
            then  
                if [ResourceStartEnd]="ResourceStart1"
                then  "P1"
            else "P2,S1"

        else if [Code] = "11001" 
            then 
                if [ResourceStartEnd] = "ResourceEnd1"
                then "S1"
            else "P2"
            

    //Conditional

        
        else if [Code] = "01010" 
               and ("Yesterday.Alloction3" = 1 
               or "Tomorrow.Alloction3" = 1) 
               then "1D Off Conditional"

        else if [Code] = "01011" 
               and ("Yesterday.Alloction3" = 0 
               or "Tomorrow.Alloction3" = 0) 
               then "1D Off Conditional"


        
    //1D Off

        else if [Code] = "11010" 
            and [ResourceStartEnd] <> "ResourceEnd1" 
            and [ResourceStartEnd] <> "ResourceStart1"
        then "1D Off"
     
        else if [Code] = "11011" 
            and [ResourceStartEnd] <> "ResourceEnd1" 
            and [ResourceStartEnd] <> "ResourceStart1"
        then "1D Off"
              
        else "Reducable"),
    #"Removed Other Columns" = Table.SelectColumns(#"Added NWD TYPE",{"Resource", "Day",  "ClusterNumber", "ClusterAllocation", "NWDType"})
in
    #"Removed Other Columns";

// Query: ResDayPotentialAvailabilityTABLE
// Purpose: Preserve the existing day decisions and their cluster-size context for downstream priorities.
// Notes: U/R classification rules are unchanged; coordinated day eligibility is repaired in the next stage.
shared ResDayPotentialAvailabilityTABLE = let
    Source = #"WD Type",
  


    #"POTENTIAL ALLOCATION" = Table.AddColumn(Source, "PotentialAvailability", each if    [NWDType] = "Allocated" then "U"
    
    else if [NWDType] = "Reducable" then "R"

    else if [ClusterAllocation] >= MaxShiftCluster then "R"
    
    else if [ClusterAllocation] = (MaxShiftCluster -1)
        and ([NWDType] = "1D Off"   or  [NWDType] <> "S1" or [NWDType] <> "P1") 
        then "R" 
    
    else if [ClusterAllocation] <= (MaxShiftCluster -2)
        and  ([NWDType] = "1D Off"   
        or  [NWDType] <> "S1" 
        or [NWDType] <> "P1" 
        or [NWDType] <> "S2"  
        or [NWDType] <> "S2,P2" 
        or [NWDType] <> "P1,S2"
        or [NWDType] <> "P2,S1")
        then "R"
      

   
    else "U"),
    #"Removed Other Columns" = Table.SelectColumns(#"POTENTIAL ALLOCATION",{"Resource", "Day", "NWDType", "PotentialAvailability", "ClusterAllocation"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

[ Description = "BUFFER" ]
// Query: ResPeriodShiftNWDTABLE
// Purpose: Publish day priorities with cluster-size context at one row per Resource/Period.
shared ResPeriodShiftNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"PotentialAvailability"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"NWDType"}, #"SET NWDPriority", {"NWD Type"}, "NWDPrioritiesTABLE", JoinKind.LeftOuter),
    #"Expanded NWDPrioritiesTABLE" = Table.ExpandTableColumn(#"Merged Queries", "NWDPrioritiesTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Sorted Rows" = Table.Sort(#"Expanded NWDPrioritiesTABLE",{{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Day"}, #"PeriodShiftDay B", {"Day"}, "Period.1", JoinKind.LeftOuter),
    #"Expanded PERIODSHIFTS" = Table.ExpandTableColumn(#"Merged Queries1", "Period.1", {"Period", "Shifts"}, {"Period", "Shifts"}),
    SortedPeriods = Table.Sort(#"Expanded PERIODSHIFTS", {{"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    SortedPeriods;

shared ResDayNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Resource", "Day", "PotentialAvailability"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "U")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Resource", "Day", "PotentialAvailability"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Other Columns", {"Day"}, #"PeriodShiftDay B", {"Day"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodShiftDay", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE (2)", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE (2)", {"Allocation"}, {"Allocation"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, #"C' ResSingleShiftDayTABLE", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded C'" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Sorted Rows" = Table.Sort(#"Expanded C'",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Replaced Value" = Table.ReplaceValue(#"Sorted Rows",null,0,Replacer.ReplaceValue,{"Allocation", "Availability"}),
    #"Added ROSTEREDPERIODSTATUS" = Table.AddColumn(#"Replaced Value", "RosteredPeriodStatus", each if ([Availability] = null or [Availability] = 0) 
and 
[Allocation] > 0  
then "UnavailableAllocatedPeriod"

else if [Allocation] = 0 then "RosteredDay" 
else if [Allocation] <> 0 then "AllocatedPeriod" 
else null),

    #"Removed Columns" = Table.RemoveColumns(#"Added ROSTEREDPERIODSTATUS",{"Allocation"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER;

shared UnavailableAllocated = let
    Source = Table.NestedJoin(#"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.FullOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(Source, "ResPeriodAllocationTABLE", {"Resource", "Period", "Allocation"}, {"Resource.1", "Period.1", "Allocation"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResPeriodAllocationTABLE", each ([Availability] = null or [Availability] = 0 ) and ([Allocation] <> null )),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Resource", "Period", "Availability"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Resource.1", "Resource"}, {"Period.1", "Period"}}),
    #"Added UNAVAILABLEALLOCATIONPERIOD" = Table.AddColumn(#"Renamed Columns", "UnavailableAllocated", each "UnavailableAllocationPeriod")
in
    #"Added UNAVAILABLEALLOCATIONPERIOD";

shared #"SingleShiftDay-initial" = let
    Source = ResDayAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([DayShiftCount] = 1))
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
shared #"MultiDayPeriod-Available+Allocation" = let
    Source = ResDayAvailabilityTABLE,
    #"Filtered MULTIDAY" = Table.SelectRows(Source, each ([DayShiftCount] <> 1)),
    BUFFER = Table.Buffer(#"Filtered MULTIDAY"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Resource", "Day"}, ResDayPeriodAvailabilityTABLE, {"Resource", "Day"}, "ResDayPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayPeriodAvailabilityTABLE", {"Period"}, {"AvailablePeriod"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResDayPeriodAvailabilityTABLE", {"Resource", "AvailablePeriod"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded ResPeriodAllocationTABLE",{"DayAvailbility", "DayShiftCount"})
in
    #"Removed Columns";

shared MultiShiftAvilabilityDay = let
    Source = #"MultiDayPeriod-Available+Allocation",
    #"Removed MULTIPERIODALLOCATION" = Table.RemoveColumns(Source,{"Allocation"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed MULTIPERIODALLOCATION", {"Resource", "Day"}, ResDayAllocationTABLE, {"Resource", "Day"}, "ResDayAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayAllocationTABLE", {"ResDayAllocation"}, {"ResDayAllocation"}),
    #"Filtered UNMACTHED PERIODS REMOVED" = Table.SelectRows(#"Expanded ResDayAllocationTABLE", each ([ResDayAllocation] = null))
in
    #"Filtered UNMACTHED PERIODS REMOVED";

// Query: MultishiftAvailDayPeriod-Index
// Purpose: Rank available periods by capacity-to-demand ratio with explicit zero-demand protection.
shared #"MultishiftAvailDayPeriod-Index" = let
    Source = MultiShiftAvilabilityDay,
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "AvailablePeriod"}, ResDayPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResDayPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResDayPeriodAvailabilityTABLE", {"AvailablePeriod"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodCapacityTABLE", {"Capacity"}, {"Capacity"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded PeriodCapacityTABLE", {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "PeriodDemandTABLE", {"D"}, {"D"}),
    // Preserve the established zero fallback; the priority semantics of excluding zero-demand periods remain a separate business decision.
    #"Inserted C/D" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C/D", each if [D] = null or [D] = 0 then 0 else [Capacity] / [D], type number),
    #"Merged Queries3" = Table.NestedJoin(#"Inserted C/D", {"Resource"}, #"ResPeriodAllocation-ShiftBias", {"Resource"}, "ResPeriodAllocation-ShiftBias", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocation-ShiftBias" = Table.ExpandTableColumn(#"Merged Queries3", "ResPeriodAllocation-ShiftBias", {"ShiftBias"}, {"ShiftBias"}),
    #"Sorted RES, PERIOD C/D" = Table.Sort(#"Expanded ResPeriodAllocation-ShiftBias",{{"Resource", Order.Ascending}, {"C/D", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES, PERIOD C/D", "Index", 1, 1, Int64.Type),
    #"Removed Other Columns" = Table.SelectColumns(#"Added Index",{"Resource", "Day", "AvailablePeriod", "C/D", "ShiftBias", "Index", "Role"})
in
    #"Removed Other Columns";

shared #"MultiShiftAvailDayPeriod -0 Allocation" = let
    Source = #"MultishiftAvailDayPeriod-Index",
    BUFFER1 = Table.Buffer(Source),
    #"Grouped RESDAYS" = Table.Group(BUFFER1, {"Resource", "Day"}, {{"ResGroup", each _, type table [Role=nullable text, Resource=nullable number, Day=number, AvailablePeriod=nullable number, ResDayAllocation=nullable number, Availability=number, Capacity=nullable number, D=nullable number, #"C/D"=number, ShiftBias=nullable number]}}),
    #"Added DAY INDEX" = Table.AddColumn(#"Grouped RESDAYS", "Group DAY INDEX", each Table.AddIndexColumn([ResGroup],"ResC/DPriority",1)),
    #"Expanded DAYINDEX" = Table.ExpandTableColumn(#"Added DAY INDEX", "Group DAY INDEX", {"AvailablePeriod", "C/D", "ResC/DPriority", "ShiftBias"}, {"AvailablePeriod", "C/D", "ResC/DPriority", "ShiftBias"}),
    #"Added NO ALLOCATION KEEP" = Table.AddColumn(#"Expanded DAYINDEX", "NoAllocationKeep", each if [#"ResC/DPriority"] = 1 then true else false),
    #"Changed Type" = Table.TransformColumnTypes(#"Added NO ALLOCATION KEEP",{{"NoAllocationKeep", type logical}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Changed Type",{"ResGroup", "ResC/DPriority"})
in
    #"Removed Columns3";

shared #"SingleShiftDayAvail+MultiAllocation B" = let
    Source = #"SingleShiftDay-initial",
    #"Merged Queries" = Table.NestedJoin(Source, {"Day"}, #"PeriodShiftDay B", {"Day"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodShiftDay", {"Resource", "Period", "Role"}, ResPeriodAllocationTABLE, {"Resource", "Period", "Role"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded ResPeriodAllocationTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Filtered ALLOCATED PERIODS" = Table.SelectRows(#"Sorted Rows1", each ([Allocation] <> null )),
    #"Merged Queries2" = Table.NestedJoin(#"Filtered ALLOCATED PERIODS", {"Resource", "Period"}, #"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Filtered NULL AVAILABILITY" = Table.SelectRows(#"Expanded ResPeriodAvailabilityTABLE", each ([Availability] <> null )),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NULL AVAILABILITY",{"Period", "Resource"})
in
    #"Removed Other Columns";

shared SingleShiftDayPeriodMatch = let
    Source = DayShiftMatchSPLIT,
    #"Filtered Rows" = Table.SelectRows(Source, each ([MatchCount] = "SingleShiftMatch")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"DayMatches", "MatchCount"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Day", "Resource"}, ResDayPeriodAvailabilityTABLE, {"Day", "Resource"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailability" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period", "Availability"}, {"Period", "Availability"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded ResDayPeriodAvailability",{"Day"}),
    #"Added MATCHPERIOD?" = Table.AddColumn(#"Removed Columns1", "MatchedPeriood", each if [MatchPeriod] = [Period] then true else false, type logical),
    #"Filtered Rows1" = Table.SelectRows(#"Added MATCHPERIOD?", each ([MatchedPeriood] = true)),
    #"Removed Columns2" = Table.RemoveColumns(#"Filtered Rows1",{"MatchedPeriood", "MatchPeriod", "Availability"})
in
    #"Removed Columns2";

[ Description = "BUFFER" ]
shared #"C' ResSingleShiftDayTABLE" = let
    Source = Table.Combine({MultiDayPeriodMatchTABLE, SingleShiftDayPeriods}),
    #"Appended MULTIDAYPERIOD-REDUCIBLE" = Table.Combine({Source, #"MultiDayPeriod-Reducibile"}),
    #"Removed Duplicates" = Table.Distinct(#"Appended MULTIDAYPERIOD-REDUCIBLE"),
    #"Merged Queries" = Table.NestedJoin(#"Removed Duplicates", {"Resource", "Period"}, #"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"BUFFER Expanded ResPeriodAvailabilityTABLE" = Table.Buffer(Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}))
in
    #"BUFFER Expanded ResPeriodAvailabilityTABLE";

shared #"C' SUM" = let
    Source = List.Sum(#"C' ResSingleShiftDayTABLE"[Availability])
in
    Source;

shared #"MultiDayPeriod-AvailableAllocatedMatches" = let
    Source = #"MultiDayPeriod-Available+Allocation",
  #"Filtered NULL ALLOCATION PERIOD" = Table.SelectRows(Source, each [Allocation] <> null)
in
  #"Filtered NULL ALLOCATION PERIOD"
  //Source
;

[ Description = "BUFFER" ]
shared DayShiftMatchSPLIT = let
    Source = #"MultiDayPeriod-AvailableAllocatedMatches",
    #"Grouped DAYMATCHES" = Table.Group(Source, {"Resource", "Day"}, {{"DayMatches", each Table.RowCount(_), Int64.Type}, {"MatchPeriod", each List.Average([AvailablePeriod]), type nullable number}}),
    #"SPLIT BUFFER" = Table.Buffer(Table.AddColumn(#"Grouped DAYMATCHES", "MatchCount", each if [DayMatches] = 1 then "SingleShiftMatch" else "MultipleShiftMatch"))
in
    #"SPLIT BUFFER";

shared MultiShiftDayMatches = let
    Source = DayShiftMatchSPLIT,
    #"Filtered Rows" = Table.SelectRows(Source, each ([MatchCount] = "MultipleShiftMatch")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"MatchCount", "MatchPeriod"})
in
    #"Removed Columns";

shared MultiShiftDayPeriodMatches = let
    Source = MultiShiftDayMatches,
    #"Merged Queries1" = Table.NestedJoin(Source, {"Resource", "Day"}, #"MultiDayPeriod-AvailableAllocatedMatches", {"Resource", "Day"}, "MultiDayPeriod-AvailableAllocatedMatches", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches" = Table.ExpandTableColumn(#"Merged Queries1", "MultiDayPeriod-AvailableAllocatedMatches", {"AvailablePeriod"}, {"AvailablePeriod"})
in
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches";

shared MultiPeriodPriority = let
    Source = MultiShiftDayPeriodMatches[[AvailablePeriod]],
    #"Removed Duplicates" = Table.Distinct(Source),
    #"Merged Queries" = Table.NestedJoin(#"Removed Duplicates", {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodDemandTABLE", {"AvailablePeriod"}, PeriodAllocationTABLE, {"Period"}, "PeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodAllocationTABLE", {"A"}, {"A"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded PeriodAllocationTABLE", {"AvailablePeriod"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "PeriodCapacityTABLE", {"Capacity"}, {"Capacity"}),
    #"Inserted Division" = Table.AddColumn(#"Expanded PeriodCapacityTABLE", "C/D", each [Capacity] / [D], type number),
    #"Inserted Division1" = Table.AddColumn(#"Inserted Division", "A/D", each [A] / [D], type number),
    #"Sorted C/D A/D" = Table.Sort(#"Inserted Division1",{{"C/D", Order.Ascending}, {"A/D", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted C/D A/D", "PeriodPriority", 1, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"D", "A", "Capacity", "C/D", "A/D"})
in
    #"Removed Columns";

[ Description = "BUFFER" ]
shared MultiResPeriodPriority = let
    Source = Table.NestedJoin(MultiShiftDayPeriodMatches, {"AvailablePeriod"}, MultiPeriodPriority, {"AvailablePeriod"}, "MultiPeriodPriority", JoinKind.LeftOuter),
    #"Expanded MultiPeriodPriority" = Table.ExpandTableColumn(Source, "MultiPeriodPriority", {"PeriodPriority"}, {"PeriodPriority"}),
    BUFFER = Table.Buffer(#"Expanded MultiPeriodPriority")
in
    BUFFER;

shared #"MultiResPeriodPriority-Highest" = let
    Source = MultiResPeriodPriority,
    #"Grouped Rows" = Table.Group(Source, {"Resource", "Day"}, {{"HighestPriority", each List.Max([PeriodPriority]), type nullable number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared MultiDayPeriodMatchTABLE = let
    Source = Table.NestedJoin(MultiResPeriodPriority, {"Resource", "Day"}, #"MultiResPeriodPriority-Highest", {"Resource", "Day"}, "MultiResPriodPriority-Highest", JoinKind.LeftOuter),
    #"Expanded MultiResPriodPriority-Highest" = Table.ExpandTableColumn(Source, "MultiResPriodPriority-Highest", {"HighestPriority"}, {"HighestPriority"}),
    #"Added Custom" = Table.AddColumn(#"Expanded MultiResPriodPriority-Highest", "Highest?", each [PeriodPriority]=[HighestPriority]),
    #"Filtered HIGHEST PRIORITY" = Table.SelectRows(#"Added Custom", each ([#"Highest?"] = true)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered HIGHEST PRIORITY",{"DayMatches", "PeriodPriority", "HighestPriority", "Highest?"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailablePeriod", "Period"}}),
    #"Removed Columns1" = Table.Buffer(Table.RemoveColumns(#"Renamed Columns",{"Day"}))
in
    #"Removed Columns1";

shared SingleShiftDayPeriods = let
    // SingleShiftDayAvail+Multi + SingleDayShiftMatchTABLE
    Source = Table.Combine({SingleShiftDayPeriodMatch, #"SingleShiftDayAvail+MultiAllocation B"}),
    #"Sorted Rows" = Table.Sort(Source,{{"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

[ Description = "BUFFER" ]
shared #"MultiDayPeriod-Reducibile" = let
    Source = #"MultiShiftAvailDayPeriod -0 Allocation",
    #"Filtered Rows" = Table.SelectRows(Source, each ([NoAllocationKeep] = true)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Day", "NoAllocationKeep", "C/D"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailablePeriod", "Period"}})
in
    #"Renamed Columns";

shared #"MultiDayPeriod-Remove" = let
    Source = #"MultiShiftAvailDayPeriod -0 Allocation",
    #"Filtered Rows" = Table.SelectRows(Source, each [NoAllocationKeep] = false),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{ "Day"})
in
    #"Removed Columns";

shared PeriodAllocated = let
    Source = ResPeriodAllocationTABLE,
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, #"PeriodShiftDay B", {"Period"}, "PeriodShiftDay", JoinKind.LeftOuter)
in
    #"Merged Queries";

shared PeriodCapacityTABLE = let
    Source = #"ResPeriodAvailabilityTABLE !!",
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"Capacity", each List.Sum([Availability]), type number}}),
    BUFFER = Table.Buffer(#"Grouped Rows")
in
    BUFFER;

shared ResPeriodAllocationSUM = let
    Source = List.Sum(PeriodAllocationTABLE[A])
in
    Source;

[ Description = "BUFFER" ]
shared #"Demand Prepare" = let
    Source = #"IMPORT ShiftUnitDemandHRS !!",
    #"Renamed Columns1" = Table.RenameColumns(Source,{{"ShiftPeriod", "Shift"}, {"ShiftDemandHCAverage", "DemandFTE"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"UnitShiftEffort", "ShiftDurations.ShiftDuration"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Columns",{"Shift", "Role", "Facility", "DemandFTE", "Date"}),
    BUFFER = Table.Buffer(#"Removed Other Columns"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Date", "Shift"}, #"PeriodShiftDay B", {"Date", "Shifts"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded PeriodShiftDay",{{"DemandFTE", "D"}})
in
    #"Renamed Columns";

shared #"PeriodDemandTABLE !!" = let
    Source = #"Demand Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Shift", "Role", "D", "Period"})
in
    #"Removed Other Columns";

shared PeriodDemandSUM = let
    Source = List.Sum(#"PeriodDemandTABLE !!"[D])
in
    Source;

// Query: A1PriorityInterface_CHECK
// Purpose: Report priority handoff schema and key failures without deliberately stopping publication.
// Notes: Null cluster sizes are allowed on days outside allocated clusters; calendar eligibility rules are unchanged.
shared A1PriorityInterface_CHECK = let
    DayPriorities = ResPeriodShiftNWDTABLE,
    CapPriorities = ResPeriodCapPrioritised,
    DayColumns = {"Resource", "Day", "Period", "Shifts", "ClusterAllocation", "NWDPriority", "NWDType"},
    CapColumns = {"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "ClusterAllocation", "NWDPriority", "NWDType"},
    Checks = {
        A1CheckResult("NWD priority handoff has required columns", () => List.Count(List.Difference(DayColumns, Table.ColumnNames(DayPriorities)))),
        A1CheckResult("NWD priority keys are populated", () => Table.RowCount(Table.SelectRows(DayPriorities, each [Resource] = null or [Period] = null))),
        A1CheckResult("NWD priority keys are unique", () => Table.RowCount(DayPriorities) - Table.RowCount(Table.Distinct(DayPriorities, {"Resource", "Period"}))),
        A1CheckResult("Cap priority handoff has required columns", () => List.Count(List.Difference(CapColumns, Table.ColumnNames(CapPriorities)))),
        A1CheckResult("Cap priority keys are populated", () => Table.RowCount(Table.SelectRows(CapPriorities, each [Resource] = null or [Period] = null))),
        A1CheckResult("Cap priority keys are unique", () => Table.RowCount(CapPriorities) - Table.RowCount(Table.Distinct(CapPriorities, {"Resource", "Period"})))
    },
    Output = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])
in
    Table.Buffer(Output);

// Query: A1DiagnosticSummary
// Purpose: Publish an independently evaluated check table or an explicit evaluation-error row.
shared A1DiagnosticSummary = (check as text, evaluate as function) as table =>
let
    Attempt = try Table.Buffer(evaluate()),
    Checks = if Attempt[HasError] then
            #table(type table [Check = text, Status = text, Failures = nullable number, Details = nullable text],
                {{check, "Error", null, try Attempt[Error][Message] otherwise "Evaluation unavailable"}})
        else Attempt[Value],
    WithStage = Table.AddColumn(Checks, "Stage", each "A.1", type text),
    Output = Table.ReorderColumns(WithStage, {"Stage", "Check", "Status", "Failures", "Details"})
in
    Output;

// Query: A1DiagnosticDetails
// Purpose: Keep other diagnostic modules readable when one module cannot be evaluated.
shared A1DiagnosticDetails = (check as text, evaluate as function) as table =>
let
    Attempt = try Table.Buffer(evaluate()),
    Output = if Attempt[HasError] then
            A1IssueDetails(#table(type table [Resource = nullable number], {{null}}), check,
                try Attempt[Error][Message] otherwise "Evaluation unavailable", "Not evaluated; inspect the source input")
        else Attempt[Value]
in
    Output;

// Query: CapacityDiagnostics_SUMMARY
// Purpose: Publish nonblocking A.1 check results for manual review and future runner collection.
// Output: Stage, Check, Status, Failures and Details; a failed check never deliberately stops other Resources.
shared CapacityDiagnostics_SUMMARY = let
    IdentityChecks = A1DiagnosticSummary("Resource identity checks", () => A1ResourceIdentity_CHECK),
    ContractChecks = A1DiagnosticSummary("Resource contract checks", () => A1ResourceContract_CHECK),
    ClusterChecks = A1DiagnosticSummary("Cluster lookup checks", () => A1ClusterLookup_CHECK),
    PriorityChecks = A1DiagnosticSummary("Priority handoff checks", () => A1PriorityInterface_CHECK),
    SpareChecks = A1DiagnosticSummary("Resource spare exclusions", () => Table.FromRecords({
        A1CheckResult("Resources excluded from spare calculations", () => Table.RowCount(A1ResourceSpareIssues_Prepare)),
        A1CheckResult("Malformed allocation amounts", () => Table.RowCount(A1AllocationValueIssues_Prepare)),
        A1CheckResult("Skipped allocation records", () => Table.RowCount(A1AllocationIdentity_EXCEPTIONS))
    }, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])),
    ParameterStatus = WorkingDayAllocationThreshold_Status,
    ParameterCheck = #table(type table [Stage = text, Check = text, Status = text, Failures = nullable number, Details = nullable text],
        {{"A.1", "Working-day allocation threshold", ParameterStatus[Status],
            if ParameterStatus[Status] = "Pass" then 0 else 1, ParameterStatus[Details]}}),
    Output = Table.Combine({IdentityChecks, ContractChecks, ClusterChecks, PriorityChecks, SpareChecks, ParameterCheck})
in
    Table.Buffer(Output);

// Query: CapacityDiagnostics_DETAILS
// Purpose: Publish affected identities/Resources and the exclusions applied, independently of capacity outputs.
shared CapacityDiagnostics_DETAILS = let
    IdentityDetails = A1DiagnosticDetails("Resource identity", () => A1IdentityIssues_Prepare),
    AllocationDetails = A1DiagnosticDetails("Allocation amount", () => A1AllocationValueIssues_Prepare),
    ContractDetails = A1DiagnosticDetails("Resource contract", () => A1ContractIssues_Prepare),
    PeriodDetails = A1DiagnosticDetails("Availability period", () => A1AvailabilityPeriodIssues_Prepare),
    ClusterDetails = A1DiagnosticDetails("Cluster lookup", () => A1ClusterIssues_Prepare),
    ParameterStatus = WorkingDayAllocationThreshold_Status,
    ParameterDetails = if ParameterStatus[Status] = "Pass" then Table.FirstN(IdentityDetails, 0)
        else A1DiagnosticDetails("Working-day threshold", () =>
            A1IssueDetails(A1ResourceGrid_Prepare, "Working-day threshold", ParameterStatus[Details],
                "Resource spare not evaluated; allocation evidence retained")),
    Output = Table.Combine({IdentityDetails, AllocationDetails, ContractDetails, PeriodDetails, ClusterDetails, ParameterDetails})
in
    Table.Buffer(Output);

// Query: ValidAllocation
// Purpose: Keep the existing worked-day threshold name while reading its value from Effort Management Parameters.
// Notes: ResDaysAllocated retains its strict Allocation > ValidAllocation comparison; the initial list value is 0.7.
shared ValidAllocation = WorkingDayAllocationThreshold meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

shared #"PeriodC'SUM" = let
    Source = #"PeriodC'",
    #"PeriodC'1" = Source[#"PeriodC'"],
    #"Calculated Sum" = List.Sum(#"PeriodC'1")
in
    #"Calculated Sum";

// Query: RolePathTABLE
// Purpose: Resolve this workbook's folder and dynamic role through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value rows used by RolePath, FilePath and Role.
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
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "CapacityDistrib(A.1)-shifts.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate CapacityDistrib(A.1)-shifts.xlsx.",
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

shared RolePath = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    Folder = #"Renamed Columns"{0}[Folder]
in
    Folder;

shared FilePath = let
    Source = RolePath,
    #"Converted to Table" = #table(1, {{Source}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Converted to Table", {{"Column1", each Text.BeforeDelimiter(_, "\", {1, RelativePosition.FromEnd}), type text}}),
    Column1 = #"Extracted Text Before Delimiter"{0}[Column1]
in
    Column1;

shared Role = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Role")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: IMPORT Settings Data
// Purpose: Open Settings Data.xlsx once for all Settings table extractions.
shared #"IMPORT Settings Data" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Settings Data.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: EXTRACT PermutationDimensions
// Purpose: Navigate to the Settings permutation table without applying business transformations.
shared #"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data"{[Item="PermutationDimensions", Kind="Table"]}[Data]
in
    Source;

// Query: EXTRACT Date_From
// Purpose: Validate and return the single DateFrom setting used to scope allocations.
shared #"EXTRACT Date_From" = let
    Source = #"IMPORT Settings Data"{[Item="DateFrom", Kind="Table"]}[Data],
    RequiredColumn = if Table.HasColumns(Source, {"DateFrom"}) then Source
        else error "DateFrom must contain a DateFrom column.",
    SingleRow = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn
        else error "DateFrom must contain exactly one data row.",
    Typed = Table.TransformColumnTypes(SingleRow, {{"DateFrom", type date}}),
    Value = Typed{0}[DateFrom]
in
    Value;

// Query: EXTRACT MaxAvailability
// Purpose: Preserve the legacy Settings maximum interface; employee Effective Shift Cap is canonical for A.1.
// Notes: Retained for compatibility and no longer used by the cap calculation.
shared #"EXTRACT MaxAvailability" = let
    Source = #"IMPORT Settings Data"{[Item="MaxAvailability", Kind="Table"]}[Data],
    RequiredColumn = if Table.HasColumns(Source, {"MaxAvailability"}) then Source
        else error "MaxAvailability must contain a MaxAvailability column.",
    SingleRow = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn
        else error "MaxAvailability must contain exactly one data row.",
    Typed = Table.TransformColumnTypes(SingleRow, {{"MaxAvailability", Int64.Type}}),
    Value = Typed{0}[MaxAvailability]
in
    Value;
