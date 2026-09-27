section Section1;

// Query: IMPORT AllocationExtracted
// Purpose: Read the unit allocation rows already consumed by Settings, without the removed ResRolesProportion table.
shared #"IMPORT AllocationExtracted" = let
    SourcePath = #"FilePath - 1Input" & "\1-AllocationExtracted.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "AllocationExtracted" and [Kind] = "Table"),
    MatchCount = Table.RowCount(Matches),
    Allocation = if MatchCount = 1 then Matches{0}[Data]
        else error Error.Record("Allocation import", "Expected exactly one AllocationExtracted table.", [Matches = MatchCount]),
    RequiredColumns = {"Location", "Code", "Role", "Hours", "Name"},
    AvailableColumns = Table.ColumnNames(Allocation),
    MissingColumns = List.Difference(RequiredColumns, AvailableColumns),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(Allocation, RequiredColumns)
        else error Error.Record("Allocation import", "AllocationExtracted is missing role-share fields.", [MissingColumns = MissingColumns]),
    Normalized = Table.TransformColumns(Selected, {
        {"Location", AvailabilityText, type nullable text}, {"Code", AvailabilityText, type nullable text},
        {"Role", AvailabilityText, type nullable text}, {"Name", AvailabilityText, type nullable text}}),
    Buffered = Table.Buffer(Normalized)
in Buffered;

// Query: AvailabilityText
// Purpose: Normalize optional identifiers and names without converting text identifiers to numbers.
shared AvailabilityText = (value as any) as nullable text =>
    let
        CleanAttempt = try
            let
                Converted = Text.From(value),
                Trimmed = Text.Trim(Converted)
            in
                Trimmed,
        Clean = if CleanAttempt[HasError] then null else CleanAttempt[Value],
        BlankAsNull = if Clean = "" then null else Clean
    in
        BlankAsNull;

// Query: AvailabilityUnion
// Purpose: Merge duplicate, overlapping and touching half-open datetime intervals.
shared AvailabilityUnion = (intervals as list) as list =>
    let
        CompareIntervals = (a as record, b as record) as number =>
            let
                StartComparison = if a[Start] < b[Start] then -1
                    else if a[Start] > b[Start] then 1
                    else 0,
                EndComparison = if a[End] < b[End] then -1
                    else if a[End] > b[End] then 1
                    else 0,
                Comparison = if StartComparison <> 0 then StartComparison else EndComparison
            in
                Comparison,
        Ordered = List.Sort(intervals, CompareIntervals),
        MergeInterval = (state as list, current as record) as list =>
            let
                StateIsEmpty = List.IsEmpty(state),
                Previous = if StateIsEmpty then null else List.Last(state),
                TouchesPrevious = if StateIsEmpty then false else current[Start] <= Previous[End],
                StateWithoutPrevious = if TouchesPrevious then List.RemoveLastN(state, 1) else state,
                MergedEnd = if TouchesPrevious then List.Max({Previous[End], current[End]}) else null,
                MergedInterval = if TouchesPrevious then
                    [Start = Previous[Start], End = MergedEnd]
                    else null,
                UpdatedState = if StateIsEmpty then {current}
                    else if TouchesPrevious then StateWithoutPrevious & {MergedInterval}
                    else state & {current},
                // Materialize accumulators to avoid re-traversing lazy concatenations.
                BufferedState = List.Buffer(UpdatedState)
            in
                BufferedState,
        Merged = List.Accumulate(Ordered, {}, MergeInterval)
    in
        Merged;

// Query: AvailabilitySubtract
// Purpose: Remove exclusion intervals from available intervals, retaining both sides of an internal gap.
shared AvailabilitySubtract = (available as list, excluded as list) as list =>
    let
        SubtractCut = (segments as list, cut as record) as list =>
            let
                SubtractFromSegment = (segment as record) as list =>
                    let
                        DoesNotOverlap = cut[End] <= segment[Start] or cut[Start] >= segment[End],
                        SegmentBeforeCut = if cut[Start] > segment[Start] then
                            {[Start = segment[Start], End = cut[Start]]}
                            else {},
                        SegmentAfterCut = if cut[End] < segment[End] then
                            {[Start = cut[End], End = segment[End]]}
                            else {},
                        RemainingParts = if DoesNotOverlap then {segment}
                            else SegmentBeforeCut & SegmentAfterCut
                    in
                        RemainingParts,
                SegmentParts = List.Transform(segments, SubtractFromSegment),
                RemainingSegments = List.Combine(SegmentParts)
            in
                RemainingSegments,
        Result = List.Accumulate(excluded, available, SubtractCut)
    in
        Result;

// Query: IMPORT Reconciled Workers
// Purpose: Read the authoritative worker population from the existing local reconciliation table.
shared #"IMPORT Reconciled Workers" = let
    SourcePath = #"FilePath - 1Input" & "\Worker Reconciliation.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    EmployeeTable = Navigation{[Item = "Employees_TABLE", Kind = "Table"]}[Data],
    Workers = Table.SelectColumns(EmployeeTable,
        {"EmployeeID", "Facility-Abbrev", "Role", "Employee Roster Name", "EmploymentType"}),
    WorkerColumns = Table.ColumnNames(Workers),
    TextTransformations = List.Transform(WorkerColumns,
        each {_, AvailabilityText, type nullable text}),
    Normalized = Table.TransformColumns(Workers, TextTransformations)
in Normalized;

// Query: IMPORT Worker's Report
// Purpose: Read the current unit's employee contract source once for downstream contract preparation.
// Inputs: Table1 in 1. Input/Worker's Report.xlsx.
// Output: Employee Code and Contracted FN Hours at source-row grain.
shared #"IMPORT Worker's Report" = let
    SourcePath = #"FilePath - 1Input" & "\Worker's Report.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "Table1" and [Kind] = "Table"),
    MatchCount = Table.RowCount(Matches),
    ContractTable = if MatchCount = 1 then Matches{0}[Data]
        else error Error.Record("Worker contract import", "Expected exactly one Table1 table in Worker's Report.xlsx.", [Matches = MatchCount]),
    RequiredColumns = {"Employee Code", "Contracted FN Hours"},
    AvailableColumns = Table.ColumnNames(ContractTable),
    MissingColumns = List.Difference(RequiredColumns, AvailableColumns),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(ContractTable, RequiredColumns)
        else error Error.Record("Worker contract import", "Worker's Report.xlsx/Table1 is missing required columns.", [MissingColumns = MissingColumns]),
    Buffered = Table.Buffer(Selected)
in Buffered;

// Query: IMPORT Facilities List
// Purpose: Read the Whiddon SharePoint facility reference for unit-level source filtering.
shared #"IMPORT Facilities List" =
    SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation=null, ApiVersion=15]);

// Query: LINK Facilities
// Purpose: Map each source Site Name to its canonical Facility-Abbrev.
shared #"LINK Facilities" =
let
    FacilityItems = #"IMPORT Facilities List"{[Id="1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items],
    SelectedColumns = Table.SelectColumns(FacilityItems, {"Title", "field_1"}),
    NamedColumns = Table.RenameColumns(SelectedColumns, {{"field_1", "Facility-Abbrev"}})
in
    NamedColumns;

// Query: LINK Facility Analysis
// Purpose: Select the same Facility Analysis fields used by AllocationExtracted.
shared #"LINK Facility Analysis" =
let
    AnalysisItems = #"IMPORT Facilities List"{[Id="17ed8e30-5707-4af2-8e21-d68eb8184287"]}[Items],
    SelectedColumns = Table.SelectColumns(AnalysisItems, {"Title", "EffortManagementAnalysis", "LeaveBalanceAnalysis"})
in
    SelectedColumns;

// Query: WorkerRoleMappings_Prepare
// Purpose: Resolve LINK Roles to one inspectable Role Group assignment per normalized Roster Roles key.
// Output: One row per RosterRoleKey with source counts, assigned Role Group and any blocking mapping issue.
// Notes: Identical duplicate definitions are collapsed without multiplying workers and remain visible in diagnostics.
shared WorkerRoleMappings_Prepare = let
    Selected = Table.SelectColumns(#"LINK Roles", {"Roster Roles", "Role Group"}),
    Normalized = Table.TransformColumns(Selected, {
        {"Roster Roles", AvailabilityText, type nullable text},
        {"Role Group", AvailabilityText, type nullable text}
    }),
    RosterRoleKeys = Table.AddColumn(Normalized, "RosterRoleKey",
        each WorkerRoleKey([Roster Roles]), type nullable text),
    // Ignore wholly empty list items; partially populated definitions remain visible as issues.
    NonEmptyDefinitions = Table.SelectRows(RosterRoleKeys,
        each [RosterRoleKey] <> null or [Role Group] <> null),
    CombineSortedDistinctText = (values as list) as nullable text =>
        let
            NonNullValues = List.RemoveNulls(values),
            DistinctValues = List.Distinct(NonNullValues),
            SortedValues = List.Sort(DistinctValues),
            CombinedValues = if List.IsEmpty(SortedValues) then null
                else Text.Combine(SortedValues, "; ")
        in
            CombinedValues,
    SummarizeMapping = (rows as table) as record =>
        let
            RosterRoleKey = rows{0}[RosterRoleKey],
            RosterRoleValues = Table.Column(rows, "Roster Roles"),
            RoleGroupValues = List.Sort(List.Distinct(
                List.RemoveNulls(Table.Column(rows, "Role Group")))),
            RoleMappingDefinitionRows = Table.RowCount(rows),
            DistinctRoleGroupCount = List.Count(RoleGroupValues),
            MappedRosterRole = CombineSortedDistinctText(RosterRoleValues),
            MappedRoleGroup = if DistinctRoleGroupCount = 1 then RoleGroupValues{0} else null,
            RoleMappingIssue = if RosterRoleKey = null then
                    "LINK Roles has a blank Roster Roles value"
                else if DistinctRoleGroupCount = 0 then
                    "LINK Roles mapping has no Role Group"
                else if DistinctRoleGroupCount > 1 then
                    "LINK Roles roster role maps to multiple Role Groups"
                else null,
            Summary = [
                MappedRosterRole = MappedRosterRole,
                RoleMappingDefinitionRows = RoleMappingDefinitionRows,
                DistinctRoleGroupCount = DistinctRoleGroupCount,
                MappedRoleGroup = MappedRoleGroup,
                RoleMappingIssue = RoleMappingIssue
            ]
        in
            Summary,
    GroupedMappings = Table.Group(NonEmptyDefinitions, {"RosterRoleKey"},
        {{"Mapping", SummarizeMapping, type [
            MappedRosterRole = nullable text,
            RoleMappingDefinitionRows = Int64.Type,
            DistinctRoleGroupCount = Int64.Type,
            MappedRoleGroup = nullable text,
            RoleMappingIssue = nullable text
        ]}}),
    ExpandedMappings = Table.ExpandRecordColumn(GroupedMappings, "Mapping", {
        "MappedRosterRole", "RoleMappingDefinitionRows", "DistinctRoleGroupCount",
        "MappedRoleGroup", "RoleMappingIssue"
    }),
    BufferedMappings = Table.Buffer(ExpandedMappings)
in
    BufferedMappings;

// Query: WorkerRoleMappings_DIAGNOSTICS
// Purpose: Show invalid or duplicated LINK Roles definitions before they are joined to workers.
shared WorkerRoleMappings_DIAGNOSTICS = let
    MappingIssues = Table.SelectRows(WorkerRoleMappings_Prepare, each
        [RoleMappingIssue] <> null or [RoleMappingDefinitionRows] > 1),
    SortedIssues = Table.Sort(MappingIssues, {{"RosterRoleKey", Order.Ascending}})
in
    SortedIssues;

// Query: FacilityAnalysisTABLE
// Purpose: Identify whether the folder-derived Unit is enabled for Effort Management Analysis.
shared FacilityAnalysisTABLE =
let
    Source = #"LINK Facility Analysis",
    IsCurrentUnit = (row as record) as logical =>
        let
            Title = row[Title],
            TrimmedTitle = if Title = null then null else Text.Trim(Title),
            Comparison = if TrimmedTitle = null then null
                else Comparer.OrdinalIgnoreCase(TrimmedTitle, Unit),
            MatchesUnit = Comparison = 0
        in
            MatchesUnit,
    UnitRows = Table.SelectRows(Source, IsCurrentUnit),
    UnitRowCount = Table.RowCount(UnitRows),
    ValidatedUnit = if UnitRowCount = 1 then UnitRows
        else error "Facility Analysis must identify exactly one row for the current Unit.",
    EnabledUnit = Table.SelectRows(ValidatedUnit, each [EffortManagementAnalysis] = true),
    Result = Table.RemoveColumns(EnabledUnit, {"LeaveBalanceAnalysis"})
in
    Result;

// Query: IMPORT Availability Leave Source
// Purpose: Read dated availability and leave for the unit selected by this workbook's folder.
// Inputs: Combined Output sheet in 1. Input/whiddon_availability_leave_extraction.xlsx.
shared #"IMPORT Availability Leave Source" = let
    SourcePath = #"FilePath - 1Input" & "\whiddon_availability_leave_extraction.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    CombinedOutputSheet = Navigation{[Item = "Combined Output", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(CombinedOutputSheet, [PromoteAllScalars = true]),
    // Use the same source-title to Facility-Abbrev mapping as the unit extraction.
    JoinedFacilities = Table.NestedJoin(PromotedHeaders, {"Site Name"}, #"LINK Facilities", {"Title"}, "FacilityLookup", JoinKind.LeftOuter),
    ExpandedFacilities = Table.ExpandTableColumn(JoinedFacilities, "FacilityLookup", {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    FilteredUnit = Table.SelectRows(ExpandedFacilities, each [#"Facility-Abbrev"] = Unit and AnalysisEnabled),
    SelectedColumns = Table.SelectColumns(FilteredUnit,
        {"Site Name", "Employee Name", "Payroll Code", "Department", "Date From", "Date To", "Availability or Leave Reason", "Facility-Abbrev"}),
    NormalizedIdentity = Table.TransformColumns(SelectedColumns,
        {{"Payroll Code", AvailabilityText, type nullable text}, {"Employee Name", AvailabilityText, type nullable text}}),
    IndexedRecords = Table.AddIndexColumn(NormalizedIdentity, "SourceRecord", 1, 1, Int64.Type),
    // Buffer the reduced scalar source shared by name resolution, parsing and diagnostics.
    BufferedRecords = Table.Buffer(IndexedRecords)
in
    BufferedRecords;

// Query: WorkerContracts_Prepare
// Purpose: Resolve Worker's Report to one validated fortnightly contract value per employee.
// Notes: Identical duplicate rows are allowed; blank versus populated, distinct values, non-numeric values, negative values and non-finite values are invalid.
shared WorkerContracts_Prepare = let
    NormalizedEmployee = Table.TransformColumns(#"IMPORT Worker's Report",
        {{"Employee Code", AvailabilityText, type nullable text}}),
    ParseContract = (row as record) as record =>
        let
            RawAttempt = try row[Contracted FN Hours],
            HasRawError = RawAttempt[HasError],
            RawValue = if RawAttempt[HasError] then null else RawAttempt[Value],
            NormalizedRawValue = if HasRawError then null else AvailabilityText(RawValue),
            IsBlank = not HasRawError and NormalizedRawValue = null,
            ShouldParse = not HasRawError and not IsBlank,
            NumberAttempt = if ShouldParse then try Number.From(RawValue) else null,
            ParseError = if HasRawError then true
                else if IsBlank then false
                else NumberAttempt[HasError],
            ContractValue = if ParseError or IsBlank then null else NumberAttempt[Value],
            Result = [ContractValue = ContractValue, IsBlank = IsBlank, ParseError = ParseError]
        in
            Result,
    ParsedContract = Table.AddColumn(NormalizedEmployee, "ContractParse", ParseContract,
        type [ContractValue = nullable number, IsBlank = logical, ParseError = logical]),
    ExpandedContract = Table.ExpandRecordColumn(ParsedContract, "ContractParse",
        {"ContractValue", "IsBlank", "ParseError"}),
    // Rows without an employee code cannot participate in an employee contract join.
    IdentifiedRows = Table.SelectRows(ExpandedContract, each [Employee Code] <> null),
    RenamedEmployee = Table.RenameColumns(IdentifiedRows, {{"Employee Code", "EmployeeID"}}),
    DistinctContractValues = (rows as table) as list =>
        let
            NonNullValues = List.RemoveNulls(rows[ContractValue]),
            DistinctValues = List.Distinct(NonNullValues)
        in
            DistinctValues,
    CountTrueValues = (values as list) as number =>
        let
            TrueValues = List.Select(values, each _ = true),
            Count = List.Count(TrueValues)
        in
            Count,
    Grouped = Table.Group(RenamedEmployee, {"EmployeeID"},
        {{"ContractValues", DistinctContractValues, type list},
         {"BlankRows", each CountTrueValues([IsBlank]), Int64.Type},
         {"ParseErrors", each CountTrueValues([ParseError]), Int64.Type}}),
    ResolveContract = (row as record) as record =>
        let
            Values = row[ContractValues],
            ValueCount = List.Count(Values),
            NonFiniteChecks = List.Transform(Values, each Number.IsNaN(_) or Number.Abs(_) = #infinity),
            HasNonFinite = List.AnyTrue(NonFiniteChecks),
            NegativeChecks = List.Transform(Values, each _ < 0),
            HasNegative = List.AnyTrue(NegativeChecks),
            HasMultipleValues = ValueCount > 1,
            MixesBlankAndValue = row[BlankRows] > 0 and ValueCount > 0,
            HasConflict = HasMultipleValues or MixesBlankAndValue,
            Issue = if row[ParseErrors] > 0 then "Non-numeric Contracted FN Hours"
                else if HasNonFinite then "Non-finite Contracted FN Hours"
                else if HasNegative then "Negative Contracted FN Hours"
                else if HasConflict then "Conflicting Contracted FN Hours"
                else null,
            ContractHours = if ValueCount = 1 then Values{0} else null,
            Result = [Contracted FN Hours = ContractHours, Issue = Issue]
        in
            Result,
    Resolved = Table.AddColumn(Grouped, "ContractResolution", ResolveContract,
        type [Contracted FN Hours = nullable number, Issue = nullable text]),
    ExpandedResolution = Table.ExpandRecordColumn(Resolved, "ContractResolution",
        {"Contracted FN Hours", "Issue"}),
    Output = Table.SelectColumns(ExpandedResolution, {"EmployeeID", "Contracted FN Hours", "Issue"}),
    BufferedOutput = Table.Buffer(Output)
in
    BufferedOutput;

// Query: IMPORT Availability Settings
// Purpose: Share the Settings workbook navigation among the shift, calendar, allowance and maximum-availability extracts.
shared #"IMPORT Availability Settings" =
    let
        SourcePath = #"FilePath - 2Calculations" & "\Settings Data.xlsx",
        SourceBinary = File.Contents(SourcePath),
        Navigation = Excel.Workbook(SourceBinary, null, true),
        IsRequiredObject = (row as record) as logical =>
            let
                IsScalarSetting = List.Contains({"ShiftDuration", "MaxAvailability"}, row[Item]),
                IsTableOrName = List.Contains({"Table", "DefinedName"}, row[Kind]),
                IsTabularSetting = List.Contains({"ShiftPeriod", "PermutationDimensions"}, row[Item]),
                IsTable = row[Kind] = "Table",
                IsRequired = (IsScalarSetting and IsTableOrName) or (IsTabularSetting and IsTable)
            in
                IsRequired,
        Required = Table.SelectRows(Navigation, IsRequiredObject),
        // Buffer only required data tables; navigation buffering alone is shallow.
        BufferedData = Table.TransformColumns(Required, {{"Data", Table.Buffer, type table}}),
        BufferedNavigation = Table.Buffer(BufferedData)
    in
        BufferedNavigation;

// Query: EXTRACT EffectiveShiftHrs
// Purpose: Import the full-shift effective-hours allowance; never derive it from roster-day caps.
// Inputs: ShiftDuration table or named range, with one ShiftDuration value.
shared #"EXTRACT EffectiveShiftHrs" = let
    Matches = Table.SelectRows(#"IMPORT Availability Settings",
        each [Item] = "ShiftDuration" and List.Contains({"Table", "DefinedName"}, [Kind])),
    MatchCount = Table.RowCount(Matches),
    Data = if MatchCount = 1 then Matches{0}[Data]
        else error "Expected exactly one ShiftDuration table or named range in Settings Data.",
    HasExpectedHeader = Table.HasColumns(Data, "ShiftDuration"),
    WithHeaders = if HasExpectedHeader then Data
        else Table.PromoteHeaders(Data, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "ShiftDuration"),
    ValueCount = List.Count(Values),
    Allowance = if ValueCount = 1 then Number.From(Values{0})
        else error "ShiftDuration must contain exactly one allowance.",
    IsNotANumber = if Allowance = null then false else Number.IsNaN(Allowance),
    IsNotPositive = if Allowance = null then false else Allowance <= 0,
    IsInfinite = if Allowance = null then false else Allowance = #infinity,
    Validated = if Allowance = null then error "ShiftDuration is blank."
        else if IsNotANumber or IsNotPositive or IsInfinite then
            error "ShiftDuration must be a positive finite number."
        else Allowance
in Validated;

// Query: EXTRACT MaxAvailability
// Purpose: Validate the Settings maximum number of shifts available to one resource for the roster.
// Inputs: MaxAvailability table or named range, containing exactly one positive whole number.
shared #"EXTRACT MaxAvailability" = let
    Matches = Table.SelectRows(#"IMPORT Availability Settings",
        each [Item] = "MaxAvailability" and List.Contains({"Table", "DefinedName"}, [Kind])),
    MatchCount = Table.RowCount(Matches),
    Data = if MatchCount = 1 then Matches{0}[Data]
        else error Error.Record("Maximum availability settings", "Expected exactly one MaxAvailability table or defined name in Settings Data.", [Matches = MatchCount]),
    HasExpectedHeader = Table.HasColumns(Data, "MaxAvailability"),
    WithHeaders = if HasExpectedHeader then Data
        else Table.PromoteHeaders(Data, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "MaxAvailability"),
    ValueCount = List.Count(Values),
    Maximum = if ValueCount = 1 then Number.From(Values{0})
        else error "MaxAvailability must contain exactly one value.",
    IsNotANumber = if Maximum = null then false else Number.IsNaN(Maximum),
    AbsoluteMaximum = if Maximum = null then null else Number.Abs(Maximum),
    IsInfinite = AbsoluteMaximum = #infinity,
    IsNotPositive = if Maximum = null then false else Maximum <= 0,
    RoundedDown = if Maximum = null then null else Number.RoundDown(Maximum),
    IsNotWhole = if Maximum = null then false else Maximum <> RoundedDown,
    IsInvalid = Maximum = null or IsNotANumber or IsInfinite or IsNotPositive or IsNotWhole,
    Validated = if IsInvalid then
            error "MaxAvailability must be one positive whole number."
        else Int64.From(Maximum)
in Validated;

// Query: IMPORT Role Lists
// Purpose: Read the Whiddon SharePoint list navigation used by the approved LINK Roles query.
// Notes: Retain the site, implementation and list ID used by the TE 1-AllocationExtracted source.
shared #"IMPORT Role Lists" =
let
    ListNavigation = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI",
        [Implementation="2.0", ViewMode="All"]),
    BufferedLists = Table.Buffer(ListNavigation)
in
    BufferedLists;

// Query: LINK Roles
// Purpose: Read the authoritative roster-role to Role Group assignments used by TE AllocationExtraction.
// Output: The five approved role-list fields, including Roster Roles and Role Group.
shared #"LINK Roles" =
let
    ListNavigation = #"IMPORT Role Lists",
    RoleDefinitions = ListNavigation{[Id="6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items],
    RoleColumns = Table.SelectColumns(RoleDefinitions,
        {"Roster Roles", "DC Category", "DC Role", "Direct Care %", "Role Group"}),
    BufferedRoles = Table.Buffer(RoleColumns)
in
    BufferedRoles;

// Query: WorkerRoleKey
// Purpose: Match worker roster roles to LINK Roles without case or surrounding-space differences.
shared WorkerRoleKey = (roleName as any) as nullable text =>
    let
        CleanRole = AvailabilityText(roleName),
        RoleKey = if CleanRole = null then null else Text.Upper(CleanRole)
    in
        RoleKey;

// Query: AvailabilityLeaveReasons
// Purpose: List exact recognised leave descriptions; new descriptions require an explicit mapping.
// Notes: Includes source descriptions confirmed in the 197-record diagnosis. Matching ignores surrounding spaces and letter case.
shared AvailabilityLeaveReasons = #table(type table [Reason = text], {
    {"Annual Rec Lve"},
    {"Maternity Lve Unpaid"},
    {"Unpaid Leave"},
    {"Study Leave Unpaid"},
    {"Workers Comp Not Worked Sec 40"},
    {"Personal Sick Leave Unpaid"},
    {"Personal Sick Leave Paid"},
    {"WC Sec 37 13-130wks unfit"},
    {"WC Sec 36 3-13 wks unfit"},
    {"Personal Carer Leave Paid"},
    {"Personal Carer Leave Unpaid"},
    {"PH Worked Leave - Taken"},
    {"Long Service Lve"},
    {"Compassionate Lve Paid"},
    {"Personal Emergency Leave (PEL)"}
});

// Query: AvailabilityIntervalHours
// Purpose: Sum clock hours in an already disjoint interval list.
shared AvailabilityIntervalHours = (intervals as list) as number =>
    let
        Durations = List.Transform(intervals, each [End] - [Start]),
        Hours = List.Transform(Durations, Duration.TotalHours),
        TotalHours = List.Accumulate(Hours, 0, (total, hours) => total + hours)
    in
        TotalHours;

// Query: AvailabilityClipIntervals
// Purpose: Retain half-open intersections with one shift or calendar range.
shared AvailabilityClipIntervals = (intervals as list, start as datetime, end as datetime) as list =>
    let
        Overlapping = List.Select(intervals, each [Start] < end and [End] > start),
        Clipped = List.Transform(Overlapping, each
            [Start = List.Max({[Start], start}), End = List.Min({[End], end})]),
        Buffered = List.Buffer(Clipped)
    in
        Buffered;

// Query: AvailabilityShiftMeasurement
// Purpose: Measure a single shift and check interval invariants while its small lists are in scope.
shared AvailabilityShiftMeasurement = (baseline as list, unavail as list, leave as list,
    shiftStart as datetime, shiftEnd as datetime, fullDayLeave as logical) as record =>
    let
        ClippedBaseline = AvailabilityClipIntervals(baseline, shiftStart, shiftEnd),
        ClippedUnavailability = AvailabilityClipIntervals(unavail, shiftStart, shiftEnd),
        ClippedLeave = AvailabilityClipIntervals(leave, shiftStart, shiftEnd),
        // UNION across both categories prevents overlapping leave and UNAVAIL being deducted twice.
        CombinedCuts = ClippedUnavailability & ClippedLeave,
        UnionedCuts = AvailabilityUnion(CombinedCuts),
        BufferedCuts = List.Buffer(UnionedCuts),
        RemainingSegments = AvailabilitySubtract(ClippedBaseline, BufferedCuts),
        LeaveAdjustedSegments = if fullDayLeave then {} else RemainingSegments,
        BufferedSegments = List.Buffer(LeaveAdjustedSegments),
        SegmentCount = List.Count(BufferedSegments),
        SingleSegment = SegmentCount = 1,
        StartsAtShiftStart = if SingleSegment then BufferedSegments{0}[Start] = shiftStart else false,
        EndsAtShiftEnd = if SingleSegment then BufferedSegments{0}[End] = shiftEnd else false,
        FullShift = SingleSegment and StartsAtShiftStart and EndsAtShiftEnd,
        PositiveDurationChecks = List.Transform(BufferedSegments,
            each [Start] < [End]),
        InsideShiftChecks = List.Transform(BufferedSegments, (segment) =>
            segment[Start] >= shiftStart and segment[End] <= shiftEnd),
        AvoidsCutChecks = List.Transform(BufferedSegments, (segment) =>
            let
                CutOverlapChecks = List.Transform(BufferedCuts, (cut) =>
                    segment[Start] < cut[End] and segment[End] > cut[Start]),
                HasCutOverlap = List.AnyTrue(CutOverlapChecks),
                AvoidsCuts = not HasCutOverlap
            in
                AvoidsCuts),
        InsideBaselineChecks = List.Transform(BufferedSegments, (segment) =>
            let
                BaselineContainmentChecks = List.Transform(ClippedBaseline, (window) =>
                    segment[Start] >= window[Start] and segment[End] <= window[End]),
                IsInsideBaseline = List.AnyTrue(BaselineContainmentChecks)
            in
                IsInsideBaseline),
        AllDurationsPositive = List.AllTrue(PositiveDurationChecks),
        AllSegmentsInsideShift = List.AllTrue(InsideShiftChecks),
        AllSegmentsAvoidCuts = List.AllTrue(AvoidsCutChecks),
        AllSegmentsInsideBaseline = List.AllTrue(InsideBaselineChecks),
        SegmentsValid = AllDurationsPositive and AllSegmentsInsideShift
            and AllSegmentsAvoidCuts and AllSegmentsInsideBaseline,
        SegmentEndTimes = List.Transform(BufferedSegments, each [End]),
        SegmentStartTimes = List.Transform(BufferedSegments, each [Start]),
        PreviousSegmentEnds = if SegmentCount < 2 then {} else List.RemoveLastN(SegmentEndTimes, 1),
        NextSegmentStarts = if SegmentCount < 2 then {} else List.Skip(SegmentStartTimes, 1),
        AdjacentEndpoints = List.Zip({PreviousSegmentEnds, NextSegmentStarts}),
        NoOverlapChecks = List.Transform(AdjacentEndpoints, each _{0} <= _{1}),
        Disjoint = List.AllTrue(NoOverlapChecks),
        BaselineHours = AvailabilityIntervalHours(ClippedBaseline),
        LeaveHours = AvailabilityIntervalHours(ClippedLeave),
        UnavailabilityHours = AvailabilityIntervalHours(ClippedUnavailability),
        CombinedAbsenceHours = AvailabilityIntervalHours(BufferedCuts),
        AvailableHours = AvailabilityIntervalHours(BufferedSegments),
        Measurement = [
            BaselineHours = BaselineHours,
            LeaveHoursInShift = LeaveHours,
            UnavailHoursInShift = UnavailabilityHours,
            CombinedAbsenceHoursInShift = CombinedAbsenceHours,
            AvailableHours = AvailableHours,
            FullShift = FullShift,
            SegmentsValid = SegmentsValid,
            Disjoint = Disjoint
        ]
    in
        Measurement;

// Query: AvailabilityResidualHours
// Purpose: Apply the approved allowance-minus-absence rule, capped by residual clock availability.
// Notes: A complete shift receives the allowance even when its clock length differs; no break/minimum rule.
shared AvailabilityResidualHours = (remainingHours as number, absenceHours as number,
    fullShift as logical, fullDayLeave as logical, allowance as number) as number =>
    let
        IsBlocked = fullDayLeave or remainingHours <= 0,
        ResidualAllowance = allowance - absenceHours,
        CappedAtRemainingHours = List.Min({remainingHours, ResidualAllowance}),
        NonNegativeResidual = List.Max({0, CappedAtRemainingHours}),
        EffectiveHours = if IsBlocked then 0
            else if fullShift then allowance
            else NonNegativeResidual
    in
        EffectiveHours;

// Query: AvailabilityRecordKind
// Purpose: Separate AVAIL, UNAVAIL and recognised leave without interpreting unknown reasons as leave.
shared AvailabilityRecordKind = let
    LeaveReasons = AvailabilityLeaveReasons[Reason],
    TrimmedLeaveReasons = List.Transform(LeaveReasons, Text.Trim),
    UppercaseLeaveReasons = List.Transform(TrimmedLeaveReasons, Text.Upper),
    LeaveNames = List.Buffer(UppercaseLeaveReasons),
    ClassifyReason = (reason as any) as text =>
        let
            Clean = AvailabilityText(reason),
            ContainsUnavailability = if Clean = null then false else Text.Contains(Clean, "UNAVAIL"),
            ContainsAvailability = if Clean = null then false else Text.Contains(Clean, "AVAIL"),
            UppercaseReason = if Clean = null then null else Text.Upper(Clean),
            IsRecognisedLeave = if UppercaseReason = null then false else List.Contains(LeaveNames, UppercaseReason),
            // Preserve case-sensitive source AVAIL/UNAVAIL rules, with UNAVAIL taking precedence.
            RecordKind = if Clean = null then "Unknown"
                else if ContainsUnavailability then "UNAVAIL"
                else if ContainsAvailability then "AVAIL"
                else if IsRecognisedLeave then "Leave"
                else "Unknown"
        in
            RecordKind
in
    ClassifyReason;

// Query: AvailabilityDailyWindows
// Purpose: Reset the availability baseline on each calendar day before applying exclusions.
// Inputs: Valid AVAIL intervals and complete calendar boundaries, with an exclusive midnight end.
// Output: Unioned baseline intervals; a day without AVAIL is fully available.
// Notes: An interval spanning midnight affects both dates only for its recorded duration.
shared AvailabilityDailyWindows = (intervals as list, calendarStart as datetime, calendarEnd as datetime) as list =>
    let
        // Reuse the worker's AVAIL list while checking each date in the roster horizon.
        Available = List.Buffer(intervals),
        Days = List.Generate(() => calendarStart, each _ < calendarEnd,
            each _ + #duration(1, 0, 0, 0)),
        DailyWindows = List.Transform(Days, (dayStart) =>
            let
                DayEnd = dayStart + #duration(1, 0, 0, 0),
                Overlapping = List.Select(Available,
                    each [Start] < DayEnd and [End] > dayStart),
                Clipped = List.Transform(Overlapping, each [
                    Start = List.Max({[Start], dayStart}),
                    End = List.Min({[End], DayEnd})
                ]),
                // Only AVAIL on this date restricts this date. Tomorrow is evaluated afresh.
                Baseline = if List.IsEmpty(Overlapping) then {[Start = dayStart, End = DayEnd]}
                    else Clipped
            in Baseline),
        CombinedWindows = List.Combine(DailyWindows),
        // Merge touching windows across midnight so a fully available night remains one segment.
        UnionedWindows = AvailabilityUnion(CombinedWindows),
        Result = List.Buffer(UnionedWindows)
    in
        Result;

// Query: AvailabilityLeaveDayTotals
// Purpose: Split and union recognised leave by calendar date, never including UNAVAIL.
// Inputs: One worker/facility's recognised leave intervals only.
// Output: Date, raw record-hour sum, distinct leave hours and source-piece count.
shared AvailabilityLeaveDayTotals = (intervals as list) as table =>
    let
        SplitIntervalByDay = (interval as record) as list =>
            let
                FirstDate = Date.From(interval[Start]),
                FirstDayStart = DateTime.From(FirstDate),
                DayStarts = List.Generate(
                    () => FirstDayStart,
                    each _ < interval[End],
                    each _ + #duration(1, 0, 0, 0)
                ),
                Pieces = List.Transform(DayStarts, (dayStart) =>
                    let
                        DayEnd = dayStart + #duration(1, 0, 0, 0),
                        PieceDate = Date.From(dayStart),
                        PieceStart = List.Max({interval[Start], dayStart}),
                        PieceEnd = List.Min({interval[End], DayEnd}),
                        Piece = [Date = PieceDate, Start = PieceStart, End = PieceEnd]
                    in
                        Piece)
            in
                Pieces,
        PiecesByInterval = List.Transform(intervals, SplitIntervalByDay),
        Pieces = List.Combine(PiecesByInterval),
        Rows = Table.FromRecords(Pieces, type table [Date = date, Start = datetime, End = datetime]),
        SumRecordHours = (dayRows as table) as number =>
            let
                Records = Table.ToRecords(dayRows),
                Durations = List.Transform(Records, each [End] - [Start]),
                Hours = List.Transform(Durations, Duration.TotalHours),
                Total = List.Sum(Hours)
            in
                Total,
        SumDistinctHours = (dayRows as table) as number =>
            let
                SelectedIntervals = Table.SelectColumns(dayRows, {"Start", "End"}),
                IntervalRecords = Table.ToRecords(SelectedIntervals),
                UnionedIntervals = AvailabilityUnion(IntervalRecords),
                Total = AvailabilityIntervalHours(UnionedIntervals)
            in
                Total,
        Days = Table.Group(Rows, {"Date"}, {
            {"SumOfLeaveRecordHours", SumRecordHours, type number},
            {"DistinctLeaveHours", SumDistinctHours, type number},
            {"LeaveRecordCount", each Table.RowCount(_), Int64.Type}})
    in
        Days;

// Query: AvailabilityRecords
// Purpose: Parse source intervals and retain separate record kinds and traceable validation issues.
// Notes: Available remains a compatibility flag; only RecordKind = Leave contributes to daily leave hours.
shared AvailabilityRecords = let
    ParseAvailabilityRecord = (row as record) as record =>
        let
            StartAttempt = try DateTime.From(row[Date From]),
            EndAttempt = try DateTime.From(row[Date To]),
            Start = if StartAttempt[HasError] then null else StartAttempt[Value],
            End = if EndAttempt[HasError] then null else EndAttempt[Value],
            Reason = AvailabilityText(row[Availability or Leave Reason]),
            Kind = AvailabilityRecordKind(Reason),
            ReasonDescription = if Reason = null then "(blank)" else Reason,
            IsAvailable = if Kind = "Unknown" then null else Kind = "AVAIL",
            Issue = if row[Payroll Code] = null then "Missing or invalid payroll code"
                else if Kind = "Unknown" then "Unrecognised or missing reason: " & ReasonDescription
                else if Start = null or End = null then "Missing or invalid datetime"
                else if End <= Start then "End must be later than start"
                else null,
            Result = [
                Start = Start,
                End = End,
                RecordKind = Kind,
                OriginalReason = Reason,
                Available = IsAvailable,
                Issue = Issue
            ]
        in
            Result,
    Parsed = Table.AddColumn(#"IMPORT Availability Leave Source", "Parsed", ParseAvailabilityRecord),
    Expanded = Table.ExpandRecordColumn(Parsed, "Parsed", {"Start", "End", "RecordKind", "OriginalReason", "Available", "Issue"}),
    // Buffer only scalar fields reused by identity-independent diagnostics and eligible interval staging.
    SelectedScalars = Table.SelectColumns(Expanded,
        {"Payroll Code", "Facility-Abbrev", "SourceRecord", "OriginalReason", "Start", "End", "RecordKind", "Available", "Issue"}),
    BufferedScalars = Table.Buffer(SelectedScalars)
in
    BufferedScalars;

// Query: ReconciledWorkers_Prepare
// Purpose: Assign LINK Roles Role Groups and resolve worker names while retaining explicit eligibility diagnostics.
// Output: One row per reconciled employee/facility/source-role, including excluded rows.
// Notes: Worker Reconciliation Role joins to LINK Roles Roster Roles; Role Group becomes the downstream Role.
shared ReconciledWorkers_Prepare = let
    DistinctNonNullValues = (values as list) as list =>
        let
            NonNullValues = List.RemoveNulls(values),
            DistinctValues = List.Distinct(NonNullValues)
        in
            DistinctValues,
    SourceNames = Table.Group(#"IMPORT Availability Leave Source", {"Payroll Code", "Facility-Abbrev"},
        {{"SourceNames", each DistinctNonNullValues([Employee Name]), type list}}),
    CombineEmploymentTypes = (values as list) as text =>
        let
            DistinctValues = DistinctNonNullValues(values),
            CombinedValues = Text.Combine(DistinctValues, "; ")
        in
            CombinedValues,
    CombineSourceRoles = (values as list) as nullable text =>
        let
            DistinctValues = DistinctNonNullValues(values),
            SortedValues = List.Sort(DistinctValues),
            CombinedValues = if List.IsEmpty(SortedValues) then null
                else Text.Combine(SortedValues, "; ")
        in
            CombinedValues,
    SourceRoleKeys = Table.AddColumn(#"IMPORT Reconciled Workers", "SourceRoleKey",
        each WorkerRoleKey([Role]), type nullable text),
    // Group on the normalized roster-role key before mapping so case or outer-space variants do not create duplicate workers.
    WorkerGroups = Table.Group(SourceRoleKeys,
        {"EmployeeID", "Facility-Abbrev", "SourceRoleKey"},
        {{"SourceRoles", each DistinctNonNullValues([Role]), type list},
         {"RosterNames", each DistinctNonNullValues([Employee Roster Name]), type list},
         {"EmploymentType", each CombineEmploymentTypes([EmploymentType]), type text}}),
    SourceRoleLabels = Table.AddColumn(WorkerGroups, "SourceRole",
        each CombineSourceRoles([SourceRoles]), type nullable text),
    RemovedSourceRoleLists = Table.RemoveColumns(SourceRoleLabels, {"SourceRoles"}),
    // Join the worker's normalized source role to the authoritative LINK Roles Roster Roles mapping.
    RoleMappingLookup = Table.SelectRows(WorkerRoleMappings_Prepare,
        each [RosterRoleKey] <> null),
    JoinedRoleMappings = Table.NestedJoin(RemovedSourceRoleLists, {"SourceRoleKey"},
        RoleMappingLookup, {"RosterRoleKey"}, "RoleMapping", JoinKind.LeftOuter),
    CountedRoleMappings = Table.AddColumn(JoinedRoleMappings, "RoleMappingMatches",
        each Table.RowCount([RoleMapping]), Int64.Type),
    ExpandedRoleMappings = Table.ExpandTableColumn(CountedRoleMappings, "RoleMapping", {
        "MappedRosterRole", "RoleMappingDefinitionRows", "DistinctRoleGroupCount",
        "MappedRoleGroup", "RoleMappingIssue"
    }, {
        "MappedRosterRole", "RoleMappingDefinitionRows", "DistinctRoleGroupCount",
        "MappedRoleGroup", "ConfiguredRoleMappingIssue"
    }),
    FilledRoleMappingCounts = Table.ReplaceValue(ExpandedRoleMappings, null, 0,
        Replacer.ReplaceValue, {"RoleMappingDefinitionRows", "DistinctRoleGroupCount"}),
    ResolvedRoleMappingIssues = Table.AddColumn(FilledRoleMappingCounts, "RoleMappingIssue", each
        if [SourceRoleKey] = null then null
        else if [RoleMappingMatches] = 0 then "Roster role absent from LINK Roles"
        else [ConfiguredRoleMappingIssue], type nullable text),
    MappedRole = Table.AddColumn(ResolvedRoleMappingIssues, "Role", each
        if [RoleMappingIssue] = null then [MappedRoleGroup] else null, type nullable text),
    // Contract hours belong to the employee, not to an employee-period or role split.
    JoinedContracts = Table.NestedJoin(MappedRole, {"EmployeeID"},
        WorkerContracts_Prepare, {"EmployeeID"}, "Contract", JoinKind.LeftOuter),
    ExpandedContracts = Table.ExpandTableColumn(JoinedContracts, "Contract",
        {"Contracted FN Hours", "Issue"}, {"Contracted FN Hours", "ContractIssue"}),
    JoinedNames = Table.NestedJoin(ExpandedContracts, {"EmployeeID", "Facility-Abbrev"},
        SourceNames, {"Payroll Code", "Facility-Abbrev"}, "Names", JoinKind.LeftOuter),
    ResolveNameCandidates = (row as record) as list =>
        let
            HasRosterNames = not List.IsEmpty(row[RosterNames]),
            HasSourceNames = not Table.IsEmpty(row[Names]),
            SourceNameCandidates = if HasSourceNames then row[Names]{0}[SourceNames] else {},
            Candidates = if HasRosterNames then row[RosterNames] else SourceNameCandidates
        in
            Candidates,
    ResolvedNames = Table.AddColumn(JoinedNames, "NameCandidates", ResolveNameCandidates, type list),
    ResolveBaseName = (names as list) as nullable text =>
        let
            NameCount = List.Count(names),
            BaseName = if NameCount = 1 then names{0} else null
        in
            BaseName,
    NamedWorkers = Table.AddColumn(ResolvedNames, "BaseName", each ResolveBaseName([NameCandidates]), type nullable text),
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    DetermineEligibilityIssue = (row as record) as nullable text =>
        let
            Issue = if row[EmployeeID] = null then "Missing employee ID"
                else if row[#"Facility-Abbrev"] = null then "Missing facility"
                else if row[#"Facility-Abbrev"] <> Unit or not AnalysisEnabled then "Outside unit analysis scope"
                else if row[SourceRoleKey] = null then "Missing role"
                else if row[RoleMappingIssue] <> null then row[RoleMappingIssue]
                else if row[Role] = null then "LINK Roles did not assign a Role Group"
                else if row[ContractIssue] <> null then row[ContractIssue]
                else if row[BaseName] = null then "Missing or ambiguous worker name"
                else null
        in
            Issue,
    Eligibility = Table.AddColumn(NamedWorkers, "Issue", DetermineEligibilityIssue, type nullable text),
    Result = Table.RemoveColumns(Eligibility,
        {"Names", "RosterNames", "NameCandidates", "ConfiguredRoleMappingIssue"}),
    // Nested name lists have been removed; reuse this scalar preparation for eligibility and diagnostics.
    BufferedResult = Table.Buffer(Result)
in
    BufferedResult;

// Query: AllocationRoleHours_Prepare
// Purpose: Aggregate approved-role allocation hours by employee, mapped facility and role.
// Notes: Facility-Abbrev is the facility identity; raw Locations are retained only as source provenance.
shared AllocationRoleHours_Prepare = let
    ApprovedRoleNames = {"AIN", "EN", "RN"},
    ApprovedRoles = Table.SelectRows(#"IMPORT AllocationExtracted", each List.Contains(ApprovedRoleNames, [Role])),
    JoinedFacilities = Table.NestedJoin(ApprovedRoles, {"Location"}, #"LINK Facilities",
        {"Title"}, "FacilityLookup", JoinKind.LeftOuter),
    CountedMappings = Table.AddColumn(JoinedFacilities, "FacilityMatches", each
        Table.RowCount([FacilityLookup]), Int64.Type),
    InvalidMappings = Table.SelectRows(CountedMappings, each [FacilityMatches] <> 1),
    ValidatedMappings = if Table.IsEmpty(InvalidMappings) then CountedMappings
        else error Error.Record("Allocation role facility mapping",
            "Each allocation Location must match exactly one Facilities List row.",
            Table.SelectColumns(InvalidMappings, {"Location", "Code", "Role", "FacilityMatches"})),
    ExpandedFacilities = Table.ExpandTableColumn(ValidatedMappings, "FacilityLookup",
        {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    UnitRows = Table.SelectRows(ExpandedFacilities, each [#"Facility-Abbrev"] = Unit and AnalysisEnabled),
    ParseAllocatedHours = (value as any) as nullable number =>
        let
            Attempt = try Number.From(value),
            ParsedValue = if Attempt[HasError] then null else Attempt[Value]
        in
            ParsedValue,
    ParsedHours = Table.AddColumn(UnitRows, "AllocatedHours", each ParseAllocatedHours([Hours]), type nullable number),
    HasInvalidAllocation = (row as record) as logical =>
        let
            AllocatedHours = row[AllocatedHours],
            IsMissing = AllocatedHours = null,
            IsNegative = if IsMissing then false else AllocatedHours < 0,
            IsNotANumber = if IsMissing then false else Number.IsNaN(AllocatedHours),
            IsInfinite = if IsMissing then false else Number.Abs(AllocatedHours) = #infinity,
            IsInvalid = row[Code] = null or row[Name] = null or IsMissing
                or IsNegative or IsNotANumber or IsInfinite
        in
            IsInvalid,
    InvalidRows = Table.SelectRows(ParsedHours, HasInvalidAllocation),
    InvalidRowDetails = Table.SelectColumns(InvalidRows, {"Location", "Code", "Role", "Name", "Hours"}),
    ValidatedRows = if Table.IsEmpty(InvalidRows) then ParsedHours
        else error Error.Record("Allocation role hours",
            "Approved-role allocation rows need employee IDs, names and finite nonnegative hours.",
            InvalidRowDetails),
    DistinctNames = (rows as table) as list =>
        let
            Values = rows[Name],
            DistinctValues = List.Distinct(Values)
        in
            DistinctValues,
    SortedSourceFacilities = (rows as table) as list =>
        let
            Values = rows[Location],
            DistinctValues = List.Distinct(Values),
            SortedValues = List.Sort(DistinctValues)
        in
            SortedValues,
    GroupedHours = Table.Group(ValidatedRows, {"Code", "Facility-Abbrev", "Role"},
        {{"AllocatedHours", each List.Sum([AllocatedHours]), type number},
         {"RosterNames", DistinctNames, type list},
         {"SourceFacilities", SortedSourceFacilities, type list}}),
    ResolveRosterName = (row as record) as text =>
        let
            NameCount = List.Count(row[RosterNames]),
            RosterName = if NameCount = 1 then row[RosterNames]{0}
        else error Error.Record("Allocation role name", "An employee/facility/role has conflicting roster names.",
            [EmployeeID = row[Code], Facility = row[#"Facility-Abbrev"], Role = row[Role],
             SourceFacilities = row[SourceFacilities], Names = row[RosterNames]])
        in
            RosterName,
    ResolvedNames = Table.AddColumn(GroupedHours, "RosterName", ResolveRosterName, type text),
    NamedKeys = Table.RenameColumns(ResolvedNames, {{"Code", "EmployeeID"}}),
    Result = Table.SelectColumns(NamedKeys,
        {"EmployeeID", "Facility-Abbrev", "Role", "RosterName", "AllocatedHours", "SourceFacilities"}),
    BufferedResult = Table.Buffer(Result)
in
    BufferedResult;

// Query: AllocationRoleShares
// Purpose: Calculate unrounded AIN/EN/RN proportions using all approved-role hours per employee and mapped facility.
// Notes: The same employee ID may occur at multiple facilities; shares are calculated separately by Facility-Abbrev.
shared AllocationRoleShares = let
    Source = AllocationRoleHours_Prepare,
    WorkerTotals = Table.Group(Source, {"EmployeeID", "Facility-Abbrev"},
        {{"TotalHours", each List.Sum([AllocatedHours]), type number},
         {"RoleCount", each
            let
                DistinctRoles = List.Distinct([Role]),
                RoleCount = List.Count(DistinctRoles)
            in
                RoleCount, Int64.Type}}),
    JoinedTotals = Table.NestedJoin(Source,
        {"EmployeeID", "Facility-Abbrev"}, WorkerTotals,
        {"EmployeeID", "Facility-Abbrev"}, "WorkerTotal", JoinKind.LeftOuter),
    ExpandedTotals = Table.ExpandTableColumn(JoinedTotals, "WorkerTotal",
        {"TotalHours", "RoleCount"}, {"TotalHours", "RoleCount"}),
    InvalidTotals = Table.SelectRows(ExpandedTotals, each [TotalHours] <= 0),
    ValidatedTotals = if Table.IsEmpty(InvalidTotals) then ExpandedTotals
        else error "Approved-role allocation hours must total more than zero per employee and mapped facility.",
    Shares = Table.AddColumn(ValidatedTotals, "RoleShare", each
        [AllocatedHours] / [TotalHours], type number),
    BufferedShares = Table.Buffer(Shares)
in
    BufferedShares;

// Query: ReconciledWorkers_Eligible
// Purpose: Publish the eligible unit identities with the existing multiple-role display-name convention.
// Notes: Membership is independent of available hours, so wholly unavailable workers remain identifiable.
shared ReconciledWorkers_Eligible = let
    Included = Table.SelectRows(ReconciledWorkers_Prepare, each [Issue] = null),
    SelectedWorkers = Table.SelectColumns(Included,
        {"EmployeeID", "Facility-Abbrev", "Role", "BaseName", "EmploymentType", "Contracted FN Hours"}),
    UniqueWorkers = Table.Distinct(SelectedWorkers),
    RoleCounts = Table.Group(UniqueWorkers, {"EmployeeID", "Facility-Abbrev"},
        {{"RoleCount", each
            let
                DistinctRoles = List.Distinct([Role]),
                RoleCount = List.Count(DistinctRoles)
            in
                RoleCount, Int64.Type}}),
    JoinedCounts = Table.NestedJoin(UniqueWorkers, {"EmployeeID", "Facility-Abbrev"},
        RoleCounts, {"EmployeeID", "Facility-Abbrev"}, "Roles", JoinKind.LeftOuter),
    SelectedAllocationRoleCounts = Table.SelectColumns(AllocationRoleShares,
        {"EmployeeID", "Facility-Abbrev", "RoleCount"}),
    AllocationRoleCounts = Table.Distinct(SelectedAllocationRoleCounts),
    JoinedAllocationRoles = Table.NestedJoin(JoinedCounts, {"EmployeeID", "Facility-Abbrev"},
        AllocationRoleCounts, {"EmployeeID", "Facility-Abbrev"}, "AllocationRoles", JoinKind.LeftOuter),
    // Use employee ID and facility, not a possibly repeated name, for the multi-role suffix convention.
    ResolveDisplayName = (row as record) as text =>
        let
            HasMultipleReconciledRoles = row[Roles]{0}[RoleCount] > 1,
            HasAllocationRoles = not Table.IsEmpty(row[AllocationRoles]),
            HasMultipleAllocationRoles = if HasAllocationRoles
                then row[AllocationRoles]{0}[RoleCount] > 1
                else false,
            NeedsRoleSuffix = HasMultipleReconciledRoles or HasMultipleAllocationRoles,
            DisplayName = if NeedsRoleSuffix
                then row[BaseName] & " (" & row[Role] & ")"
                else row[BaseName]
        in
            DisplayName,
    DisplayNames = Table.AddColumn(JoinedAllocationRoles, "Name", ResolveDisplayName, type text),
    // Worker Reconciliation selects the preferred roster role; retain its LINK Roles Role Group for downstream reporting.
    PreferredRole = Table.AddColumn(DisplayNames, "PreferredRole", each [Role], type text),
    RemovedNestedTables = Table.RemoveColumns(PreferredRole, {"Roles", "AllocationRoles"}),
    BufferedWorkers = Table.Buffer(RemovedNestedTables)
in
    BufferedWorkers;

// Query: WorkerIdentityFailures_DIAGNOSTICS
// Purpose: Identify the worker rows behind the blocking reconciliation checks.
// Output: One row per affected worker/check, including employee, facility, role, name and failure detail.
// Notes: StaffListMaster requires one preferred-role resource per employee and facility.
shared WorkerIdentityFailures_DIAGNOSTICS = let
    Workers = ReconciledWorkers_Eligible,
    WorkerKeyCounts = Table.Group(Workers, {"EmployeeID", "Facility-Abbrev", "Role"},
        {{"Rows", each Table.RowCount(_), Int64.Type}}),
    JoinedWorkerKeys = Table.NestedJoin(Workers, {"EmployeeID", "Facility-Abbrev", "Role"},
        WorkerKeyCounts, {"EmployeeID", "Facility-Abbrev", "Role"}, "WorkerKeyCount", JoinKind.LeftOuter),
    ExpandedWorkerKeys = Table.ExpandTableColumn(JoinedWorkerKeys, "WorkerKeyCount", {"Rows"}, {"KeyRows"}),
    DuplicateWorkerKeys = Table.SelectRows(ExpandedWorkerKeys, each [KeyRows] > 1),
    MarkedWorkerKeys = Table.AddColumn(DuplicateWorkerKeys, "Check", each "Unique worker keys", type text),
    WorkerKeyDetails = Table.AddColumn(MarkedWorkerKeys, "Details", each
        Number.ToText([KeyRows]) & " rows share this employee/facility/role key.", type text),
    CombineSortedDistinctText = (values as list) as text =>
        let
            DistinctValues = List.Distinct(values),
            SortedValues = List.Sort(DistinctValues),
            CombinedValues = Text.Combine(SortedValues, ", ")
        in
            CombinedValues,
    PreferredRoleCounts = Table.Group(Workers, {"EmployeeID", "Facility-Abbrev"},
        {{"Rows", each Table.RowCount(_), Int64.Type},
         {"Roles", each CombineSortedDistinctText([Role]), type text}}),
    JoinedPreferredRoles = Table.NestedJoin(Workers, {"EmployeeID", "Facility-Abbrev"},
        PreferredRoleCounts, {"EmployeeID", "Facility-Abbrev"}, "PreferredRoleCount", JoinKind.LeftOuter),
    ExpandedPreferredRoles = Table.ExpandTableColumn(JoinedPreferredRoles, "PreferredRoleCount",
        {"Rows", "Roles"}, {"PreferredRows", "PreferredRoles"}),
    MultiplePreferredRoles = Table.SelectRows(ExpandedPreferredRoles, each [PreferredRows] > 1),
    MarkedPreferredRoles = Table.AddColumn(MultiplePreferredRoles, "Check", each
        "One preferred role per employee and facility", type text),
    PreferredRoleDetails = Table.AddColumn(MarkedPreferredRoles, "Details", each
        "Preferred-role rows: " & [PreferredRoles] & ". StaffListMaster requires one resource for this employee/facility.", type text),
    PublishedNameCounts = Table.Group(Workers, {"Name"},
        {{"Rows", each Table.RowCount(_), Int64.Type}}),
    JoinedPublishedNames = Table.NestedJoin(Workers, {"Name"}, PublishedNameCounts,
        {"Name"}, "PublishedNameCount", JoinKind.LeftOuter),
    ExpandedPublishedNames = Table.ExpandTableColumn(JoinedPublishedNames, "PublishedNameCount",
        {"Rows"}, {"NameRows"}),
    DuplicatePublishedNames = Table.SelectRows(ExpandedPublishedNames, each [NameRows] > 1),
    MarkedPublishedNames = Table.AddColumn(DuplicatePublishedNames, "Check", each
        "Unambiguous published names", type text),
    PublishedNameDetails = Table.AddColumn(MarkedPublishedNames, "Details", each
        Number.ToText([NameRows]) & " worker rows publish the same name.", type text),
    ContractIssues = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit and [Role] <> null and [ContractIssue] <> null),
    NamedContractIssues = Table.RenameColumns(ContractIssues, {{"BaseName", "Name"}}),
    MarkedContracts = Table.AddColumn(NamedContractIssues, "Check", each
        "Valid relevant worker contracts", type text),
    ContractDetails = Table.AddColumn(MarkedContracts, "Details", each [ContractIssue], type text),
    RoleMappingIssues = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit and [SourceRoleKey] <> null and [RoleMappingIssue] <> null),
    NamedRoleMappingIssues = Table.RenameColumns(RoleMappingIssues, {{"BaseName", "Name"}}),
    MarkedRoleMappings = Table.AddColumn(NamedRoleMappingIssues, "Check", each
        "Current worker roles map through LINK Roles", type text),
    RoleMappingDetails = Table.AddColumn(MarkedRoleMappings, "Details", each
        [RoleMappingIssue] & ": source role '" & [SourceRole]
            & "' (key '" & [SourceRoleKey] & "').", type text),
    Combined = Table.Combine({WorkerKeyDetails, PreferredRoleDetails, PublishedNameDetails,
        ContractDetails, RoleMappingDetails}),
    Result = Table.SelectColumns(Combined,
        {"Check", "EmployeeID", "Facility-Abbrev", "Role", "Name", "Details"}),
    SortedResult = Table.Sort(Result,
        {{"Check", Order.Ascending}, {"EmployeeID", Order.Ascending}, {"Role", Order.Ascending}})
in
    SortedResult;

// Query: WorkerAvailability_DIAGNOSTICS
// Purpose: Show excluded reconciliation rows and extraction workers absent from the selected register.
shared WorkerAvailability_DIAGNOSTICS = let
    Excluded = Table.SelectRows(ReconciledWorkers_Prepare, each [Issue] <> null),
    SelectedSourceWorkers = Table.SelectColumns(#"IMPORT Availability Leave Source",
        {"Payroll Code", "Facility-Abbrev"}),
    SourceWorkers = Table.Distinct(SelectedSourceWorkers),
    NotRegistered = Table.NestedJoin(SourceWorkers, {"Payroll Code", "Facility-Abbrev"},
        #"IMPORT Reconciled Workers", {"EmployeeID", "Facility-Abbrev"}, "Register", JoinKind.LeftAnti),
    RemovedRegister = Table.RemoveColumns(NotRegistered, {"Register"}),
    MissingWorkers = Table.RenameColumns(RemovedRegister, {{"Payroll Code", "EmployeeID"}}),
    MarkedMissing = Table.AddColumn(MissingWorkers, "Issue", each "Extraction worker absent from reconciliation", type text),
    CombinedDiagnostics = Table.Combine({Excluded, MarkedMissing})
in
    CombinedDiagnostics;

// Query: AvailabilityRecords_Invalid
// Purpose: Identify invalid intervals for included workers, plus records without a usable worker ID.
shared AvailabilityRecords_Invalid = let
    Invalid = Table.SelectRows(AvailabilityRecords, each [Issue] <> null),
    SelectedWorkerKeys = Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"}),
    WorkerKeys = Table.Distinct(SelectedWorkerKeys),
    Matched = Table.NestedJoin(Invalid, {"Payroll Code", "Facility-Abbrev"},
        WorkerKeys, {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.LeftOuter),
    RequiredFailures = Table.SelectRows(Matched, each [Payroll Code] = null or not Table.IsEmpty([Worker])),
    RemovedWorkerMatch = Table.RemoveColumns(RequiredFailures, {"Worker"})
in
    RemovedWorkerMatch;

// Query: AvailabilityReasons_DIAGNOSTICS
// Purpose: Show unknown source reasons and their record IDs for explicit mapping review.
shared AvailabilityReasons_DIAGNOSTICS = let
    Unknown = Table.SelectRows(AvailabilityRecords, each [RecordKind] = "Unknown"),
    SelectedDetails = Table.SelectColumns(Unknown,
        {"SourceRecord", "Payroll Code", "Facility-Abbrev", "OriginalReason", "Start", "End", "Issue"})
in
    SelectedDetails;

// Query: AvailabilityShiftDefinitions
// Purpose: Validate the role-specific shift endpoints and durations from Settings Data.
// Notes: StartDay/EndDay are fractions of a day; EndTimeCheck is a clock time; StartTime is hours.
shared AvailabilityShiftDefinitions = let
    Data = #"IMPORT Availability Settings"{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    Selected = Table.SelectColumns(Data, {"Role", "ShiftPeriod", "StartTime", "StartDay", "DurationOfShifts", "EndDay", "EndTimeCheck"}),
    NursingRoles = Table.SelectRows(Selected, each List.Contains({"RN", "AIN", "EN"}, [Role])),
    ParseDefinition = (row as record) as record =>
        let
            Start = Number.From(row[StartDay]),
            End = Number.From(row[EndDay]),
            Hours = Number.From(row[DurationOfShifts]),
            StartHours = Number.From(row[StartTime]),
            EndClock = Time.From(row[EndTimeCheck]),
            EndHourSeconds = Time.Hour(EndClock) * 3600,
            EndMinuteSeconds = Time.Minute(EndClock) * 60,
            EndSeconds = Time.Second(EndClock),
            ClockEnd = (EndHourSeconds + EndMinuteSeconds + EndSeconds) / 86400,
            AdjustedEnd = if End <= Start then End + 1 else End,
            ClockHours = (AdjustedEnd - Start) * 24,
            HasAllValues = List.NonNullCount({Start, End, Hours, StartHours, EndClock}) = 5,
            StartIsDayFraction = Start >= 0 and Start < 1,
            EndIsDayFraction = End >= 0 and End < 1,
            DurationIsValid = Hours > 0 and Hours < 24,
            StartHoursMatch = Number.Abs(StartHours - Start * 24) < 0.00000001,
            EndClockMatches = Number.Abs(ClockEnd - End) < 0.00000001,
            DurationMatches = Number.Abs(ClockHours - Hours) < 0.00000001,
            ShiftNameIsValid = List.Contains({"AM", "PM", "NIGHT"}, row[ShiftPeriod]),
            IsValid = HasAllValues and StartIsDayFraction and EndIsDayFraction
                and DurationIsValid and StartHoursMatch and EndClockMatches
                and DurationMatches and ShiftNameIsValid,
            StartSeconds = Number.Round(Start * 86400, 6),
            EndOffsetSeconds = Number.Round(AdjustedEnd * 86400, 6),
            Definition = if IsValid then [
                StartOffset = #duration(0, 0, 0, StartSeconds),
                EndOffset = #duration(0, 0, 0, EndOffsetSeconds)
            ] else error "Invalid or inconsistent role-specific ShiftPeriod definition."
        in
            Definition,
    Parsed = Table.AddColumn(NursingRoles, "Definition", ParseDefinition),
    Expanded = Table.ExpandRecordColumn(Parsed, "Definition", {"StartOffset", "EndOffset"}),
    DistinctDefinitions = Table.Distinct(Expanded, {"Role", "ShiftPeriod"}),
    DefinitionCount = Table.RowCount(Expanded),
    DistinctDefinitionCount = Table.RowCount(DistinctDefinitions),
    Unique = if DefinitionCount = DistinctDefinitionCount then Expanded
        else error "Duplicate role/shift definitions in ShiftPeriod.",
    // Force endpoint validation before a missing calendar match could conceal it.
    DefinitionRecords = Table.ToRecords(Unique),
    EndpointChecks = List.Transform(DefinitionRecords, each [EndOffset] > [StartOffset]),
    EndpointsAreValid = List.AllTrue(EndpointChecks),
    Validated = if EndpointsAreValid then Unique
        else error "Invalid shift endpoints."
in
    Validated;

// Query: AvailabilityShiftWindows
// Purpose: Build one dated role/shift window for every Settings permutation, including the final night.
shared AvailabilityShiftWindows = let
    Data = #"IMPORT Availability Settings"{[Item = "PermutationDimensions", Kind = "Table"]}[Data],
    Selected = Table.SelectColumns(Data, {"Date", "Shifts", "RolesList", "Period"}),
    Renamed = Table.RenameColumns(Selected, {{"Shifts", "Shift"}, {"RolesList", "Role"}}),
    NursingRoles = Table.SelectRows(Renamed, each List.Contains({"RN", "AIN", "EN"}, [Role])),
    Typed = Table.TransformColumnTypes(NursingRoles, {{"Date", type date}, {"Period", Int64.Type}, {"Role", type text}, {"Shift", type text}}),
    CalendarIsEmpty = Table.IsEmpty(Typed),
    MissingKeyRows = Table.SelectRows(Typed, each [Date] = null or [Period] = null or [Shift] = null),
    MissingKeyCount = Table.RowCount(MissingKeyRows),
    DistinctDatedShifts = Table.Distinct(Typed, {"Date", "Role", "Shift"}),
    DistinctRolePeriods = Table.Distinct(Typed, {"Role", "Period"}),
    CalendarRowCount = Table.RowCount(Typed),
    DatedShiftCount = Table.RowCount(DistinctDatedShifts),
    RolePeriodCount = Table.RowCount(DistinctRolePeriods),
    HasDuplicateKeys = CalendarRowCount <> DatedShiftCount or CalendarRowCount <> RolePeriodCount,
    ValidCalendar = if CalendarIsEmpty then error "Settings roster calendar is empty."
        else if MissingKeyCount > 0 then
            error "Settings roster calendar contains missing keys."
        else if HasDuplicateKeys then
            error "Settings roster calendar contains duplicate keys."
        else Typed,
    Joined = Table.NestedJoin(ValidCalendar, {"Role", "Shift"},
        AvailabilityShiftDefinitions, {"Role", "ShiftPeriod"}, "Definition", JoinKind.LeftOuter),
    IncompleteDefinitions = Table.SelectRows(Joined, each Table.RowCount([Definition]) <> 1),
    IncompleteDefinitionCount = Table.RowCount(IncompleteDefinitions),
    Complete = if IncompleteDefinitionCount = 0 then Joined
        else error "A configured period lacks exactly one role-specific shift definition.",
    Expanded = Table.ExpandTableColumn(Complete, "Definition", {"StartOffset", "EndOffset"}),
    Starts = Table.AddColumn(Expanded, "ShiftStart", each
        let
            ShiftDate = DateTime.From([Date]),
            ShiftStart = ShiftDate + [StartOffset]
        in
            ShiftStart, type datetime),
    Ends = Table.AddColumn(Starts, "ShiftEnd", each
        let
            ShiftDate = DateTime.From([Date]),
            ShiftEnd = ShiftDate + [EndOffset]
        in
            ShiftEnd, type datetime),
    RosterStart = List.Min(ValidCalendar[Date]),
    Weeks = Table.AddColumn(Ends, "Week", each
        let
            ElapsedDuration = [Date] - RosterStart,
            ElapsedDays = Duration.Days(ElapsedDuration),
            CompletedWeeks = Number.IntegerDivide(ElapsedDays, 7),
            Week = 1 + CompletedWeeks
        in
            Week, Int64.Type),
    Days = Table.AddColumn(Weeks, "Day", each Date.ToText([Date], "ddd", "en-US"), type text),
    Result = Table.RemoveColumns(Days, {"StartOffset", "EndOffset"}),
    BufferedResult = Table.Buffer(Result)
in
    BufferedResult;

// Query: AvailabilityEligibleRecords
// Purpose: Keep valid eligible-worker records on complete calendar dates covered by the roster.
// Notes: Unique worker/facility keys prevent multiple roles multiplying the source records.
shared AvailabilityEligibleRecords = let
    SelectedWorkerKeys = Table.SelectColumns(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"}),
    WorkerKeys = Table.Distinct(SelectedWorkerKeys),
    FirstShiftStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    FirstShiftDate = Date.From(FirstShiftStart),
    CalendarStart = DateTime.From(FirstShiftDate),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    HorizonEndTime = Time.From(HorizonEnd),
    HorizonEndDate = Date.From(HorizonEnd),
    HorizonEndMidnight = DateTime.From(HorizonEndDate),
    EndsAtMidnight = HorizonEndTime = #time(0, 0, 0),
    CalendarEnd = if EndsAtMidnight then HorizonEnd
        else HorizonEndMidnight + #duration(1, 0, 0, 0),
    IsEligibleRecord = (row as record) as logical =>
        let
            HasNoIssue = row[Issue] = null,
            StartsBeforeCalendarEnd = row[Start] < CalendarEnd,
            EndsAfterCalendarStart = row[End] > CalendarStart,
            IsEligible = HasNoIssue and StartsBeforeCalendarEnd and EndsAfterCalendarStart
        in
            IsEligible,
    Valid = Table.SelectRows(AvailabilityRecords, IsEligibleRecord),
    Joined = Table.NestedJoin(Valid, {"Payroll Code", "Facility-Abbrev"}, WorkerKeys,
        {"EmployeeID", "Facility-Abbrev"}, "Worker", JoinKind.Inner),
    Narrow = Table.RemoveColumns(Joined, {"Worker", "Available"}),
    ClipStart = (value as datetime) as datetime =>
        let
            ClippedStart = List.Max({value, CalendarStart})
        in
            ClippedStart,
    ClipEnd = (value as datetime) as datetime =>
        let
            ClippedEnd = List.Min({value, CalendarEnd})
        in
            ClippedEnd,
    Clipped = Table.TransformColumns(Narrow, {
        {"Start", ClipStart, type datetime},
        {"End", ClipEnd, type datetime}
    }),
    BufferedRecords = Table.Buffer(Clipped)
in
    BufferedRecords;

// Query: WorkerLeaveDays
// Purpose: Decide full-day leave once per worker/facility/date, independently of role and UNAVAIL.
// Notes: A full-day decision blocks every configured shift STARTING on that date, including NIGHT.
shared WorkerLeaveDays = let
    Leave = Table.SelectRows(AvailabilityEligibleRecords, each [RecordKind] = "Leave"),
    BuildLeaveDayTotals = (rows as table) as table =>
        let
            SelectedIntervals = Table.SelectColumns(rows, {"Start", "End"}),
            IntervalRecords = Table.ToRecords(SelectedIntervals),
            DayTotals = AvailabilityLeaveDayTotals(IntervalRecords)
        in
            DayTotals,
    Workers = Table.Group(Leave, {"Payroll Code", "Facility-Abbrev"},
        {{"Days", BuildLeaveDayTotals, type table}}),
    Expanded = Table.ExpandTableColumn(Workers, "Days", {"Date", "SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount"}),
    Identified = Table.RenameColumns(Expanded, {{"Payroll Code", "EmployeeID"}}),
    Threshold = #"EXTRACT EffectiveShiftHrs",
    Blocked = Table.AddColumn(Identified, "FullDayLeaveBlocked", each [DistinctLeaveHours] >= Threshold, type logical),
    // Scalar daily results are reused by each role/shift and its diagnostic rows.
    BufferedDays = Table.Buffer(Blocked)
in
    BufferedDays;

// Query: WorkerAvailabilityRules
// Purpose: Build date-specific availability windows and separate raw leave and UNAVAIL intervals.
// Notes: UNAVAIL is never promoted to a full-day exclusion; WorkerLeaveDays owns that leave-only decision.
shared WorkerAvailabilityRules = let
    HorizonStart = List.Min(AvailabilityShiftWindows[ShiftStart]),
    HorizonEnd = List.Max(AvailabilityShiftWindows[ShiftEnd]),
    HorizonStartDate = Date.From(HorizonStart),
    CalendarStart = DateTime.From(HorizonStartDate),
    HorizonEndTime = Time.From(HorizonEnd),
    HorizonEndDate = Date.From(HorizonEnd),
    HorizonEndMidnight = DateTime.From(HorizonEndDate),
    EndsAtMidnight = HorizonEndTime = #time(0, 0, 0),
    CalendarEnd = if EndsAtMidnight then HorizonEnd
        else HorizonEndMidnight + #duration(1, 0, 0, 0),
    Intervals = (records as table, kind as text) as list =>
        let
            MatchingRecords = Table.SelectRows(records, each [RecordKind] = kind),
            SelectedIntervals = Table.SelectColumns(MatchingRecords, {"Start", "End"}),
            IntervalRecords = Table.ToRecords(SelectedIntervals),
            UnionedIntervals = AvailabilityUnion(IntervalRecords),
            BufferedIntervals = List.Buffer(UnionedIntervals)
        in
            BufferedIntervals,
    AvailableIntervals = (records as table) as list =>
        let
            RawAvailableIntervals = Intervals(records, "AVAIL"),
            DailyAvailableWindows = AvailabilityDailyWindows(RawAvailableIntervals, CalendarStart, CalendarEnd)
        in
            DailyAvailableWindows,
    UnavailIntervals = (records as table) as list =>
        let
            RawUnavailIntervals = Intervals(records, "UNAVAIL")
        in
            RawUnavailIntervals,
    LeaveIntervals = (records as table) as list =>
        let
            RawLeaveIntervals = Intervals(records, "Leave")
        in
            RawLeaveIntervals,
    Grouped = Table.Group(AvailabilityEligibleRecords, {"Payroll Code", "Facility-Abbrev"}, {
        {"AvailableIntervals", AvailableIntervals, type list},
        {"UnavailIntervals", UnavailIntervals, type list},
        {"LeaveIntervals", LeaveIntervals, type list}
    }),
    Joined = Table.NestedJoin(ReconciledWorkers_Eligible, {"EmployeeID", "Facility-Abbrev"},
        Grouped, {"Payroll Code", "Facility-Abbrev"}, "Rules", JoinKind.LeftOuter),
    ResolveRule = (rules as table) as record =>
        let
            RulesAreEmpty = Table.IsEmpty(rules),
            DefaultRule = [
                AvailableIntervals = {[Start = HorizonStart, End = HorizonEnd]},
                UnavailIntervals = {},
                LeaveIntervals = {}
            ],
            ExistingRule = if RulesAreEmpty then null
                else Record.SelectFields(rules{0}, {"AvailableIntervals", "UnavailIntervals", "LeaveIntervals"}),
            Rule = if RulesAreEmpty then DefaultRule else ExistingRule
        in
            Rule,
    ReadRule = (row as record) as record =>
        let
            Rule = ResolveRule(row[Rules])
        in
            Rule,
    Defaults = Table.AddColumn(Joined, "Rule", ReadRule),
    RemovedRules = Table.RemoveColumns(Defaults, {"Rules"}),
    Expanded = Table.ExpandRecordColumn(RemovedRules, "Rule",
        {"AvailableIntervals", "UnavailIntervals", "LeaveIntervals"}),
    // Nested interval lists are explicitly buffered in their constructors; the outer table is small.
    BufferedRules = Table.Buffer(Expanded)
in
    BufferedRules;

// Query: WorkerShiftRoleCoverage_DIAGNOSTICS
// Purpose: Show every eligible worker role and the number of dated Settings shifts matched to it.
// Output: One row per eligible worker/facility/role; MissingShiftWindow remains inspectable without raising an error.
shared WorkerShiftRoleCoverage_DIAGNOSTICS = let
    CountShiftWindows = (rows as table) as number =>
        let
            ShiftWindowCount = Table.RowCount(rows)
        in
            ShiftWindowCount,
    ShiftCounts = Table.Group(AvailabilityShiftWindows, {"Role"},
        {{"ShiftWindowCount", CountShiftWindows, Int64.Type}}),
    WorkerRoles = Table.SelectColumns(WorkerAvailabilityRules,
        {"EmployeeID", "Facility-Abbrev", "Name", "Role"}),
    MatchedRoles = Table.NestedJoin(WorkerRoles, {"Role"}, ShiftCounts,
        {"Role"}, "ConfiguredShifts", JoinKind.LeftOuter),
    ResolveShiftWindowCount = (configuredShifts as table) as number =>
        let
            ConfiguredShiftsAreEmpty = Table.IsEmpty(configuredShifts),
            ConfiguredShiftCount = if ConfiguredShiftsAreEmpty then 0
                else configuredShifts{0}[ShiftWindowCount]
        in
            ConfiguredShiftCount,
    ReadShiftWindowCount = (row as record) as number =>
        let
            ShiftWindowCount = ResolveShiftWindowCount(row[ConfiguredShifts])
        in
            ShiftWindowCount,
    CountedShifts = Table.AddColumn(MatchedRoles, "ShiftWindowCount", ReadShiftWindowCount, Int64.Type),
    HasMissingShiftWindow = (row as record) as logical =>
        let
            IsMissing = row[ShiftWindowCount] = 0
        in
            IsMissing,
    MissingShiftWindow = Table.AddColumn(CountedShifts, "MissingShiftWindow", HasMissingShiftWindow, type logical),
    Result = Table.RemoveColumns(MissingShiftWindow, {"ConfiguredShifts"})
in Result;

// Query: WorkerShiftCalendar
// Purpose: Match eligible workers to dated, role-specific shift windows before leave or availability calculations.
// Notes: A missing role is visible in WorkerShiftRoleCoverage_DIAGNOSTICS and blocks downstream publication.
shared WorkerShiftCalendar = let
    MissingRoles = Table.SelectRows(WorkerShiftRoleCoverage_DIAGNOSTICS, each [MissingShiftWindow]),
    ValidatedRoles = if Table.IsEmpty(MissingRoles) then WorkerAvailabilityRules
        else error Error.Record("Worker shift role coverage",
            "Eligible worker roles have no configured Settings roster shifts. Inspect WorkerShiftRoleCoverage_DIAGNOSTICS.",
            [Unit = Unit, MissingWorkers = MissingRoles]),
    MatchedWindows = Table.NestedJoin(ValidatedRoles, {"Role"}, AvailabilityShiftWindows,
        {"Role"}, "RosterShifts", JoinKind.Inner),
    DatedShifts = Table.ExpandTableColumn(MatchedWindows, "RosterShifts",
        {"Date", "Shift", "Period", "ShiftStart", "ShiftEnd", "Week", "Day"})
in DatedShifts;

// Query: WorkerShiftDayAssessment
// Purpose: Attach each worker's date-level leave decision to the matched shift calendar.
// Notes: The leave decision is shared by all shifts starting on that date.
shared WorkerShiftDayAssessment = let
    JoinedLeaveDays = Table.NestedJoin(WorkerShiftCalendar, {"EmployeeID", "Facility-Abbrev", "Date"},
        WorkerLeaveDays, {"EmployeeID", "Facility-Abbrev", "Date"}, "LeaveDay", JoinKind.LeftOuter),
    ResolveDayAssessment = (leaveDay as table) as record =>
        let
            HasLeaveDay = not Table.IsEmpty(leaveDay),
            DefaultAssessment = [
                SumOfLeaveRecordHours = 0,
                DistinctLeaveHours = 0,
                LeaveRecordCount = 0,
                FullDayLeaveBlocked = false
            ],
            LeaveAssessment = if HasLeaveDay then Record.SelectFields(leaveDay{0},
                {"SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked"})
                else null,
            Assessment = if HasLeaveDay then LeaveAssessment else DefaultAssessment
        in
            Assessment,
    DayAssessment = Table.AddColumn(JoinedLeaveDays, "DayAssessment",
        each ResolveDayAssessment([LeaveDay])),
    RemovedLeaveDay = Table.RemoveColumns(DayAssessment, {"LeaveDay"}),
    DayColumns = Table.ExpandRecordColumn(RemovedLeaveDay, "DayAssessment",
        {"SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked"})
in
    DayColumns;

// Query: WorkerShiftSegments
// Purpose: Measure baseline, leave, UNAVAIL and residual clock hours for each worker's dated shift.
// Output: Scalar measurements and checks, including zero-hour shifts; no historical lists on expanded rows.
shared WorkerShiftSegments = let
    MeasureShift = (row as record) as record =>
        let
            Measurement = AvailabilityShiftMeasurement(
                row[AvailableIntervals],
                row[UnavailIntervals],
                row[LeaveIntervals],
                row[ShiftStart],
                row[ShiftEnd],
                row[FullDayLeaveBlocked]
            )
        in
            Measurement,
    Measured = Table.AddColumn(WorkerShiftDayAssessment, "Measurement", MeasureShift),
    SelectedScalars = Table.SelectColumns(Measured,
        {"EmployeeID", "Facility-Abbrev", "Role", "Name", "Date", "Shift", "Week", "Day", "ShiftStart", "ShiftEnd",
         "SumOfLeaveRecordHours", "DistinctLeaveHours", "LeaveRecordCount", "FullDayLeaveBlocked", "Measurement"}),
    ScalarRows = Table.ExpandRecordColumn(SelectedScalars, "Measurement",
        {"BaselineHours", "LeaveHoursInShift", "UnavailHoursInShift", "CombinedAbsenceHoursInShift",
         "AvailableHours", "FullShift", "SegmentsValid", "Disjoint"}),
    BufferedRows = Table.Buffer(ScalarRows)
in
    BufferedRows;

// Query: ResDayShift_Calculated
// Purpose: Deduct distinct shift absences from ShiftDuration and respect the remaining AVAIL windows.
// Notes: Full-day leave blocks all shifts on the date; a zero-hour shift alone never blocks other shifts.
shared ResDayShift_Calculated = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    WithAllowance = Table.AddColumn(WorkerShiftSegments, "ShiftDuration", each Allowance, type number),
    CalculateEffectiveHours = (row as record) as number =>
        let
            EffectiveHours = AvailabilityResidualHours(
                row[AvailableHours],
                row[CombinedAbsenceHoursInShift],
                row[FullShift],
                row[FullDayLeaveBlocked],
                Allowance
            )
        in
            EffectiveHours,
    Effective = Table.AddColumn(WithAllowance, "EffectiveShiftHrs", CalculateEffectiveHours, type number),
    DetermineAvailabilityDecision = (row as record) as text =>
        let
            Decision = if row[FullDayLeaveBlocked] then
                "All shifts blocked: daily recognised leave reaches ShiftDuration"
                else if row[BaselineHours] = 0 then
                    "Unavailable: outside this date's AVAIL windows"
                else if row[AvailableHours] = 0 then
                    "Unavailable: recorded absences cover available time in this shift"
                else if row[CombinedAbsenceHoursInShift] >= Allowance then
                    "Unavailable: shift absences exhaust ShiftDuration"
                else if row[FullShift] then
                    "Full shift available"
                else
                    "Residual availability"
        in
            Decision,
    Outcome = Table.AddColumn(Effective, "AvailabilityDecision",
        DetermineAvailabilityDecision, type text)
in
    Outcome;

// Query: WorkerRolePipeline_DIAGNOSTICS
// Purpose: Trace every prepared worker-role row through eligibility, rule construction and positive calculated shifts.
// Output: One row per ReconciledWorkers_Prepare row, with downstream counts and the last reached stage.
// Notes: SourceRole joins to LINK Roles Roster Roles; MappedRoleGroup supplies the downstream Role.
shared WorkerRolePipeline_DIAGNOSTICS = let
    KeyColumns = {"EmployeeID", "Facility-Abbrev", "Role"},
    PreparedColumns = {
        "EmployeeID", "Facility-Abbrev", "BaseName", "SourceRole", "SourceRoleKey",
        "RoleMappingMatches", "MappedRosterRole", "RoleMappingDefinitionRows",
        "DistinctRoleGroupCount", "MappedRoleGroup", "RoleMappingIssue", "Role",
        "EmploymentType", "Contracted FN Hours", "ContractIssue", "Issue"
    },
    Prepared = Table.SelectColumns(ReconciledWorkers_Prepare, PreparedColumns),
    PreparedMappedRoleCounts = Table.Group(Prepared, KeyColumns,
        {{"PreparedMappedRoleRows", each Table.RowCount(_), Int64.Type}}),
    JoinedPreparedCounts = Table.NestedJoin(Prepared, KeyColumns,
        PreparedMappedRoleCounts, KeyColumns, "PreparedStage", JoinKind.LeftOuter),
    ExpandedPreparedCounts = Table.ExpandTableColumn(JoinedPreparedCounts, "PreparedStage",
        {"PreparedMappedRoleRows"}, {"PreparedMappedRoleRows"}),

    EligibleKeys = Table.SelectColumns(ReconciledWorkers_Eligible, KeyColumns),
    EligibleCounts = Table.Group(EligibleKeys, KeyColumns,
        {{"EligibleRows", each Table.RowCount(_), Int64.Type}}),
    JoinedEligible = Table.NestedJoin(ExpandedPreparedCounts, KeyColumns,
        EligibleCounts, KeyColumns, "EligibleStage", JoinKind.LeftOuter),
    ExpandedEligible = Table.ExpandTableColumn(JoinedEligible, "EligibleStage",
        {"EligibleRows"}, {"EligibleRows"}),

    RuleKeys = Table.SelectColumns(WorkerAvailabilityRules, KeyColumns),
    RuleCounts = Table.Group(RuleKeys, KeyColumns,
        {{"RuleRows", each Table.RowCount(_), Int64.Type}}),
    JoinedRules = Table.NestedJoin(ExpandedEligible, KeyColumns,
        RuleCounts, KeyColumns, "RuleStage", JoinKind.LeftOuter),
    ExpandedRules = Table.ExpandTableColumn(JoinedRules, "RuleStage",
        {"RuleRows"}, {"RuleRows"}),

    ConfiguredShiftCounts = Table.Group(AvailabilityShiftWindows, {"Role"},
        {{"ConfiguredShiftRows", each Table.RowCount(_), Int64.Type}}),
    JoinedConfiguredShifts = Table.NestedJoin(ExpandedRules, {"Role"},
        ConfiguredShiftCounts, {"Role"}, "ConfiguredShiftStage", JoinKind.LeftOuter),
    ExpandedConfiguredShifts = Table.ExpandTableColumn(JoinedConfiguredShifts,
        "ConfiguredShiftStage", {"ConfiguredShiftRows"}, {"ConfiguredShiftRows"}),

    SummarizeCalculatedRows = (rows as table) as record =>
        let
            PositiveRows = Table.SelectRows(rows, each [EffectiveShiftHrs] > 0),
            PositiveRowCount = Table.RowCount(PositiveRows),
            PositiveEffectiveHours = if PositiveRowCount = 0 then 0
                else List.Sum(PositiveRows[EffectiveShiftHrs]),
            Summary = [
                CalculatedShiftRows = Table.RowCount(rows),
                PositiveShiftRows = PositiveRowCount,
                PositiveDates = List.Count(List.Distinct(PositiveRows[Date])),
                PositiveEffectiveHours = PositiveEffectiveHours
            ]
        in
            Summary,
    CalculatedCountsAttempt = try
        let
            CalculatedSelected = Table.SelectColumns(ResDayShift_Calculated,
                KeyColumns & {"Date", "EffectiveShiftHrs"}),
            CalculatedCounts = Table.Group(CalculatedSelected, KeyColumns,
                {{"CalculatedSummary", SummarizeCalculatedRows,
                    type [CalculatedShiftRows = Int64.Type, PositiveShiftRows = Int64.Type,
                        PositiveDates = Int64.Type, PositiveEffectiveHours = number]}}),
            ExpandedCalculatedCounts = Table.ExpandRecordColumn(CalculatedCounts,
                "CalculatedSummary",
                {"CalculatedShiftRows", "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours"}),
            BufferedCalculatedCounts = Table.Buffer(ExpandedCalculatedCounts)
        in
            BufferedCalculatedCounts,
    CalculatedTraceAvailable = not CalculatedCountsAttempt[HasError],
    ReadAttemptError = (attempt as record) as nullable text =>
        let
            MessageAttempt = if attempt[HasError] then try attempt[Error][Message] else null,
            Message = if not attempt[HasError] then null
                else if MessageAttempt[HasError] then "Evaluation failed"
                else if MessageAttempt[Value] = null then "Evaluation failed"
                else Text.From(MessageAttempt[Value])
        in
            Message,
    CalculatedTraceError = ReadAttemptError(CalculatedCountsAttempt),
    EmptyCalculatedCounts = #table(
        {"EmployeeID", "Facility-Abbrev", "Role", "CalculatedShiftRows",
            "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours"},
        {}),
    CalculatedCounts = if CalculatedTraceAvailable then CalculatedCountsAttempt[Value]
        else EmptyCalculatedCounts,
    JoinedCalculated = Table.NestedJoin(ExpandedConfiguredShifts, KeyColumns,
        CalculatedCounts, KeyColumns, "CalculatedStage", JoinKind.LeftOuter),
    ExpandedCalculated = Table.ExpandTableColumn(JoinedCalculated, "CalculatedStage",
        {"CalculatedShiftRows", "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours"},
        {"CalculatedShiftRows", "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours"}),

    FilledCounts = Table.ReplaceValue(ExpandedCalculated, null, 0,
        Replacer.ReplaceValue, {
            "PreparedMappedRoleRows", "EligibleRows", "RuleRows", "ConfiguredShiftRows", "CalculatedShiftRows",
            "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours"
        }),
    AddedMappedRoleCollision = Table.AddColumn(FilledCounts, "MappedRoleCollision", each
        [Role] <> null and [PreparedMappedRoleRows] > 1, type logical),
    AddedCalculatedTraceStatus = Table.AddColumn(AddedMappedRoleCollision, "CalculatedTraceAvailable",
        each CalculatedTraceAvailable, type logical),
    AddedCalculatedTraceError = Table.AddColumn(AddedCalculatedTraceStatus, "CalculatedTraceError",
        each CalculatedTraceError, type nullable text),
    TraceOutcome = Table.AddColumn(AddedCalculatedTraceError, "TraceOutcome", each
        if [RoleMappingIssue] <> null then "Excluded by LINK Roles mapping: " & [RoleMappingIssue]
        else if [Issue] <> null then "Excluded in ReconciledWorkers_Prepare: " & [Issue]
        else if [EligibleRows] = 0 then "Missing from ReconciledWorkers_Eligible"
        else if [RuleRows] = 0 then "Missing from WorkerAvailabilityRules"
        else if [ConfiguredShiftRows] = 0 then "No configured shift windows for mapped role"
        else if not [CalculatedTraceAvailable] then "Calculated-shift trace unavailable: " & [CalculatedTraceError]
        else if [CalculatedShiftRows] = 0 then "Rules reached; no calculated shift rows"
        else if [PositiveShiftRows] = 0 then "Calculated shifts reached; all EffectiveShiftHrs are zero"
        else "Positive ResDayShift candidate; subject to publication checks",
        type text),
    Reordered = Table.ReorderColumns(TraceOutcome, {
        "EmployeeID", "Facility-Abbrev", "BaseName", "SourceRole", "SourceRoleKey",
        "RoleMappingMatches", "MappedRosterRole", "RoleMappingDefinitionRows",
        "DistinctRoleGroupCount", "MappedRoleGroup", "RoleMappingIssue", "Role",
        "Issue", "ContractIssue", "EmploymentType", "Contracted FN Hours",
        "PreparedMappedRoleRows", "MappedRoleCollision", "EligibleRows", "RuleRows",
        "ConfiguredShiftRows", "CalculatedTraceAvailable", "CalculatedTraceError",
        "CalculatedShiftRows", "PositiveShiftRows", "PositiveDates", "PositiveEffectiveHours", "TraceOutcome"
    }),
    Sorted = Table.Sort(Reordered, {
        {"EmployeeID", Order.Ascending},
        {"Facility-Abbrev", Order.Ascending},
        {"SourceRole", Order.Ascending}
    })
in
    Sorted;

// Query: AvailabilityCheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared AvailabilityCheckResult = (name as text, evaluate as function) as record =>
    let
        Evaluation = try evaluate(),
        HasError = Evaluation[HasError],
        Failures = if HasError then null else Evaluation[Value],
        Status = if HasError then "Fail"
            else if Failures = 0 then "Pass"
            else "Fail",
        MessageAttempt = if HasError then try Evaluation[Error][Message] else null,
        Details = if not HasError then null
            else if MessageAttempt[HasError] then "Evaluation failed"
            else MessageAttempt[Value],
        Result = [
            Check = name,
            Status = Status,
            Failures = Failures,
            Details = Details
        ]
    in
        Result;

// Query: ReconciledWorkers_CHECK
// Purpose: Gate worker outputs using identity checks without evaluating shift intervals.
shared ReconciledWorkers_CHECK = let
    Workers = ReconciledWorkers_Eligible,
    InvalidRoleDefinitions = Table.SelectRows(WorkerRoleMappings_Prepare, each
        [RoleMappingIssue] <> null),
    RelevantRoleMappingIssues = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit and [SourceRoleKey] <> null and [RoleMappingIssue] <> null),
    RelevantContractIssues = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit and [Role] <> null and [ContractIssue] <> null),
    DuplicateCount = (rows as table, keys as list) as number =>
        let
            RowCount = Table.RowCount(rows),
            DistinctRows = Table.Distinct(rows, keys),
            DistinctRowCount = Table.RowCount(DistinctRows),
            Duplicates = RowCount - DistinctRowCount
        in
            Duplicates,
    CheckUniqueWorkerKeys = () as number =>
        DuplicateCount(Workers, {"EmployeeID", "Facility-Abbrev", "Role"}),
    CheckPreferredRoleKeys = () as number =>
        DuplicateCount(Workers, {"EmployeeID", "Facility-Abbrev"}),
    CheckRoleDefinitions = () as number =>
        Table.RowCount(InvalidRoleDefinitions),
    CheckWorkerRoleMappings = () as number =>
        Table.RowCount(RelevantRoleMappingIssues),
    CheckRelevantContracts = () as number =>
        Table.RowCount(RelevantContractIssues),
    CheckPublishedNames = () as number =>
        DuplicateCount(Workers, {"Name"}),
    Checks = {
        AvailabilityCheckResult("Unique worker keys", CheckUniqueWorkerKeys),
        AvailabilityCheckResult("One preferred role per employee and facility", CheckPreferredRoleKeys),
        AvailabilityCheckResult("Valid LINK Roles definitions", CheckRoleDefinitions),
        AvailabilityCheckResult("Current worker roles map through LINK Roles", CheckWorkerRoleMappings),
        AvailabilityCheckResult("Valid relevant worker contracts", CheckRelevantContracts),
        // Names alone must remain unique because the existing master list uses that key.
        AvailabilityCheckResult("Unambiguous published names", CheckPublishedNames)
    },
    CheckTable = Table.FromRecords(Checks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    // Small scalar check rows are reused by staff and shift publication within an evaluation.
    BufferedChecks = Table.Buffer(CheckTable)
in
    BufferedChecks;

// Query: AvailabilityRules_TEST
// Purpose: Exercise production helpers against fixed expected answers without opening source workbooks.
// Notes: Synthetic datetimes, hours and reason labels only; this is also a required publication check.
shared AvailabilityRules_TEST = let
    Origin = #datetime(2026, 7, 22, 0, 0, 0),
    At = (hour as number) as datetime =>
        let
            Seconds = hour * 3600,
            Offset = #duration(0, 0, 0, Seconds),
            Timestamp = Origin + Offset
        in
            Timestamp,
    PairToInterval = (pair as list) as record =>
        let
            StartHour = pair{0},
            EndHour = pair{1},
            Start = At(StartHour),
            End = At(EndHour),
            Interval = [Start = Start, End = End]
        in
            Interval,
    Windows = (pairs as list) as list =>
        let
            Intervals = List.Transform(pairs, PairToInterval),
            UnionedIntervals = AvailabilityUnion(Intervals)
        in
            UnionedIntervals,
    Cases = #table(type table [Case = text, StartHour = number, EndHour = number, Base = list,
        Unavail = list, Leave = list, Allowance = number, ExpectedDayHours = number, ExpectedEffective = number], {
        {"No records", 6, 14, {{0, 48}}, {}, {}, 7.6, 0, 7.6},
        {"Short UNAVAIL preserves AM residual", 6, 14, {{0, 48}}, {{13, 15}}, {}, 7.6, 0, 6.6},
        {"Same UNAVAIL preserves PM residual", 14, 22, {{0, 48}}, {{13, 15}}, {}, 7.6, 0, 6.6},
        {"UNAVAIL alone never blocks other shifts", 14, 22, {{0, 48}}, {{6, 14}}, {}, 7.6, 0, 7.6},
        {"UNAVAIL exhausts its own shift", 6, 14, {{0, 48}}, {{6, 14}}, {}, 7.6, 0, 0},
        {"UNAVAIL excluded from daily leave tally", 14, 22, {{0, 48}}, {{6, 10}}, {{10, 14}}, 7.6, 4, 7.6},
        {"10-15 leave AM", 6, 14, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 3.6},
        {"10-15 leave PM", 14, 22, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 6.6},
        {"10-15 leave NIGHT", 22, 30, {{0, 48}}, {}, {{10, 15}}, 7.6, 5, 7.6},
        {"Full-day leave blocks AM", 6, 14, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Full-day leave blocks PM", 14, 22, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Full-day leave blocks entire dated NIGHT", 22, 30, {{0, 48}}, {}, {{8, 16}}, 7.6, 8, 0},
        {"Duplicate four-hour leave counts once", 6, 14, {{0, 48}}, {}, {{6, 10}, {6, 10}}, 7.6, 4, 3.6},
        {"Separate leave periods reach threshold", 14, 22, {{0, 48}}, {}, {{6, 10}, {12, 16}}, 7.6, 8, 0},
        {"Cross-category overlap deducted once", 6, 14, {{0, 48}}, {{8, 12}}, {{10, 14}}, 7.6, 4, 1.6},
        {"Exact midnight end does not affect next day", 30, 38, {{0, 48}}, {}, {{0, 24}}, 7.6, 0, 7.6},
        {"AVAIL window caps residual hours", 6, 14, {{9, 13}}, {}, {{10, 11}}, 7.6, 1, 3},
        {"AVAIL does not cover this shift", 14, 22, {{9, 13}}, {}, {}, 7.6, 0, 0},
        {"Full 7.5-clock-hour shift gets allowance", 22.5, 30, {{0, 48}}, {}, {}, 7.6, 0, 7.6},
        {"Changed Settings allowance", 6, 14, {{0, 48}}, {}, {{10, 15}}, 8, 5, 4},
        {"Exactly ShiftDuration blocks the day", 22, 30, {{0, 48}}, {}, {{6, 13.6}}, 7.6, 7.6, 0},
        {"Midnight-to-midnight UNAVAIL does not block NIGHT after midnight", 22, 30, {{0, 48}}, {{0, 24}}, {}, 7.6, 0, 5.6}
    }),
    CaseRecords = Table.ToRecords(Cases),
    EvaluateCase = (testCase as record) as number =>
        let
            // Raw leave rows feed daily totals so the duplicate-handling test is meaningful.
            RawLeave = List.Transform(testCase[Leave], PairToInterval),
            DayTotals = AvailabilityLeaveDayTotals(RawLeave),
            ShiftStart = At(testCase[StartHour]),
            ShiftEnd = At(testCase[EndHour]),
            ShiftDate = Date.From(ShiftStart),
            MatchingDay = Table.SelectRows(DayTotals, each [Date] = ShiftDate),
            HasMatchingDay = not Table.IsEmpty(MatchingDay),
            DayHours = if HasMatchingDay then MatchingDay{0}[DistinctLeaveHours] else 0,
            Blocked = DayHours >= testCase[Allowance],
            BaselineWindows = Windows(testCase[Base]),
            UnavailabilityWindows = Windows(testCase[Unavail]),
            LeaveWindows = Windows(testCase[Leave]),
            Measurement = AvailabilityShiftMeasurement(
                BaselineWindows,
                UnavailabilityWindows,
                LeaveWindows,
                ShiftStart,
                ShiftEnd,
                Blocked
            ),
            Effective = AvailabilityResidualHours(
                Measurement[AvailableHours],
                Measurement[CombinedAbsenceHoursInShift],
                Measurement[FullShift],
                Blocked,
                testCase[Allowance]
            ),
            DayHoursDifference = Number.Abs(DayHours - testCase[ExpectedDayHours]),
            EffectiveHoursDifference = Number.Abs(Effective - testCase[ExpectedEffective]),
            DayHoursMatch = DayHoursDifference < 0.00000001,
            EffectiveHoursMatch = EffectiveHoursDifference < 0.00000001,
            Passed = DayHoursMatch and EffectiveHoursMatch
                and Measurement[SegmentsValid] and Measurement[Disjoint],
            Failures = if Passed then 0 else 1
        in
            Failures,
    BuildCaseCheck = (testCase as record) as record =>
        let
            Evaluate = () as number => EvaluateCase(testCase),
            Check = AvailabilityCheckResult(testCase[Case], Evaluate)
        in
            Check,
    RuleChecks = List.Transform(CaseRecords, BuildCaseCheck),
    CheckClassification = () as number =>
        let
            Unavailability = AvailabilityRecordKind("UNAVAIL - Unpaid Leave"),
            Availability = AvailabilityRecordKind("AVAIL"),
            Leave = AvailabilityRecordKind("Maternity Lve Unpaid"),
            Unmapped = AvailabilityRecordKind("Unmapped reason"),
            Blank = AvailabilityRecordKind(null),
            Passed = Unavailability = "UNAVAIL"
                and Availability = "AVAIL"
                and Leave = "Leave"
                and Unmapped = "Unknown"
                and Blank = "Unknown",
            Failures = if Passed then 0 else 1
        in
            Failures,
    CheckDailyReset = () as number =>
        let
            RestrictedWindows = Windows({{9, 13}}),
            CalendarStart = At(0),
            CalendarEnd = At(48),
            Baseline = AvailabilityDailyWindows(RestrictedWindows, CalendarStart, CalendarEnd),
            FirstDayStart = At(0),
            FirstDayEnd = At(24),
            SecondDayStart = At(24),
            SecondDayEnd = At(48),
            FirstDayIntervals = AvailabilityClipIntervals(Baseline, FirstDayStart, FirstDayEnd),
            SecondDayIntervals = AvailabilityClipIntervals(Baseline, SecondDayStart, SecondDayEnd),
            FirstDayHours = AvailabilityIntervalHours(FirstDayIntervals),
            SecondDayHours = AvailabilityIntervalHours(SecondDayIntervals),
            Passed = FirstDayHours = 4 and SecondDayHours = 24,
            Failures = if Passed then 0 else 1
        in
            Failures,
    OtherChecks = {
        AvailabilityCheckResult("Classification separates leave and UNAVAIL", CheckClassification),
        AvailabilityCheckResult("AVAIL resets independently on next date", CheckDailyReset)
    },
    AllChecks = RuleChecks & OtherChecks,
    CheckTable = Table.FromRecords(AllChecks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    BufferedChecks = Table.Buffer(CheckTable)
in
    BufferedChecks;

// Query: ResDayShift_CHECK
// Purpose: Gate shift publication using input checks and materialized scalar measurement checks.
shared ResDayShift_CHECK = let
    Allowance = #"EXTRACT EffectiveShiftHrs",
    // Scalar columns only; reuse these rows for every shift check in this evaluation.
    Calculated = Table.Buffer(ResDayShift_Calculated),
    DuplicateCount = (rows as table, keys as list) as number =>
        let
            RowCount = Table.RowCount(rows),
            DistinctRows = Table.Distinct(rows, keys),
            DistinctRowCount = Table.RowCount(DistinctRows),
            Duplicates = RowCount - DistinctRowCount
        in
            Duplicates,
    CheckKnownRules = () as number =>
        let
            FailedRules = Table.SelectRows(AvailabilityRules_TEST, each [Status] <> "Pass"),
            FailureCount = Table.RowCount(FailedRules)
        in
            FailureCount,
    CheckUniqueLeaveDays = () as number =>
        DuplicateCount(WorkerLeaveDays, {"EmployeeID", "Facility-Abbrev", "Date"}),
    CheckRoleCoverage = () as number =>
        let
            MissingRoles = Table.SelectRows(WorkerShiftRoleCoverage_DIAGNOSTICS,
                each [MissingShiftWindow]),
            FailureCount = Table.RowCount(MissingRoles)
        in
            FailureCount,
    CheckShiftEndpoints = () as number =>
        let
            ShiftWindowCount = Table.RowCount(AvailabilityShiftWindows),
            CanValidate = Allowance > 0 and ShiftWindowCount > 0,
            InvalidEndpoints = if CanValidate then
                Table.SelectRows(Calculated, each [ShiftEnd] <= [ShiftStart])
                else #table({}, {}),
            InvalidEndpointCount = if CanValidate then Table.RowCount(InvalidEndpoints) else 1
        in
            InvalidEndpointCount,
    CheckInvalidSourceRecords = () as number =>
        Table.RowCount(AvailabilityRecords_Invalid),
    CheckUniqueWorkerShifts = () as number =>
        DuplicateCount(Calculated, {"EmployeeID", "Facility-Abbrev", "Role", "Date", "Shift"}),
    CheckSegmentBounds = () as number =>
        let
            InvalidSegments = Table.SelectRows(Calculated, each not [SegmentsValid]),
            FailureCount = Table.RowCount(InvalidSegments)
        in
            FailureCount,
    MeasurementIsInvalid = (row as record) as logical =>
        let
            ShiftClockDuration = row[ShiftEnd] - row[ShiftStart],
            ShiftClockHours = Duration.TotalHours(ShiftClockDuration),
            ResidualAllowance = Allowance - row[CombinedAbsenceHoursInShift],
            CappedAtAvailableHours = List.Min({row[AvailableHours], ResidualAllowance}),
            NonNegativeResidual = List.Max({0, CappedAtAvailableHours}),
            ExpectedEffectiveHours = if row[FullDayLeaveBlocked] or row[AvailableHours] <= 0 then 0
                else if row[FullShift] then Allowance
                else NonNegativeResidual,
            EffectiveDifference = Number.Abs(row[EffectiveShiftHrs] - ExpectedEffectiveHours),
            ExpectedFullDayBlock = row[DistinctLeaveHours] >= Allowance,
            IsNotDisjoint = not row[Disjoint],
            HasNegativeAvailableHours = row[AvailableHours] < 0,
            ExceedsShiftClock = row[AvailableHours] > ShiftClockHours + 0.00000001,
            HasEffectiveMismatch = EffectiveDifference > 0.00000001,
            HasBlockMismatch = row[FullDayLeaveBlocked] <> ExpectedFullDayBlock,
            BlockedRowHasHours = row[FullDayLeaveBlocked]
                and (row[EffectiveShiftHrs] <> 0 or row[AvailableHours] <> 0),
            DoubleCountsAbsence = row[CombinedAbsenceHoursInShift]
                > row[LeaveHoursInShift] + row[UnavailHoursInShift] + 0.00000001,
            EffectiveBelowZero = row[EffectiveShiftHrs] < 0,
            EffectiveAboveAllowance = row[EffectiveShiftHrs] > Allowance,
            IsInvalid = IsNotDisjoint or HasNegativeAvailableHours or ExceedsShiftClock
                or HasEffectiveMismatch or HasBlockMismatch or BlockedRowHasHours
                or DoubleCountsAbsence or EffectiveBelowZero or EffectiveAboveAllowance
        in
            IsInvalid,
    CheckMeasurements = () as number =>
        let
            InvalidMeasurements = Table.SelectRows(Calculated, MeasurementIsInvalid),
            FailureCount = Table.RowCount(InvalidMeasurements)
        in
            FailureCount,
    CheckPublishedScope = () as number =>
        let
            ApprovedRoles = {"RN", "AIN", "EN"},
            InvalidScope = Table.SelectRows(Calculated, each
                [#"Facility-Abbrev"] <> Unit or not List.Contains(ApprovedRoles, [Role])),
            FailureCount = Table.RowCount(InvalidScope)
        in
            FailureCount,
    Checks = {
        AvailabilityCheckResult("Known-answer availability rules", CheckKnownRules),
        AvailabilityCheckResult("Unique worker/date leave totals", CheckUniqueLeaveDays),
        AvailabilityCheckResult("Eligible worker roles have roster shifts", CheckRoleCoverage),
        AvailabilityCheckResult("Settings shift endpoints", CheckShiftEndpoints),
        AvailabilityCheckResult("Invalid source records", CheckInvalidSourceRecords),
        AvailabilityCheckResult("Unique worker shift rows", CheckUniqueWorkerShifts),
        AvailabilityCheckResult("Segments obey availability and exclusions", CheckSegmentBounds),
        AvailabilityCheckResult("Disjoint segments and reconciled hours", CheckMeasurements),
        AvailabilityCheckResult("Published facility and roles", CheckPublishedScope)
    },
    ShiftCheckTable = Table.FromRecords(Checks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Result = Table.Combine({ReconciledWorkers_CHECK, ShiftCheckTable}),
    // Buffer the small check output; buffers are scoped to an evaluation, not shared across refreshes.
    BufferedResult = Table.Buffer(Result)
in
    BufferedResult;

// Query: ResDayShift
// Purpose: Publish validated worker/date/shift availability with reconciled worker and facility identifiers.
// Output: The original seven columns followed by ID (reconciled EmployeeID) and Facility-Abbrev, both text.
// Notes: CapacityDistrib still assigns one availability unit per row; partial-hour weighting is deferred.
shared ResDayShift = let
    Failures = Table.SelectRows(ResDayShift_CHECK, each [Status] <> "Pass"),
    FailureNames = Text.Combine(Failures[Check], "; "),
    FailureMessage = "Required availability checks failed: " & FailureNames
        & ". Inspect ResDayShift_CHECK.",
    Checked = if Table.IsEmpty(Failures) then ResDayShift_Calculated
        else error Error.Record("ResDayShift validation",
            FailureMessage, Failures),
    // Preserve the original seven columns in order and append identifiers for future downstream matching.
    Positive = Table.SelectRows(Checked, each [EffectiveShiftHrs] > 0),
    Identified = Table.RenameColumns(Positive, {{"EmployeeID", "ID"}}),
    Output = Table.SelectColumns(Identified, {"Role", "Week", "Day", "Shift", "EffectiveShiftHrs", "Date", "Name", "ID", "Facility-Abbrev"}),
    Typed = Table.TransformColumnTypes(Output, {{"Role", type text}, {"Week", Int64.Type}, {"Day", type text},
        {"Shift", type text}, {"EffectiveShiftHrs", type number}, {"Date", type date}, {"Name", type text},
        {"ID", type text}, {"Facility-Abbrev", type text}}),
    Sorted = Table.Sort(Typed,
        {{"Week", Order.Ascending}, {"Date", Order.Ascending}, {"Role", Order.Ascending},
         {"Name", Order.Ascending}, {"Shift", Order.Ascending}})
in
    Sorted;

// Query: AllocationRoleShares_CHECK
// Purpose: Expose role-share uniqueness and per-worker reconciliation before capacity capping.
shared AllocationRoleShares_CHECK = let
    Shares = AllocationRoleShares,
    WorkerTotals = Table.Group(Shares, {"EmployeeID", "Facility-Abbrev"},
        {{"ShareSum", each List.Sum([RoleShare]), type number}}),
    CheckUniqueShareKeys = () as number =>
        let
            RowCount = Table.RowCount(Shares),
            DistinctShares = Table.Distinct(Shares, {"EmployeeID", "Facility-Abbrev", "Role"}),
            DistinctRowCount = Table.RowCount(DistinctShares),
            DuplicateCount = RowCount - DistinctRowCount
        in
            DuplicateCount,
    CheckShareTotals = () as number =>
        let
            InvalidTotals = Table.SelectRows(WorkerTotals, each
                Number.Abs([ShareSum] - 1) > 0.00000001),
            FailureCount = Table.RowCount(InvalidTotals)
        in
            FailureCount,
    Checks = {
        AvailabilityCheckResult("Unique allocation role-share keys", CheckUniqueShareKeys),
        AvailabilityCheckResult("Allocation role shares sum to one", CheckShareTotals)
    },
    CheckTable = Table.FromRecords(Checks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])
in
    CheckTable;

// Query: MultiRoles
// Purpose: Publish role-specific roster names for employees allocated to more than one approved role.
// Output: Name, MultiRole and unrounded RoleShare; the null legacy column only keeps MutliRoleCheck's Name import working.
shared MultiRoles = let
    MultiRoleShares = Table.SelectRows(AllocationRoleShares, each [RoleCount] > 1),
    Selected = Table.SelectColumns(MultiRoleShares, {"RosterName", "Role", "RoleShare"}),
    Named = Table.RenameColumns(Selected, {{"RosterName", "Name"}, {"Role", "MultiRole"}}),
    // MutliRoleCheck currently types this obsolete field but then selects Name only; never use it as a share.
    LegacyImportColumn = Table.AddColumn(Named, "AINC4HrsAvilPref", each null, type nullable number),
    Output = Table.SelectColumns(LegacyImportColumn,
        {"Name", "AINC4HrsAvilPref", "MultiRole", "RoleShare"}),
    Sorted = Table.Sort(Output, {{"Name", Order.Ascending}, {"MultiRole", Order.Ascending}})
in
    Sorted;

// Query: RoleResDayAvailabilityCapped
// Purpose: Apply the existing shift-count cap using unrounded allocation shares for multi-role employees.
// Notes: Join by employee and facility rather than name; workers absent from allocation retain a full single-role cap.
shared RoleResDayAvailabilityCapped = let
    Failures = Table.SelectRows(AllocationRoleShares_CHECK, each [Status] <> "Pass"),
    ValidatedShares = if Table.IsEmpty(Failures) then AllocationRoleShares
        else error Error.Record("Allocation role shares", "Role-share checks failed.", Failures),
    ShiftCounts = Table.Group(ResDayShift, {"ID", "Facility-Abbrev", "Role", "Name"},
        {{"Count", each Table.RowCount(_), Int64.Type}}),
    SettingsMaximum = #"EXTRACT MaxAvailability",
    WithMaximum = Table.AddColumn(ShiftCounts, "StdRosterDays", each SettingsMaximum, Int64.Type),
    SelectedAllocationRoleCounts = Table.SelectColumns(ValidatedShares,
        {"EmployeeID", "Facility-Abbrev", "RoleCount"}),
    AllocationRoleCounts = Table.Distinct(SelectedAllocationRoleCounts),
    JoinedRoleCounts = Table.NestedJoin(WithMaximum, {"ID", "Facility-Abbrev"},
        AllocationRoleCounts, {"EmployeeID", "Facility-Abbrev"}, "AllocationRoles", JoinKind.LeftOuter),
    JoinedShares = Table.NestedJoin(JoinedRoleCounts, {"ID", "Facility-Abbrev", "Role"},
        ValidatedShares, {"EmployeeID", "Facility-Abbrev", "Role"}, "AllocationShare", JoinKind.LeftOuter),
    HasInvalidShareMatch = (row as record) as logical =>
        let
            ShareMatchCount = Table.RowCount(row[AllocationShare]),
            HasNoShare = ShareMatchCount = 0,
            HasDuplicateShares = ShareMatchCount > 1,
            HasAllocationRoles = not Table.IsEmpty(row[AllocationRoles]),
            HasMultipleAllocationRoles = if HasAllocationRoles
                then row[AllocationRoles]{0}[RoleCount] > 1
                else false,
            MissingRequiredShare = HasNoShare and HasMultipleAllocationRoles,
            IsInvalid = HasDuplicateShares or MissingRequiredShare
        in
            IsInvalid,
    MissingOrDuplicateShares = Table.SelectRows(JoinedShares, HasInvalidShareMatch),
    InvalidShareDetails = Table.SelectColumns(MissingOrDuplicateShares,
        {"ID", "Facility-Abbrev", "Role", "Name"}),
    ValidatedMatches = if Table.IsEmpty(MissingOrDuplicateShares) then JoinedShares
        else error Error.Record("Availability role share", "A multi-role worker lacks one matching allocation share.",
            InvalidShareDetails),
    ResolveRoleShare = (matches as table) as number =>
        let
            HasMatch = not Table.IsEmpty(matches),
            RoleShare = if HasMatch then matches{0}[RoleShare] else 1
        in
            RoleShare,
    RoleShares = Table.AddColumn(ValidatedMatches, "RoleShare", each
        ResolveRoleShare([AllocationShare]), type number),
    // Preserve the prior cap rule: only counts above the Settings maximum are apportioned by role share.
    CalculateAvailabilityCap = (row as record) as number =>
        let
            Count = Number.From(row[Count]),
            StandardRosterDays = Number.From(row[StdRosterDays]),
            AvailabilityCapped = if Count <= StandardRosterDays then Count
                else StandardRosterDays * row[RoleShare]
        in
            AvailabilityCapped,
    Capped = Table.AddColumn(RoleShares, "AvailabilityCapped",
        CalculateAvailabilityCap, type number),
    Result = Table.SelectColumns(Capped, {"Count", "StdRosterDays", "AvailabilityCapped", "Name", "Role"})
in
    Result;

// Query: RoleAvailabilityCapped
// Purpose: Sum the capped worker availability to one published value per role.
shared RoleAvailabilityCapped = let
    Source = RoleResDayAvailabilityCapped,
    SumAvailability = (rows as table) as number =>
        let
            AvailabilityValues = rows[AvailabilityCapped],
            AvailabilityTotal = List.Sum(AvailabilityValues)
        in
            AvailabilityTotal,
    GroupedByRole = Table.Group(Source, {"Role"},
        {{"AvailabilityCapped", SumAvailability, type number}})
in
    GroupedByRole;

// Query: Availability-StaffList
// Purpose: Publish eligible staff identity and fortnightly contract fields, including workers with no available shifts.
// Output: One row per reconciled preferred employee/resource with facility, employment type and Contracted FN Hours.
shared #"Availability-StaffList" = let
    Failures = Table.SelectRows(ReconciledWorkers_CHECK, each [Status] <> "Pass"),
    FailureNames = Text.Combine(Failures[Check], "; "),
    FailureMessage = "Required worker identity checks failed: " & FailureNames
        & ". Inspect ReconciledWorkers_CHECK, WorkerIdentityFailures_DIAGNOSTICS and WorkerRoleMappings_DIAGNOSTICS.",
    Workers = if Table.IsEmpty(Failures) then ReconciledWorkers_Eligible
        else error Error.Record("Availability staff validation",
            FailureMessage, Failures),
    SelectedOutput = Table.SelectColumns(Workers,
        {"EmployeeID", "Facility-Abbrev", "Name", "Role", "PreferredRole", "EmploymentType", "Contracted FN Hours"}),
    Output = Table.Distinct(SelectedOutput),
    Sorted = Table.Sort(Output,
        {{"Facility-Abbrev", Order.Ascending}, {"PreferredRole", Order.Ascending}, {"Name", Order.Ascending}})
in
    Sorted;

// Query: IMPORT CentriSyncPaths
// Purpose: Read the fixed public-machine path mapping used by the workbook resolver.
shared #"IMPORT CentriSyncPaths" =
    let
        SourcePath = "C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx",
        SourceBinary = File.Contents(SourcePath),
        Navigation = Excel.Workbook(SourceBinary, null, true)
    in
        Navigation;

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook to its unit folder using FilePathUrl and CentriSyncPaths.
// Output: Existing Variable Name/Value interface plus the authoritative Unit Root.
shared UnitL1PathTABLE = let
    NormalizePath = (value as nullable text) as nullable text =>
        let
            Trimmed = if value = null then null else Text.Trim(value),
            BackslashSeparated = if Trimmed = null then null else Text.Replace(Trimmed, "/", "\"),
            WithoutTrailingSeparator = if BackslashSeparated = null then null
                else Text.TrimEnd(BackslashSeparated, "\")
        in
            WithoutTrailingSeparator,
    CurrentWorkbook = Excel.CurrentWorkbook(),
    PathInput = CurrentWorkbook{[Name = "FilePathUrl"]}[Content],
    // Excel exposes this workbook's existing FilePathUrl named cell as Column1.
    // Also accept the standard FilePath table through the same resolver.
    HasFilePathColumn = Table.HasColumns(PathInput, {"FilePath"}),
    PathInputColumns = Table.ColumnNames(PathInput),
    IsNamedCellShape = PathInputColumns = {"Column1"},
    PathColumn = if HasFilePathColumn then Table.SelectColumns(PathInput, {"FilePath"})
        else if IsNamedCellShape then
            Table.RenameColumns(PathInput, {{"Column1", "FilePath"}})
        else error "FilePathUrl must be a one-row FilePath table or a single named cell.",
    PathRowCount = Table.RowCount(PathColumn),
    PathRows = if PathRowCount = 1 then PathColumn
        else error "FilePathUrl must contain exactly one path value.",
    PathValue = PathRows{0}[FilePath],
    PathText = Text.From(PathValue),
    RawPath = NormalizePath(PathText),
    NonBlankPath = if RawPath <> null and RawPath <> "" then RawPath
        else error "FilePathUrl is blank. Save and recalculate this workbook.",
    // CELL("filename") includes [workbook.xlsx]sheet; discard the sheet suffix.
    HasWorkbookBrackets = Text.Contains(NonBlankPath, "["),
    PathBeforeWorkbook = if HasWorkbookBrackets then Text.BeforeDelimiter(NonBlankPath, "[") else null,
    BracketedWorkbookName = if HasWorkbookBrackets
        then Text.BetweenDelimiters(NonBlankPath, "[", "]")
        else null,
    WorkbookPath = if HasWorkbookBrackets then
        PathBeforeWorkbook & BracketedWorkbookName
        else NonBlankPath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FileNameComparison = Comparer.OrdinalIgnoreCase(InputFileName, "Capacity-ShiftAvailability.xlsx"),
    IsExpectedWorkbook = FileNameComparison = 0,
    FilePath = if IsExpectedWorkbook then WorkbookPath
        else error "FilePathUrl must identify Capacity-ShiftAvailability.xlsx.",
    MappingSource = #"IMPORT CentriSyncPaths",
    MappingTable = MappingSource{[Item = "CentriSyncPaths", Kind = "Table"]}[Data],
    SelectedMapping = Table.SelectColumns(MappingTable,
        {"SharepointRootUrl", "SyncedFolderRootPath"}),
    NormalizeAnyPath = (value as any) as nullable text =>
        let
            Converted = Text.From(value),
            Normalized = NormalizePath(Converted)
        in
            Normalized,
    Mapping = Table.TransformColumns(SelectedMapping, {
        {"SharepointRootUrl", NormalizeAnyPath, type nullable text},
        {"SyncedFolderRootPath", NormalizeAnyPath, type nullable text}}),
    SharePoint = Table.AddColumn(Mapping, "MatchRoot", each [SharepointRootUrl], type nullable text),
    Documents = Table.AddColumn(Mapping, "MatchRoot", each if [SharepointRootUrl] = null then null
        else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    Local = Table.AddColumn(Mapping, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    CombinedCandidates = Table.Combine({SharePoint, Documents, Local}),
    Candidates = Table.AddColumn(CombinedCandidates, "RootLength",
        each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    // The folder boundary prevents one mapping from matching a similarly prefixed folder.
    FilePathLength = Text.Length(FilePath),
    IsMatchingCandidate = (row as record) as logical =>
        let
            MatchRoot = row[MatchRoot],
            SyncedRoot = row[SyncedFolderRootPath],
            HasMatchRoot = MatchRoot <> null and MatchRoot <> "",
            HasSyncedRoot = SyncedRoot <> null and SyncedRoot <> "",
            StartsWithRoot = if HasMatchRoot
                then Text.StartsWith(FilePath, MatchRoot, Comparer.OrdinalIgnoreCase)
                else false,
            EndsAtRoot = FilePathLength = row[RootLength],
            BoundaryCharacter = if EndsAtRoot then null
                else Text.Range(FilePath, row[RootLength], 1),
            HasFolderBoundary = EndsAtRoot or BoundaryCharacter = "\",
            IsMatch = HasMatchRoot and HasSyncedRoot and StartsWithRoot and HasFolderBoundary
        in
            IsMatch,
    Matching = Table.SelectRows(Candidates, IsMatchingCandidate),
    Ordered = Table.Sort(Matching, {{"RootLength", Order.Descending}}),
    Longest = if Table.IsEmpty(Ordered) then error "FilePathUrl has no CentriSyncPaths mapping." else Ordered{0},
    LongestRootLength = Longest[RootLength],
    SameLength = Table.SelectRows(Ordered, each [RootLength] = LongestRootLength),
    SameLengthDestinations = SameLength[SyncedFolderRootPath],
    DistinctDestinations = List.Distinct(SameLengthDestinations, Comparer.OrdinalIgnoreCase),
    DestinationCount = List.Count(DistinctDestinations),
    Best = if DestinationCount = 1 then Longest
        else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    RelativeWithSeparator = Text.Range(FilePath, Best[RootLength]),
    Relative = Text.TrimStart(RelativeWithSeparator, "\"),
    LocalWorkbook = if Relative = "" then Best[SyncedFolderRootPath]
        else Best[SyncedFolderRootPath] & "\" & Relative,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbook, "\", {0, RelativePosition.FromEnd}),
    AllFolderSegments = Text.Split(WorkbookFolder, "\"),
    FolderSegments = List.Select(AllFolderSegments, each _ <> ""),
    UnitsIndex = List.PositionOf(FolderSegments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    CalcIndex = List.PositionOf(FolderSegments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    FolderSegmentCount = List.Count(FolderSegments),
    HasUnitsSegment = UnitsIndex >= 0,
    HasExpectedCalculationSegment = CalcIndex = UnitsIndex + 2,
    CalculationFolderIsLast = FolderSegmentCount = CalcIndex + 1,
    ValidFolder = HasUnitsSegment and HasExpectedCalculationSegment and CalculationFolderIsLast,
    RootPath = if ValidFolder then WorkbookFolder else error "Workbook must be under UNITS/<unit>/2. Calculations.",
    UnitRoot = Text.BeforeDelimiter(RootPath, "\", {0, RelativePosition.FromEnd}),
    CareIndex = List.PositionOf(FolderSegments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    HasClientSegment = CareIndex >= 0 and FolderSegmentCount > CareIndex + 1,
    Client = if HasClientSegment then FolderSegments{CareIndex + 1} else null,
    Unit = FolderSegments{UnitsIndex + 1},
    OutputRows = {
        {"UserName", null},
        {"Root Path", RootPath},
        {"Unit Root", UnitRoot},
        {"Client", Client},
        {"Date", null},
        {"Unit", Unit},
        {"FileName", InputFileName}
    },
    Output = #table({"Variable Name", "Value"}, OutputRows),
    BufferedOutput = Table.Buffer(Output)
in
    BufferedOutput;

// Query: Unit
// Purpose: Expose the resolved workbook folder name as the canonical unit filter value.
shared Unit =
let
    UnitRows = Table.SelectRows(UnitL1PathTABLE, each [Variable Name] = "Unit"),
    UnitRowCount = Table.RowCount(UnitRows),
    UnitName = if UnitRowCount = 1 then UnitRows{0}[Value]
        else error "UnitL1PathTABLE must identify one unit folder.",
    TrimmedUnitName = if UnitName = null then null else Text.Trim(UnitName),
    ValidatedUnitName = if TrimmedUnitName = null or TrimmedUnitName = ""
        then error "The workbook path did not identify a unit folder."
        else UnitName
in
    ValidatedUnitName;

// Query: FilePath - 2Calculations
// Purpose: Expose the resolved workbook folder to existing calculation imports.
shared #"FilePath - 2Calculations" =
let
    RootPathRows = Table.SelectRows(UnitL1PathTABLE, each [Variable Name] = "Root Path"),
    RootPathRowCount = Table.RowCount(RootPathRows),
    RootPath = if RootPathRowCount = 1 then RootPathRows{0}[Value]
        else error "UnitL1PathTABLE must identify one calculation root path."
in
    RootPath;

// Query: FilePath - 1Input
// Purpose: Build the sibling input folder from the resolved unit root.
shared #"FilePath - 1Input" =
let
    UnitRootRows = Table.SelectRows(UnitL1PathTABLE, each [Variable Name] = "Unit Root"),
    UnitRootRowCount = Table.RowCount(UnitRootRows),
    UnitRoot = if UnitRootRowCount = 1 then UnitRootRows{0}[Value]
        else error "UnitL1PathTABLE must identify one unit root path.",
    InputPath = UnitRoot & "\1. Input"
in
    InputPath;