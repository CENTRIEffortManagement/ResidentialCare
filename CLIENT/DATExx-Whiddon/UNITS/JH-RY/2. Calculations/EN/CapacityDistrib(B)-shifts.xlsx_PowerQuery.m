// Power Query from: CapacityDistrib(B)-shifts.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\AIN\CapacityDistrib(B)-shifts.xlsx
// Extracted: 2026-10-05T07:55:23.444Z

section Section1;

// Query: fnBInputIssues
// Purpose: Attribute input schema, cell and duplicate-key failures without stopping unrelated Resources.
// Inputs: A captured input evaluation, its required columns and join keys.
// Output: Diagnostic detail rows; a null Resource blocks all spare when attribution is impossible.
// Notes: Unmatched allocation rows are reported separately without causing scope-wide spare exclusion.
shared fnBInputIssues = (inputName as text, snapshot as record, requiredColumns as list, keys as list, skipUnmatchedAllocation as logical) as table =>
let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    RoleValue = try Role otherwise null,
    MakeIssue = (status as text, resource as nullable number, period as nullable number, reason as text) as record =>
        [Stage = "B", Check = inputName, Status = status, Resource = resource, Role = RoleValue, Name = null, Date = null, Period = period, Reason = reason,
         Action = if resource = null then "Exclude all spare calculations; retain allocation evidence" else "Exclude spare calculations for this Resource; retain allocation evidence"],
    ReadFailed = snapshot[HasError],
    Raw = if ReadFailed then #table({}, {}) else snapshot[Value],
    MissingColumns = if ReadFailed then requiredColumns else List.Difference(requiredColumns, Table.ColumnNames(Raw)),
    SchemaIssues = if ReadFailed then {MakeIssue("Error", null, null, try snapshot[Error][Message] otherwise "Input evaluation failed")}
        else if not List.IsEmpty(MissingColumns) then {MakeIssue("Error", null, null, "Missing required columns: " & Text.Combine(MissingColumns, ", "))} else {},
    HasSchema = not ReadFailed and List.IsEmpty(MissingColumns),
    UsableSchema = if HasSchema then Raw else #table(requiredColumns, {}),
    RowsToCheck = if skipUnmatchedAllocation then Table.SelectRows(UsableSchema, each (try [Resource] otherwise null) <> null) else UsableSchema,
    ErrorRows = Table.SelectRowsWithErrors(RowsToCheck, requiredColumns),
    CellIssues = List.Transform(Table.ToRecords(ErrorRows), each MakeIssue("Error", try _[Resource] otherwise null, try _[Period] otherwise null, "Required input value contains an error")),
    WithoutErrors = Table.RemoveRowsWithErrors(RowsToCheck, requiredColumns),
    IsMissingKey = (row as record) as logical => List.AnyTrue(List.Transform(keys, (key) =>
        let value = Record.Field(row, key) in value = null or (Value.Is(value, type text) and Text.Trim(value) = ""))),
    MissingKeys = Table.SelectRows(WithoutErrors, IsMissingKey),
    MissingKeyIssues = List.Transform(Table.ToRecords(MissingKeys), each MakeIssue("Fail", try _[Resource] otherwise null, try _[Period] otherwise null, "Required join key is missing")),
    CompleteKeys = Table.SelectRows(WithoutErrors, each not IsMissingKey(_)),
    KeyCounts = Table.Group(CompleteKeys, keys, {{"KeyCount", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.SelectRows(KeyCounts, each [KeyCount] > 1),
    DuplicateIssues = List.Transform(Table.ToRecords(DuplicateKeys), each MakeIssue("Fail", try _[Resource] otherwise null, try _[Period] otherwise null, "Duplicate join key; rows=" & Text.From(_[KeyCount]))),
    Combined = SchemaIssues & CellIssues & MissingKeyIssues & DuplicateIssues,
    // Every issue supplies all fields; preserve required types without nullable-only missing-field substitution.
    Output = Table.FromRecords(Combined, DetailType)
in
    Table.Buffer(Output);

// Query: fnBPreparedInput
// Purpose: Return schema-safe rows for calculations after retaining failures in separate diagnostics.
// Notes: Empty fallback tables are accompanied by Error diagnostics and scope-wide spare exclusion.
shared fnBPreparedInput = (snapshot as record, requiredColumns as list, emptyTable as table) as table =>
let
    Raw = if snapshot[HasError] then emptyTable else snapshot[Value],
    HasSchema = List.IsEmpty(List.Difference(requiredColumns, Table.ColumnNames(Raw))),
    WithSchema = if HasSchema then Raw else emptyTable,
    WithoutErrors = Table.RemoveRowsWithErrors(WithSchema, requiredColumns),
    WithKeys = Table.SelectRows(WithoutErrors, each (if Table.HasColumns(WithoutErrors, {"Resource"}) then [Resource] <> null else true)
        and (if Table.HasColumns(WithoutErrors, {"Period"}) then [Period] <> null else true))
in
    Table.Buffer(WithKeys);

// Query: BContractInput_RESULT
// Purpose: Capture contract import failures for honest diagnostics and Resource-scoped spare exclusion.
shared BContractInput_RESULT = try Table.Buffer(#"IMPORT ResourceContract");

// Query: BOriginalAvailabilityInput_RESULT
// Purpose: Capture the existing A.1 original-availability import without substituting another source.
shared BOriginalAvailabilityInput_RESULT = try Table.Buffer(#"IMPORT AvailabilityOriginal");

// Query: BCappedAvailabilityInput_RESULT
// Purpose: Capture A.2 C# plus the Stage 1 Resource mask and allocation evidence.
shared BCappedAvailabilityInput_RESULT = try Table.Buffer(#"IMPORT ResPeriodAvailabilityCapped(C#)");

// Query: BPriorityInput_RESULT
// Purpose: Capture the true NWD priority table and typed input errors before Resource spare exclusions.
shared BPriorityInput_RESULT = try let
    Source = #"EXTRACT ResPeriodNWDTABLE",
    TypedPriorities = Table.TransformColumnTypes(Source, {{"Resource", Int64.Type}, {"Period", Int64.Type}, {"NWDPriority", Int64.Type}})
in
    Table.Buffer(TypedPriorities);

// Query: BDemandInput_RESULT
// Purpose: Capture demand input failures so surviving allocation evidence can still be published.
shared BDemandInput_RESULT = try Table.Buffer(#"IMPORT Demand");

// Query: BAllocationInput_RESULT
// Purpose: Capture the unchanged actual-allocation interface for calculation-side checks.
shared BAllocationInput_RESULT = try Table.Buffer(ResPeriodAllocationTABLE);

// Query: BOriginalAvailability_Prepare
// Purpose: Prepare usable original-availability rows without discarding the raw import diagnostics.
shared BOriginalAvailability_Prepare = fnBPreparedInput(BOriginalAvailabilityInput_RESULT, {"Role", "Resource", "Period", "Availability"},
    #table(type table [Role = nullable text, Resource = nullable number, Period = nullable number, Availability = nullable number], {}));

// Query: BCappedAvailability_Prepare
// Purpose: Prepare usable raw C# rows before masking; diagnostics and lineage inspect this unmasked view.
shared BCappedAvailability_Prepare = fnBPreparedInput(BCappedAvailabilityInput_RESULT, {"Role", "Resource", "Period", "AvailabilityCapped"},
    #table(type table [Role = nullable text, Resource = nullable number, Period = nullable number, AvailabilityCapped = nullable number], {}));

// Query: BPriority_Prepare
// Purpose: Select usable NWD priorities after capturing typed input failures for nonblocking diagnostics.
shared BPriority_Prepare = let
    PreparedInput = fnBPreparedInput(BPriorityInput_RESULT, {"Resource", "Period", "NWDPriority"},
        #table(type table [Resource = nullable number, Period = nullable number, NWDPriority = nullable number], {})),
    SelectedPriorities = Table.SelectColumns(PreparedInput, {"Resource", "Period", "NWDPriority"}),
    SortedPriorities = Table.Sort(SelectedPriorities, {{"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    Table.Buffer(SortedPriorities);

// Query: BDemand_Prepare
// Purpose: Prepare demand for existing B arithmetic; failures remain visible and block optional spare.
shared BDemand_Prepare = fnBPreparedInput(BDemandInput_RESULT, {"Period", "D", "Role", "Shift", "Facility", "Date"},
    #table(type table [Period = nullable number, D = nullable number, Role = nullable text, Shift = nullable text, Facility = nullable text, Date = nullable date], {}));

// Query: BAllocationFacts_Prepare
// Purpose: Retain usable individual allocation facts for the existing period-allocation arithmetic.
shared BAllocationFacts_Prepare = fnBPreparedInput(BAllocationInput_RESULT, {"Resource", "Period", "Allocation"},
    #table(type table [Resource = nullable number, Period = nullable number, Allocation = nullable number], {}));

// Query: BAllocation_CalculationTABLE
// Purpose: Supply one join row per Resource/Period while preserving all actual rows in ResPeriodAllocationTABLE.
// Notes: Duplicate facts are diagnosed; this grouped view prevents multiplicative calculation joins.
shared BAllocation_CalculationTABLE = let
    GroupedFacts = Table.Group(BAllocationFacts_Prepare, {"Resource", "Period"}, {
        {"Allocation", each List.Sum([Allocation]), type nullable number},
        {"HasAllocationEvidence", each List.AnyTrue(List.Transform([Allocation], each _ <> null and _ > 0)), type logical}
    })
in
    Table.Buffer(GroupedFacts);

// Query: BUnmatchedAllocation_EXCEPTIONS
// Purpose: Flag skipped allocation rows with unusable Resource identity without stopping other Resources.
// Output: The allocation quantity remains visible as evidence; these exceptions do not block all spare calculations.
shared BUnmatchedAllocation_EXCEPTIONS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text, AllocationEvidence = nullable number],
    RoleValue = try Role otherwise null,
    Raw = if BAllocationInput_RESULT[HasError] then #table({"Resource"}, {}) else BAllocationInput_RESULT[Value],
    Rows = if not Table.HasColumns(Raw, {"Resource"}) then {} else Table.ToRecords(Table.SelectRows(Raw, each (try [Resource] otherwise null) = null)),
    Issues = List.Transform(Rows, each [Stage = "B", Check = "Allocation input", Status = if (try _[Resource] otherwise "Error") = "Error" then "Error" else "Fail",
        Resource = null, Role = try _[Role] otherwise RoleValue, Name = try _[Name] otherwise null, Date = try _[Date] otherwise null, Period = try _[Period] otherwise null,
        Reason = "Allocation has no usable Resource match", Action = "Allocation skipped; evidence retained; other Resources continue", AllocationEvidence = try _[Allocation] otherwise null]),
    Output = Table.FromRecords(Issues, DetailType)
in
    Table.Buffer(Output);

// Query: BInputIssues_DETAILS
// Purpose: Collect Resource-local key failures and scope-wide source/schema errors before calculations.
shared BInputIssues_DETAILS = Table.Combine({
    fnBInputIssues("Demand input", BDemandInput_RESULT, {"Period", "D", "Role", "Shift", "Facility", "Date"}, {"Period"}, false),
    fnBInputIssues("Allocation input", BAllocationInput_RESULT, {"Resource", "Period", "Allocation"}, {"Resource", "Period"}, true),
    fnBInputIssues("C# input", BCappedAvailabilityInput_RESULT, {"Role", "Resource", "Period", "AvailabilityCapped"}, {"Resource", "Period"}, false),
    fnBInputIssues("Original availability input", BOriginalAvailabilityInput_RESULT, {"Role", "Resource", "Period", "Availability"}, {"Resource", "Period"}, false),
    fnBInputIssues("Priority input", BPriorityInput_RESULT, {"Resource", "Period", "NWDPriority"}, {"Resource", "Period"}, false),
    fnBInputIssues("Resource contract keys", BContractInput_RESULT, {"Resource", "Effective Shift Cap", "Limit Basis", "Worker Record Status"}, {"Resource"}, false),
    BUnmatchedAllocation_EXCEPTIONS
});

// Query: BContractIssues_DETAILS
// Purpose: Attribute invalid, missing and disallowed fallback caps to affected Resources without inventing caps.
shared BContractIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    RoleValue = try Role otherwise null,
    MakeIssue = (resource as nullable number, reason as text) as record => [Stage = "B", Check = "Resource contracts", Status = "Fail", Resource = resource, Role = RoleValue, Name = null, Date = null, Period = null, Reason = reason,
        Action = if resource = null then "Exclude all spare calculations; retain allocation evidence" else "Exclude spare calculations for this Resource; retain allocation evidence"],
    Contracts = fnBPreparedInput(BContractInput_RESULT, {"Resource", "Effective Shift Cap", "Limit Basis", "Worker Record Status"},
        #table(type table [Resource = nullable number, #"Effective Shift Cap" = nullable number, #"Limit Basis" = nullable text, #"Worker Record Status" = nullable text], {})),
    InvalidRows = Table.SelectRows(Contracts, each [Effective Shift Cap] = null or Number.IsNaN([Effective Shift Cap]) or [Effective Shift Cap] = #infinity or [Effective Shift Cap] = -#infinity or [Effective Shift Cap] <= 0
        or [Effective Shift Cap] <> Number.RoundDown([Effective Shift Cap]) or [Limit Basis] = null or Text.Trim([Limit Basis]) = ""),
    InvalidIssues = List.Transform(Table.ToRecords(InvalidRows), each MakeIssue(_[Resource], "Effective Shift Cap or Limit Basis is missing or invalid")),
    PositiveOriginal = Table.SelectRows(BOriginalAvailability_Prepare, each [Availability] <> null and [Availability] > 0),
    PositiveCapped = Table.SelectRows(BCappedAvailability_Prepare, each [AvailabilityCapped] <> null and [AvailabilityCapped] > 0),
    PositiveResources = Table.Distinct(Table.Combine({Table.SelectColumns(PositiveOriginal, {"Resource"}), Table.SelectColumns(PositiveCapped, {"Resource"})})),
    MissingCaps = Table.NestedJoin(PositiveResources, {"Resource"}, Contracts, {"Resource"}, "Contracts", JoinKind.LeftAnti),
    MissingIssues = List.Transform(MissingCaps[Resource], each MakeIssue(_, "Positive availability has no Resource contract")),
    FallbackResources = Table.SelectRows(Contracts, each [Worker Record Status] = "Not found"),
    InvalidFallback = Table.NestedJoin(Table.Distinct(Table.SelectColumns(PositiveOriginal, {"Resource"})), {"Resource"}, FallbackResources, {"Resource"}, "Fallback", JoinKind.Inner),
    FallbackIssues = List.Transform(InvalidFallback[Resource], each MakeIssue(_, "Missing-worker fallback has positive AIN availability")),
    Output = Table.FromRecords(InvalidIssues & MissingIssues & FallbackIssues, DetailType)
in
    Table.Buffer(Output);

// Query: BMaskIssues_DETAILS
// Purpose: Require explicit upstream Resource permission; missing masks are Error rather than an implied Pass.
shared BMaskIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    RoleValue = try Role otherwise null,
    MakeIssue = (status as text, resource as nullable number, reason as text) as record => [Stage = "B", Check = "Upstream spare permission", Status = status, Resource = resource, Role = RoleValue, Name = null, Date = null, Period = null, Reason = reason,
        Action = if resource = null then "Exclude all spare calculations; retain allocation evidence" else "Exclude spare calculations for this Resource; retain allocation evidence"],
    RequiredMetadata = {"SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence", "OriginalAvailability"},
    ReadIssues = (snapshot as record, inputName as text) as list =>
        let
            Raw = if snapshot[HasError] then #table({}, {}) else snapshot[Value],
            MissingMetadata = List.Difference(RequiredMetadata, Table.ColumnNames(Raw)),
            SchemaIssues = if snapshot[HasError] then {} else if not List.IsEmpty(MissingMetadata) then
                {MakeIssue("Error", null, inputName & " is missing Stage 1 metadata: " & Text.Combine(MissingMetadata, ", "))} else {},
            Rows = if snapshot[HasError] or not List.IsEmpty(MissingMetadata) then {} else Table.ToRecords(Raw),
            HasValidMetadata = (row as record) as logical => try Value.Is(row[SpareCalculationAllowed], type logical)
                and Value.Is(row[HasAllocationEvidence], type logical)
                and (row[SpareIssueReason] = null or Value.Is(row[SpareIssueReason], type text))
                and (row[OriginalAvailability] = null or Value.Is(row[OriginalAvailability], type number)) otherwise false,
            BlockedRows = List.Select(Rows, each not HasValidMetadata(_) or (try _[SpareCalculationAllowed] otherwise false) <> true),
            RowIssues = List.Transform(BlockedRows, each MakeIssue(
                if HasValidMetadata(_) then "Fail" else "Error",
                try _[Resource] otherwise null,
                inputName & ": " & (if not HasValidMetadata(_) then "Stage 1 metadata is invalid or contains an error" else
                    (try if _[SpareIssueReason] = null or Text.Trim(_[SpareIssueReason]) = "" then "Spare permission is false or unusable" else _[SpareIssueReason] otherwise "Spare permission is false or unusable"))))
        in SchemaIssues & RowIssues,
    Issues = ReadIssues(BOriginalAvailabilityInput_RESULT, "A.1 original availability") & ReadIssues(BCappedAvailabilityInput_RESULT, "A.2 C#"),
    Output = Table.Distinct(Table.FromRecords(Issues, DetailType))
in
    Table.Buffer(Output);

// Query: BAvailabilityLineage_Prepare
// Purpose: Inspect unmasked positive A.2 cells against original A.1 availability before Resource exclusion.
shared BAvailabilityLineage_Prepare = let
    PositiveCapped = Table.SelectRows(BCappedAvailability_Prepare, each [AvailabilityCapped] <> null and [AvailabilityCapped] > 0),
    Original = Table.SelectColumns(BOriginalAvailability_Prepare, {"Role", "Resource", "Period", "Availability"}),
    Joined = Table.NestedJoin(PositiveCapped, {"Role", "Resource", "Period"}, Original, {"Role", "Resource", "Period"}, "OriginalRows", JoinKind.LeftOuter),
    Assessed = Table.AddColumn(Joined, "LineageIssue", each if Table.IsEmpty([OriginalRows]) then "No original A.1 availability row"
        else if Table.RowCount([OriginalRows]) <> 1 then "Ambiguous original A.1 availability rows"
        else if [OriginalRows]{0}[Availability] = null or [OriginalRows]{0}[Availability] <= 0 then "Original A.1 availability is not positive" else null, type nullable text),
    Failures = Table.SelectRows(Assessed, each [LineageIssue] <> null),
    Output = Table.RemoveColumns(Failures, {"OriginalRows"})
in
    Table.Buffer(Output);

// Query: BLineageIssues_DETAILS
// Purpose: Exclude spare for each Resource with unsupported C# while preserving allocation evidence.
shared BLineageIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    Rows = List.Transform(Table.ToRecords(BAvailabilityLineage_Prepare), each [Stage = "B", Check = "Availability lineage", Status = "Fail", Resource = _[Resource], Role = _[Role], Name = null, Date = null, Period = _[Period], Reason = _[LineageIssue], Action = "Exclude spare calculations for this Resource; retain allocation evidence"])
in
    Table.FromRecords(Rows, DetailType);

// Query: BPreCalculationIssues_DETAILS
// Purpose: Collect input issues before calculations, keeping Resource permission independent of output diagnostics.
shared BPreCalculationIssues_DETAILS = Table.Buffer(Table.Combine({BInputIssues_DETAILS, BContractIssues_DETAILS, BMaskIssues_DETAILS, BLineageIssues_DETAILS}));

// Query: BRedistribution_RESULT
// Purpose: Capture optional redistribution failure without concealing it or discarding allocation evidence.
shared BRedistribution_RESULT = try let
    Rows = Table.Buffer(#"ReDistribPeriodAvailbility TABLE"),
    ErrorRows = Table.SelectRowsWithErrors(Rows),
    CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else error Error.Record("CapacityDistribB.RedistributionErrors", "Optional redistribution contains cell errors.", ErrorRows)
in
    CheckedRows;

// Query: BFinalReduction_RESULT
// Purpose: Capture optional final-reduction failure; final capacity then excludes all unallocated spare.
shared BFinalReduction_RESULT = try let
    Rows = Table.Buffer(SubtractOverallocatedResources),
    ErrorRows = Table.SelectRowsWithErrors(Rows),
    CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else error Error.Record("CapacityDistribB.FinalReductionErrors", "Optional final reduction contains cell errors.", ErrorRows)
in
    CheckedRows;

// Query: BOptionalCalculationIssues_DETAILS
// Purpose: Report every optional-calculation fallback as Error with scope-wide spare exclusion.
// Notes: These post-calculation checks do not feed BResourceSpareEligibility and cannot form a dependency cycle.
shared BOptionalCalculationIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    RoleValue = try Role otherwise null,
    CaptureIssue = (checkName as text, snapshot as record) as list => if snapshot[HasError] then
        {[Stage = "B", Check = checkName, Status = "Error", Resource = null, Role = RoleValue, Name = null, Date = null, Period = null,
          Reason = try snapshot[Error][Message] otherwise "Optional calculation failed", Action = "Exclude all spare calculations for this output stage; retain allocation evidence"]} else {},
    Issues = CaptureIssue("Redistribution calculation", BRedistribution_RESULT) & CaptureIssue("Final reduction calculation", BFinalReduction_RESULT)
in
    Table.FromRecords(Issues, DetailType);

// Query: CapacityDiagnostics_DETAILS
// Purpose: Publish truthful Stage 1 issues including errors that caused optional-calculation fallback.
// Output: Resource-level evidence; a null Resource denotes a source/schema or calculation issue affecting all spare.
shared CapacityDiagnostics_DETAILS = Table.Buffer(Table.Combine({BPreCalculationIssues_DETAILS, BOptionalCalculationIssues_DETAILS}));

// Query: CapacityDiagnostics_SUMMARY
// Purpose: Summarise real Pass, Fail and Error statuses without disguising failed inputs as successful calculations.
shared CapacityDiagnostics_SUMMARY = let
    CheckNames = {"Demand input", "Allocation input", "C# input", "Original availability input", "Priority input", "Resource contract keys", "Resource contracts", "Upstream spare permission", "Availability lineage", "Redistribution calculation", "Final reduction calculation"},
    SummaryRows = List.Transform(CheckNames, (checkName) =>
        let Issues = Table.SelectRows(CapacityDiagnostics_DETAILS, each [Check] = checkName)
        in [Stage = "B", Check = checkName, Status = if List.Contains(Issues[Status], "Error") then "Error" else if Table.IsEmpty(Issues) then "Pass" else "Fail",
            Failures = Table.RowCount(Issues), Details = if Table.IsEmpty(Issues) then null else Text.Combine(List.Distinct(Issues[Reason]), "; ")]),
    Output = Table.FromRecords(SummaryRows, type table [Stage = text, Check = text, Status = text, Failures = nullable number, Details = nullable text])
in
    Table.Buffer(Output);

// Query: BResourceSpareEligibility
// Purpose: Combine upstream permission and B-local issues at Resource grain before any spare calculation.
// Notes: Diagnostics depend only on captured inputs; this view never reads C## or C### and cannot form an output cycle.
shared BResourceSpareEligibility = let
    Resources = Table.Distinct(Table.Combine({Table.SelectColumns(BOriginalAvailability_Prepare, {"Resource"}), Table.SelectColumns(BCappedAvailability_Prepare, {"Resource"}), Table.SelectColumns(BAllocation_CalculationTABLE, {"Resource"})})),
    GlobalIssues = Table.SelectRows(BPreCalculationIssues_DETAILS, each [Resource] = null and [Action] <> "Allocation skipped; evidence retained; other Resources continue"),
    ResourceIssues = Table.SelectRows(BPreCalculationIssues_DETAILS, each [Resource] <> null),
    JoinedIssues = Table.NestedJoin(Resources, {"Resource"}, ResourceIssues, {"Resource"}, "ResourceIssues", JoinKind.LeftOuter),
    WithPermission = Table.AddColumn(JoinedIssues, "SpareCalculationAllowed", each Table.IsEmpty(GlobalIssues) and Table.IsEmpty([ResourceIssues]), type logical),
    WithReason = Table.AddColumn(WithPermission, "SpareIssueReason", each if [SpareCalculationAllowed] then null else Text.Combine(List.Distinct(GlobalIssues[Reason] & [ResourceIssues][Reason]), "; "), type nullable text),
    Output = Table.RemoveColumns(WithReason, {"ResourceIssues"})
in
    Table.Buffer(Output);

// Query: fnAddIndexedRunningTotal
// Purpose: Calculate inclusive running totals in one pass through the existing indexed priority order.
// Inputs: Ordered rows with contiguous indexes starting at 2, a group-start index and a prefix count.
// Output: The original rows and columns followed by the requested total column, retaining type any.
// Notes: Buffer locally so validation, scalar projections and output rows use the same evaluated order.
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
            error Error.Record("CapacityDistribB.RunningTotalSchema", "Required running-total columns are missing.", MissingColumns)
        else if Table.HasColumns(orderedInput, {outputColumn}) then
            error Error.Record("CapacityDistribB.RunningTotalSchema", "The running-total output column already exists.", outputColumn)
        else orderedInput,
    // These buffers apply to this invocation only; they do not cache separately refreshed queries.
    StableInput = Table.Buffer(CheckedSchema),
    Values = List.Buffer(Table.Column(StableInput, valueColumn)),
    Indexes = List.Buffer(Table.Column(StableInput, indexColumn)),
    GroupStarts = List.Buffer(Table.Column(StableInput, startColumn)),
    PrefixCounts = List.Buffer(Table.Column(StableInput, countColumn)),
    RowCount = Table.RowCount(StableInput),
    // Positional ranges require unique consecutive indexes and contiguous groups; never repair bad joins silently.
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
            "CapacityDistribB.RunningTotalOrder",
            "Running-total indexes or group boundaries are invalid. Check joins and the existing priority order.",
            [RowPosition = FirstInvalidPosition + 1, IndexColumn = indexColumn, StartColumn = startColumn, CountColumn = countColumn]
        ),
    // Emit one scalar per row without repeatedly slicing/summing prefixes or appending to a growing list.
    // List.Sum of the two scalars preserves null for an entirely null prefix and ignores null otherwise.
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
        ValidatedInput, outputColumn,
        (row as record) => RunningTotals{Record.Field(row, indexColumn) - 2},
        type any
    )
in
    WithRunningTotal;


shared #"ResMaxReduction (C# Unassigned)" = let
    Source = ResPeriodOverallocationReduction,
    #"Grouped RESMAXADDITION" = Table.Group(Source, {"Resource"}, {{"MaxResAddition", each List.Sum([Unassigned]), type nullable number}}),
    #"Merged Queries1" = Table.NestedJoin(#"Grouped RESMAXADDITION", {"Resource"}, ResRosterAvailabilityCapReduction, {"Resource"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction" = Table.ExpandTableColumn(#"Merged Queries1", "ResRosterAvailabilityCapReduction", {"AvailabilityReduction"}, {"AvailabilityReduction"}),
    #"Inserted Minimum1" = Table.AddColumn(#"Expanded ResRosterAvailabilityCapReduction", "ResMaxReduction", each List.Min({[MaxResAddition], [AvailabilityReduction]}), type number),
    #"Sorted Rows" = Table.Sort(#"Inserted Minimum1",{{"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared ResPeriodMaxAddition = let
    Source = Table.NestedJoin(ResPeriodOverallocationReduction, {"Resource"}, #"ResMaxReduction (C# Unassigned)", {"Resource"}, "ResMaxAddition", JoinKind.LeftOuter),
    #"Expanded ResMaxAddition" = Table.ExpandTableColumn(Source, "ResMaxAddition", {"ResMaxReduction"}, {"ResMaxReduction"})
in
    #"Expanded ResMaxAddition";

[ Description = "%#(lf)(A-D) gaps divided to D #(lf)#(lf)Not really ABi - but should still provide same ranking of prorrity" ]
// Query: PeriodUnderallcationPrioritised
// Purpose: Retain the existing redistribution priority inputs using the checked priority preparation interface.
shared PeriodUnderallcationPrioritised = let
    Source = ResPeriodOverallocationReduction,
    #"Merged Queries2" = Table.NestedJoin(Source, {"Resource", "Period"}, ResPeriodMaxAddition, {"Resource", "Period"}, "ResPeriodMaxAddition", JoinKind.LeftOuter),
    #"Expanded ResPeriodMaxAddition" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodMaxAddition", {"ResMaxReduction"}, {"ResMaxReduction"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResPeriodMaxAddition", {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded PeriodDemandTABLE", {"Resource", "Period"}, BPriority_Prepare, {"Resource", "Period"}, "EXTRACT ResPeriodNWDTABLE", JoinKind.LeftOuter),
    #"Expanded NSWPRIORITIES" = Table.ExpandTableColumn(#"Merged Queries", "EXTRACT ResPeriodNWDTABLE", {"NWDPriority"}, {"NWDPriority"}),
    BUFFER = Table.Buffer(#"Expanded NSWPRIORITIES")
in
    BUFFER;

shared #"PrioritiseRedistribAvail-Setup" = let
    Source = #"PeriodUnderallcationPrioritised",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 2, 1, Int64.Type),
    #"Added FTE" = Table.AddColumn(#"Added Index", "FTE", each 1)
in
    #"Added FTE";

// Query: PrioritiseRedistribAvail-Distrib
// Purpose: Retain redistribution rankings and resolve Resource priority ties by Period.
shared #"PrioritiseRedistribAvail-Distrib" = let
    Source = #"PrioritiseRedistribAvail-Setup",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Index", "IndexAllRows"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Resource", Order.Ascending}, {"C#/D", Order.Descending}, {"ExcessC#-D", Order.Descending}, {"NWDPriority", Order.Descending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRowPrioritySort", 2
, 1, Int64.Type)
in
    #"Added Index";

shared #"ResIndexLimits (ResMaxReduction)" = let
    Source = #"PrioritiseRedistribAvail-Distrib",
    #"Added Index" = Table.AddIndexColumn(Source, "Index2", 1, 1, Int64.Type),
    #"Grouped Rows" = Table.Group(#"Added Index", {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRowPrioritySort]), type number}, {"ResExcessAvailable", each List.Max([ResMaxReduction]), type nullable number}, {"ResAvailNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: ReDistributeResAvailability
// Purpose: Apply each resource's reduction cap in the existing redistribution priority order.
shared ReDistributeResAvailability = let
    Source = Table.NestedJoin(#"PrioritiseRedistribAvail-Distrib", {"Resource"}, #"ResIndexLimits (ResMaxReduction)", {"Resource"}, "REsLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "REsLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded REsLimits", {"Resource", "Period"}, #"ResPeriodAvailabilityCapped(C#)TABLE", {"Resource", "Period"}, "ResPeriodAvailabilityCapped(C#)TABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityCapped(C#)TABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityCapped(C#)TABLE", {"C#"}, {"C#"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded ResPeriodAvailabilityCapped(C#)TABLE",{{"IndexAllRowPrioritySort", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows1", "Subtraction", each [IndexAllRowPrioritySort] - [StartResIndex]+1),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction", Int64.Type}}),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Changed Type", "C#", "IndexAllRowPrioritySort", "StartResIndex", "Subtraction", "RunningResTotal"),
    #"Added KEEP" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "KeepRunningTotal", each if [RunningResTotal] > [ResMaxReduction] then "Remove" else if [RunningResTotal] <= [ResMaxReduction] then "Keep" else "Split")
in
    #"Added KEEP";

[ Description = "BUFFER" ]
// Query: ReDistribResAvailabilityTABLE
// Purpose: Index redistribution rows in stable Period/Resource order for period running totals.
shared ReDistribResAvailabilityTABLE = let
    Source = ReDistributeResAvailability,
    #"Sorted Rows" = Table.Sort(Source,{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Index", 2, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"D", "NWDPriority", "IndexAllRows", "FTE", "IndexAllRowPrioritySort", "StartResIndex", "ResExcessAvailable", "ResAvailNumber", "Subtraction"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER
;

shared #"PeriodIndexLimits (ExcessC#-D)" = let
    Source = ReDistribResAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([Index]), type number}, {"MaxC#-D", each List.Max([#"ExcessC#-D"]), type nullable number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
// Query: ReDistribPeriodAvailbility TABLE
// Purpose: Retain redistribution rows within both resource and period reduction limits.
shared #"ReDistribPeriodAvailbility TABLE" = let
    Source = Table.NestedJoin(ReDistribResAvailabilityTABLE, {"Period"}, #"PeriodIndexLimits (ExcessC#-D)", {"Period"}, "PeriodIndexLimits", JoinKind.LeftOuter),
    #"Expanded PeriodIndexLimits" = Table.ExpandTableColumn(Source, "PeriodIndexLimits", {"StartPeriodIndex"}, {"StartPeriodIndex"}),
    #"Added Custom" = Table.AddColumn(#"Expanded PeriodIndexLimits", "ApplicableC#", each if [KeepRunningTotal] = "Remove" then 0 else [#"C#"]),
    #"Sorted INDEX" = Table.Sort(#"Added Custom",{{"Index", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted INDEX", "Subtraction", each [Index] - [StartPeriodIndex]+1, type number),
    #"Added PERIODRUNNINGTOTAL" = fnAddIndexedRunningTotal(#"Inserted Subtraction", "ApplicableC#", "Index", "StartPeriodIndex", "Subtraction", "PeriodRunningTotal"),
    #"Added KEEP PR" = Table.AddColumn(#"Added PERIODRUNNINGTOTAL", "Keep RunningTotal", each if [RunningResTotal] <= [ResMaxReduction]
and 
[PeriodRunningTotal] <=[#"ExcessC#-D"]
then "Keep" else "Remove"),
    #"Filtered Rows" = Table.SelectRows(#"Added KEEP PR", each ([Keep RunningTotal] = "Keep")),
    BUFFER = Table.Buffer(#"Filtered Rows")
in
    BUFFER;

shared CheckReAllocation = let
    Source = #"ReDistribPeriodAvailbility TABLE",
    #"Grouped Rows1" = Table.Group(Source, {"Resource"}, {{"Max", each List.Max([RunningResTotal]), type number}, {"ReDirect", each List.Sum([#"C#"]), type nullable number}})
in
    #"Grouped Rows1";

shared ReAllocateMATRIX = let
    Source = ReDistribResAvailabilityTABLE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "C#"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Period", type text}}),
    #"Pivoted Column1" = Table.Pivot(#"Changed Type1", List.Distinct(#"Changed Type1"[Period]), "Period", "C#", List.Sum)
in
    #"Pivoted Column1";

// Query: ResRosterAvailabilityC##CapReduction
// Purpose: Calculate the existing C## cap-reduction target only for Resources allowed to calculate spare.
shared #"ResRosterAvailabilityC##CapReduction" = let
    Source = Table.SelectRows(#"C##TABLE", each [SpareCalculationAllowed] = true),
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([#"C##"]), type nullable number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped Rows",null,0,Replacer.ReplaceValue,{"RosterAvailability"}),
    // Join the cap only after aggregation so contract totals are never repeated at Resource-Period grain.
    #"Merged Resource Contract" = Table.NestedJoin(#"Replaced Value", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added AVAILCAP" = Table.AddColumn(#"Expanded Effective Shift Cap", "RosterAvailabilityCAPPED", each
        // Allocation-only Resources can remain in the grid with zero availability and no contract cap.
        if [RosterAvailability] = 0 then 0
        else if [Effective Shift Cap] = null then null
        else if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [RosterAvailability], type number),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED", "Effective Shift Cap"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityReduction", "MaxC##Reduction"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"MaxC##Reduction"] <> 0))
in
    #"Filtered Rows";

shared Role = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Role")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: IMPORTSource Settings Data
// Purpose: Import the shared Settings Data workbook for this unit.
shared #"IMPORTSource Settings Data" = let
    // Settings Data is in the parent Calculations folder for every role.
    CalculationsPath = Text.BeforeDelimiter(
        RolePath, "\", {0, RelativePosition.FromEnd}
    ),
    Source = Excel.Workbook(
        File.Contents(CalculationsPath & "\Settings Data.xlsx"),
        null,
        true
    )
in
    Source;

// Query: EXTRACT MaxAvailability
// Purpose: Extract the analysis-period worker shift cap from Settings Data.
shared #"EXTRACT MaxAvailability" = let
    SettingsTable = #"IMPORTSource Settings Data"{
        [Item = "MaxAvailability", Kind = "Table"]
    }[Data],
    TypedSettings = Table.TransformColumnTypes(
        SettingsTable, {{"MaxAvailability", Int64.Type}}
    ),
    PeriodShiftLimit = TypedSettings{0}[MaxAvailability]
in
    PeriodShiftLimit;

// Query: MaxAvailability
// Purpose: Preserve the existing cap interface for downstream B calculations.
shared MaxAvailability = #"EXTRACT MaxAvailability";

shared AllocationThreshold = 0.8 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

// Query: IMPORTSource StaffListMaster
// Purpose: Import the Unit1 StaffListMaster workbook once for B's Resource-grain contract caps.
shared #"IMPORTSource StaffListMaster" = let
    CalculationsPath = Text.BeforeDelimiter(RolePath, "\", {0, RelativePosition.FromEnd}),
    WorkbookBinary = Binary.Buffer(File.Contents(CalculationsPath & "\StaffListMaster.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORT ResourceContract
// Purpose: Extract preferred-role contracts plus approved missing-worker fallbacks from Table_Masterlist for this role.
// Output: One candidate cap row per applicable AIN Resource before B's independent checks.
shared #"IMPORT ResourceContract" = let
    Matches = Table.SelectRows(#"IMPORTSource StaffListMaster", each [Item] = "Table_Masterlist" and [Kind] = "Table"),
    MasterlistTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("B resource contract import", "Expected exactly one Table_Masterlist table in StaffListMaster.xlsx.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Resource", "Role", "PreferredRole", "Source", "Worker Record Status", "Effective Shift Cap", "Limit Basis"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MasterlistTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(MasterlistTable, RequiredColumns)
        else error Error.Record("B resource contract import", "Table_Masterlist is missing required contract columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected, {
        {"Resource", Int64.Type}, {"Role", type text}, {"PreferredRole", type text},
        {"Source", type text}, {"Worker Record Status", type text},
        {"Effective Shift Cap", type number}, {"Limit Basis", type text}
    }),
    NormalizedRoles = Table.TransformColumns(Typed, {
        {"Role", each if _ = null then null else Text.Clean(Text.Trim(_)), type nullable text},
        {"PreferredRole", each if _ = null then null else Text.Clean(Text.Trim(_)), type nullable text}
    }),
    RoleValue = Role,
    // Preserve preferred-role ownership, while allowing the approved Settings fallback for allocation-only employees absent from Worker's Report.
    ApplicableRoleCaps = Table.SelectRows(NormalizedRoles, each
        [Role] = RoleValue
            and ([PreferredRole] = RoleValue
                or ([Source] = "Allocation" and [Worker Record Status] = "Not found")))
in
    Table.Buffer(ApplicableRoleCaps);

// Query: BResourceContract_CHECK
// Purpose: Retain the legacy contract-check interface with honest Stage 1 failure statuses.
// Output: Diagnostic rows only; failed Resources lose spare calculations without stopping unrelated Resources.
shared BResourceContract_CHECK = let
    KeyIssues = Table.SelectRows(BInputIssues_DETAILS, each [Check] = "Resource contract keys"),
    MakeCheck = (checkName as text, issues as table) as record => [Check = checkName,
        Status = if List.Contains(issues[Status], "Error") then "Error" else if Table.IsEmpty(issues) then "Pass" else "Fail",
        Failures = Table.RowCount(issues), Details = if Table.IsEmpty(issues) then null else Text.Combine(List.Distinct(issues[Reason]), "; ")],
    SchemaOrCellIssues = Table.SelectRows(KeyIssues, each [Status] = "Error"),
    Checks = Table.FromRecords({
        MakeCheck("Populated Resource contract keys", Table.SelectRows(KeyIssues, each not Text.StartsWith([Reason], "Duplicate join key"))),
        MakeCheck("Unique Resource contracts", Table.Combine({SchemaOrCellIssues, Table.SelectRows(KeyIssues, each Text.StartsWith([Reason], "Duplicate join key"))})),
        MakeCheck("Valid effective Resource caps", Table.Combine({SchemaOrCellIssues, Table.SelectRows(BContractIssues_DETAILS, each Text.StartsWith([Reason], "Effective Shift Cap"))})),
        MakeCheck("Complete AIN availability contract coverage", Table.Combine({SchemaOrCellIssues, Table.SelectRows(BContractIssues_DETAILS, each [Reason] = "Positive availability has no Resource contract")})),
        MakeCheck("Missing-worker fallback has zero AIN availability", Table.Combine({SchemaOrCellIssues, Table.SelectRows(BContractIssues_DETAILS, each [Reason] = "Missing-worker fallback has positive AIN availability")}))
    }, type table [Check = text, Status = text, Failures = number, Details = nullable text])
in
    Table.Buffer(Checks);

// Query: BResourceContract
// Purpose: Publish only unique usable contract caps; affected Resources are excluded through Stage 1 diagnostics.
shared BResourceContract = let
    PreparedContracts = fnBPreparedInput(BContractInput_RESULT, {"Resource", "Effective Shift Cap", "Limit Basis", "Worker Record Status"},
        #table(type table [Resource = nullable number, #"Effective Shift Cap" = nullable number, #"Limit Basis" = nullable text, #"Worker Record Status" = nullable text], {})),
    ValidCaps = Table.SelectRows(PreparedContracts, each [Effective Shift Cap] <> null and not Number.IsNaN([Effective Shift Cap]) and [Effective Shift Cap] <> #infinity and [Effective Shift Cap] <> -#infinity and [Effective Shift Cap] > 0
        and [Effective Shift Cap] = Number.RoundDown([Effective Shift Cap]) and [Limit Basis] <> null and Text.Trim([Limit Basis]) <> ""),
    Counts = Table.Group(PreparedContracts, {"Resource"}, {{"ContractRows", each Table.RowCount(_), Int64.Type}}),
    UniqueResources = Table.SelectRows(Counts, each [ContractRows] = 1),
    UniqueCaps = Table.NestedJoin(ValidCaps, {"Resource"}, UniqueResources, {"Resource"}, "UniqueResource", JoinKind.Inner),
    ContractIssueResources = Table.Distinct(Table.SelectColumns(Table.Combine({
        Table.SelectRows(BInputIssues_DETAILS, each [Check] = "Resource contract keys" and [Resource] <> null), BContractIssues_DETAILS
    }), {"Resource"})),
    UnambiguousCaps = Table.NestedJoin(UniqueCaps, {"Resource"}, ContractIssueResources, {"Resource"}, "ContractIssues", JoinKind.LeftAnti),
    Output = Table.SelectColumns(UnambiguousCaps, {"Resource", "Effective Shift Cap", "Limit Basis"})
in
    Table.Buffer(Output);

// Query: IMPORTSource A1
// Purpose: Provide the A.1 workbook binary and navigation table for the four existing imports.
// Notes: Navigation buffering is shallow; selected table transformations remain in their owning imports.
shared #"IMPORTSource A1" = let
    // Reuse the saved file bytes within an evaluation; separately loaded outputs can still evaluate this again.
    WorkbookBinary = Binary.Buffer(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx")),
    WorkbookNavigation = Table.Buffer(Excel.Workbook(WorkbookBinary, null, true))
in
    WorkbookNavigation;

// Query: IMPORT Demand
// Purpose: Read the existing A.1 Demand_Prepare table without changing its schema or grain.
shared #"IMPORT Demand" = let
    Source = #"IMPORTSource A1",
    Demand_Prepare_Table = Source{[Item="Demand_Prepare",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Demand_Prepare_Table,{{"Shift", type text}, {"Role", type text}, {"Facility", type text}, {"D", Int64.Type}, {"Date", type date}, {"Period", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORT Allocation
// Purpose: Read A.1 resource-period allocations, preserving the existing null-period exclusion.
shared #"IMPORT Allocation" = let
    Source = #"IMPORTSource A1",
    ResPeriodAllocationTABLE_Table = Source{[Item="ResPeriodAllocationTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAllocationTABLE_Table,{{"Role", type text}, {"Allocation", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Period] <> null))
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
// Query: IMPORT AvailabilityOriginal
// Purpose: Read and retain the existing buffered original-availability table from A.1.
shared #"IMPORT AvailabilityOriginal" = let
    Source = #"IMPORTSource A1",
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;

shared #"IMPORT ResPeriodAvailabilityCapped(C#)" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodCappedAvailability_C__TABLE_Table = Source{[Item="ResPeriodCappedAvailability_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodCappedAvailability_C__TABLE_Table,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"AvailabilityCapped", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

// Query: EXTRACT ResPeriodNWDTABLE
// Purpose: Return A.1's true NWD priority table unchanged from the existing workbook import.
// Notes: The user confirmed all A.1 loaded outputs are tables; ResPeriodWDTABLE remains A.2's separate workday-status interface.
shared #"EXTRACT ResPeriodNWDTABLE" = let
    Source = #"IMPORTSource A1",
    PriorityTable = Source{[Item="ResPeriodShiftNWDTABLE",Kind="Table"]}[Data]
in
    PriorityTable;

// Query: PeriodAllocationTABLE
// Purpose: Preserve B's individual-fact AllocationThreshold arithmetic and period totals on usable allocation evidence.
shared PeriodAllocationTABLE = let
    Source = BAllocationFacts_Prepare,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Allocation", "AllocationX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Allocation", each if [AllocationX] > AllocationThreshold then 1 else [AllocationX]),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"Period"}, {{"A", each List.Sum([Allocation]), type number}})
in
    #"Grouped Rows";

shared #"PeriodA-D" = let
    Source = Table.NestedJoin(PeriodDemandTABLE, {"Period"}, PeriodAllocationTABLE, {"Period"}, "AllocationPeriodLIST", JoinKind.LeftOuter),
    #"Expanded AllocationPeriodLIST" = Table.ExpandTableColumn(Source, "AllocationPeriodLIST", {"A"}, {"A"}),
    #"Sorted Rows" = Table.Sort(#"Expanded AllocationPeriodLIST",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

// Query: ResRosterAvailabilityCapReduction
// Purpose: Calculate the existing C# cap-reduction target only for Resources allowed to calculate spare.
shared ResRosterAvailabilityCapReduction = let
    Source = Table.SelectRows(ResPeriodAvailabilityTABLE, each [SpareCalculationAllowed] = true),
    #"Grouped ROSTERAVAILABILITY" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([#"C#"]), type nullable number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped ROSTERAVAILABILITY",null,0,Replacer.ReplaceValue,{"RosterAvailability"}),
    // Join the cap only after aggregation so contract totals are never repeated at Resource-Period grain.
    #"Merged Resource Contract" = Table.NestedJoin(#"Replaced Value", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added AVAILCAP" = Table.AddColumn(#"Expanded Effective Shift Cap", "RosterAvailabilityCAPPED", each
        // Allocation-only Resources can remain in the grid with zero availability and no contract cap.
        if [RosterAvailability] = 0 then 0
        else if [Effective Shift Cap] = null then null
        else if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [RosterAvailability], type number),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED", "Effective Shift Cap"})
in
    #"Removed Columns";

[ Description = "BUFFER" ]
// Query: ResPeriodOverallocationReduction
// Purpose: Retain the existing redistribution candidates after excluding every Resource with a spare-calculation issue.
shared ResPeriodOverallocationReduction = let
    // Period A-D Overallocated
    Source = #"PeriodA-DPos (Overallocation)",
    BUFFER = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Period"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodCapacityTABLE", {"PeriodCapacity"}, {"PeriodCapacity"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodCapacityTABLE", {"Period"}, #"PeriodC#-DPos (Excess)TABLE", {"Period"}, "PeriodC#-DPos (Excess)TABLE", JoinKind.LeftOuter),
    #"Expanded PeriodC#-DPos (Excess)TABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodC#-DPos (Excess)TABLE", {"ExcessC#-D", "C#/D"}, {"ExcessC#-D", "C#/D"}),
    #"Merged Queries3" = Table.NestedJoin(#"Expanded PeriodC#-DPos (Excess)TABLE", {"Period"}, #"ResourcePeriods-EmptyTABLE", {"Period"}, "ResourcePeriods-EmptyTABLE", JoinKind.LeftOuter),
    #"Expanded ResourcePeriods-EmptyTABLE" = Table.ExpandTableColumn(#"Merged Queries3", "ResourcePeriods-EmptyTABLE", {"Resource"}, {"Resource"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded ResourcePeriods-EmptyTABLE",{"Resource", "Period", "D", "A", "A-D", "PeriodCapacity", "ExcessC#-D", "C#/D"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns1",{{"Resource", Int64.Type}}),
    // Resource-local failures remove all optional candidates before allocation and priority joins.
    AllowedResources = List.Buffer(Table.SelectRows(BResourceSpareEligibility, each [SpareCalculationAllowed])[Resource]),
    EligibleResources = Table.SelectRows(#"Changed Type", each List.Contains(AllowedResources, [Resource])),
    #"Merged Queries2" = Table.NestedJoin(EligibleResources, {"Period", "Resource"}, BAllocation_CalculationTABLE, {"Period", "Resource"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries5" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource"}, ResRosterAvailabilityCapReduction, {"Resource"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction1" = Table.ExpandTableColumn(#"Merged Queries5", "ResRosterAvailabilityCapReduction", {"AvailabilityReduction"}, {"ResAvailabilityReduction"}),
    #"Merged Queries4" = Table.NestedJoin(#"Expanded ResRosterAvailabilityCapReduction1", {"Resource", "Period"}, ResPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction" = Table.ExpandTableColumn(#"Merged Queries4", "ResRosterAvailabilityCapReduction", {"C#"}, {"C#"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResRosterAvailabilityCapReduction",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    /*#"Filtered Rows1" = Table.SelectRows(#"Expanded ResRosterAvailabilityCapReduction", each ([#"C#"] = 1) and ([ResAvailabilityReduction] <> null)),
   */ 
   #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each 
        ([ResAvailabilityReduction] <> null 
        and [ResAvailabilityReduction] <> 0) 
        and ([#"C#/D"] <> null 
        and [#"C#/D"] <> 0) 
        and ([#"C#"] <> 0) 
        and ([#"C#"] <> null
        and (([Allocation] = null )
             or ([Allocation] <> 0 ))  )
    ),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"C#", "Unassigned"}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Renamed Columns1",{"Resource", "Period", "ExcessC#-D", "C#/D", "Unassigned"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"Resource", "Period", "Unassigned"})
in
    #"Reordered Columns";

shared ResPeriodAllocationTABLE = let
    Source = #"IMPORT Allocation",
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Resource", Int64.Type}})
in
    #"Changed Type1";

shared AvailabilityOriginalMATRIX = let
    Source = #"ResourcePeriods-EmptyTABLE",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Shift", "Role", "Facility", "Date"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Resource", "Period"}, #"IMPORT AvailabilityOriginal", {"Resource", "Period"}, "IMPORT AvailabilityOriginal", JoinKind.LeftOuter),
    #"Expanded IMPORT AvailabilityOriginal" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT AvailabilityOriginal", {"Availability"}, {"Availability"}),
    #"Sorted Rows" = Table.Sort(#"Expanded IMPORT AvailabilityOriginal",{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Availability", List.Sum)
in
    #"Pivoted Column";

// Query: PeriodDemandTABLE
// Purpose: Preserve the period-demand interface using schema-safe preparation and separate error diagnostics.
shared PeriodDemandTABLE = let
    Source = BDemand_Prepare,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}})
in
    #"Changed Type";

shared PeriodDemandMATRIX = let
    Source = PeriodDemandTABLE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Period", "D"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "D", List.Sum)
in
    #"Pivoted Column";

[ Description = "BUFFER-All cells capped availability" ]
// Query: ResPeriodAvailabilityCapped(C#)TABLE
// Purpose: Apply upstream and B-local Resource permission to spare while retaining cells with actual-allocation evidence.
// Output: One Resource/Period C# row with Stage 1 permission, reason and allocation evidence.
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    Source = BCappedAvailability_Prepare,
    RoleRows = Table.SelectRows(Source, each [Role] = Role and [AvailabilityCapped] <> null),
    // A duplicate input key is already a Resource issue. Preserve one existing allocated-cell amount, never multiply it through joins.
    GroupedCells = Table.Group(RoleRows, {"Resource", "Period"}, {
        {"Role", each List.First([Role]), type nullable text},
        {"C#", each List.Max([AvailabilityCapped]), type nullable number},
        {"UpstreamAllocationEvidence", each List.AnyTrue(List.Transform(Table.ToRecords(_), each (try _[HasAllocationEvidence] otherwise false) = true)), type logical},
        {"OriginalAvailability", each List.Max(List.Transform(Table.ToRecords(_), each try _[OriginalAvailability] otherwise null)), type nullable number}
    }),
    JoinedAllocationEvidence = Table.NestedJoin(GroupedCells, {"Resource", "Period"}, BAllocation_CalculationTABLE, {"Resource", "Period"}, "AllocationEvidence", JoinKind.LeftOuter),
    WithAllocationEvidence = Table.AddColumn(JoinedAllocationEvidence, "HasAllocationEvidence", each [UpstreamAllocationEvidence]
        or (not Table.IsEmpty([AllocationEvidence]) and [AllocationEvidence]{0}[HasAllocationEvidence]), type logical),
    RemovedEvidenceJoin = Table.RemoveColumns(WithAllocationEvidence, {"UpstreamAllocationEvidence", "AllocationEvidence"}),
    JoinedPermission = Table.NestedJoin(RemovedEvidenceJoin, {"Resource"}, BResourceSpareEligibility, {"Resource"}, "SpareEligibility", JoinKind.LeftOuter),
    ExpandedPermission = Table.ExpandTableColumn(JoinedPermission, "SpareEligibility", {"SpareCalculationAllowed", "SpareIssueReason"}, {"SpareCalculationAllowed", "SpareIssueReason"}),
    WithMaskedCapacity = Table.AddColumn(ExpandedPermission, "MaskedC#", each if [SpareCalculationAllowed] = true or [HasAllocationEvidence] then [#"C#"] else null, type nullable number),
    RemovedUnmaskedCapacity = Table.RemoveColumns(WithMaskedCapacity, {"C#"}),
    RenamedCapacity = Table.RenameColumns(RemovedUnmaskedCapacity, {{"MaskedC#", "C#"}}),
    ZeroToNull = Table.ReplaceValue(RenamedCapacity, 0, null, Replacer.ReplaceValue, {"C#"})
in
    Table.Buffer(ZeroToNull);

// Query: ResPeriodAvailabilityCapped(C#)MATRIX
// Purpose: Preserve the legacy C# matrix grain without using Stage 1 metadata as pivot keys.
shared #"ResPeriodAvailabilityCapped(C#)MATRIX" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Removed Columns" = Table.SelectColumns(Source,{"Resource", "Period", "C#"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "C#", List.Sum)
in
    #"Pivoted Column";

shared ResPeriodAvailabilityTABLE = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE"
in
    Source;

// Query: CapacityDistribB_INPUT_CHECK
// Purpose: Retain the four legacy input-check rows while diagnosing raw inputs before Resource exclusion.
// Output: Honest Passed flags; failures are diagnostic and are no longer a global publication gate.
shared CapacityDistribB_INPUT_CHECK = let
    CheckKeys = (inputName as text, snapshot as record, keys as list, detailCheck as text, skipUnmatched as logical) as record =>
        let
            inputTable = if snapshot[HasError] then #table({}, {}) else snapshot[Value],
            MissingColumns = List.Difference(keys, Table.ColumnNames(inputTable)),
            HasRequiredColumns = not snapshot[HasError] and List.IsEmpty(MissingColumns),
            MatchedRows = if skipUnmatched and HasRequiredColumns then Table.SelectRows(inputTable, each (try [Resource] otherwise null) <> null) else inputTable,
            // Only key columns are buffered and scanned; the calculation's input table is not altered.
            KeyRows = if HasRequiredColumns then Table.Buffer(Table.SelectColumns(MatchedRows, keys)) else #table(keys, {}),
            ErrorKeyRows = if HasRequiredColumns then Table.RowCount(Table.SelectRowsWithErrors(KeyRows, keys)) else null,
            KeysWithoutErrors = Table.RemoveRowsWithErrors(KeyRows, keys),
            HasMissingKey = (row as record) as logical => List.AnyTrue(List.Transform(
                Record.FieldValues(row),
                (keyValue) => keyValue = null or (Value.Is(keyValue, type text) and Text.Trim(keyValue) = "")
            )),
            MissingKeyRows = if HasRequiredColumns then Table.RowCount(Table.SelectRows(KeysWithoutErrors, HasMissingKey)) else null,
            CompleteKeys = Table.SelectRows(KeysWithoutErrors, each not HasMissingKey(_)),
            KeyCounts = Table.Group(CompleteKeys, keys, {{"KeyCount", each Table.RowCount(_), Int64.Type}}),
            DuplicateKeyGroups = if HasRequiredColumns then Table.RowCount(Table.SelectRows(KeyCounts, each [KeyCount] > 1)) else null
        in [
            Input = inputName,
            Keys = Text.Combine(keys, ", "),
            MissingColumns = Text.Combine(MissingColumns, ", "),
            ErrorKeyRows = ErrorKeyRows,
            MissingKeyRows = MissingKeyRows,
            DuplicateKeyGroups = DuplicateKeyGroups,
            Passed = HasRequiredColumns and ErrorKeyRows = 0 and MissingKeyRows = 0 and DuplicateKeyGroups = 0
                and Table.IsEmpty(Table.SelectRows(BInputIssues_DETAILS, each [Check] = detailCheck))
        ],
    Checks = {
        CheckKeys("PeriodDemandTABLE", BDemandInput_RESULT, {"Period"}, "Demand input", false),
        CheckKeys("ResPeriodAllocationTABLE", BAllocationInput_RESULT, {"Resource", "Period"}, "Allocation input", true),
        CheckKeys("ResPeriodAvailabilityCapped(C#)TABLE", BCappedAvailabilityInput_RESULT, {"Resource", "Period"}, "C# input", false),
        CheckKeys("EXTRACT ResPeriodNWDTABLE", BPriorityInput_RESULT, {"Resource", "Period"}, "Priority input", false)
    },
    TypedChecks = Table.TransformColumnTypes(Table.FromRecords(Checks), {
        {"Input", type text}, {"Keys", type text}, {"MissingColumns", type text},
        {"ErrorKeyRows", Int64.Type}, {"MissingKeyRows", Int64.Type}, {"DuplicateKeyGroups", Int64.Type}, {"Passed", type logical}
    })
in
    // Reuse these four scalar diagnostic rows within the output's validation and failure reporting.
    Table.Buffer(TypedChecks);

// Query: CapacityDistribB_AVAILABILITY_LINEAGE_CHECK
// Purpose: Identify positive A.2 Resource/Period availability without positive original A.1 availability.
// Output: One diagnostic row per unsupported positive C#; affected Resources lose spare calculations.
// Notes: B must not carry stale A.2 capacity into a Resource/Period absent from the current A.1 source.
shared CapacityDistribB_AVAILABILITY_LINEAGE_CHECK = let
    Output = Table.RenameColumns(BAvailabilityLineage_Prepare, {{"AvailabilityCapped", "C#"}})
in Table.Buffer(Output);

shared #"ResPeriodAvailabilityCapped(C#)SUM" = let
    Source = #"ResPeriodAvailabilityTABLE",
    #"C#" = Source[#"C#"],
    #"Calculated Sum" = List.Sum(#"C#")
in
    #"Calculated Sum";

shared #"PeriodOverallocatedExcess%" = let
    Source = #"ResPeriodMaxC##Reduction",
    #"PERIOD C= Grouped PERIOD.C###" = Table.Group(Source, {"Period"}, {{"PeriodAvailability", each List.Sum([#"C##"]), type nullable number}}),
    #"Merged Queries" = Table.NestedJoin(#"PERIOD C= Grouped PERIOD.C###", {"Period"}, #"PeriodA-DPos (Overallocation)", {"Period"}, "PeriodA-DPos (Excess)", JoinKind.LeftOuter),
    #"Expanded PeriodA-DPos (Excess)" = Table.ExpandTableColumn(#"Merged Queries", "PeriodA-DPos (Excess)", {"D", "A-D"}, {"D", "A-D"}),
    #"Added EXCESS%" = Table.AddColumn(#"Expanded PeriodA-DPos (Excess)", "Excess%", each [#"A-D"] / [D]),
    #"Filtered NULL" = Table.SelectRows(#"Added EXCESS%", each [#"Excess%"] <> null),
    #"Sorted Rows1" = Table.Sort(#"Filtered NULL",{{"Excess%", Order.Descending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows1";

shared #"PeriodA-DPos (Overallocation)" = let
    Source = #"PeriodA-D",
    #"Inserted Subtraction" = Table.AddColumn(Source, "A-D", each [A]-[D], type number),
    #"Filtered Rows" = Table.SelectRows(#"Inserted Subtraction", each [#"A-D"] > 0)
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
// Query: PrioritiseReductionAvail-Setup
// Purpose: Attach resource-period priorities and resolve period reduction ties by Resource.
shared #"PrioritiseReductionAvail-Setup" = let
    Source = Table.NestedJoin(#"ResPeriodMaxC##Reduction", {"Period"}, #"PeriodOverallocatedExcess%", {"Period"}, "PeriodOverallocatedPriroristised", JoinKind.LeftOuter),
    #"Expanded PeriodOverallocatedPriroristised" = Table.ExpandTableColumn(Source, "PeriodOverallocatedPriroristised", {"Excess%"}, {"Excess%"}),
    // Pair Resource with Resource and Period with Period; positional join keys must use the same order.
    #"Merged Queries" = Table.NestedJoin(#"Expanded PeriodOverallocatedPriroristised", {"Resource", "Period"}, BPriority_Prepare, {"Resource", "Period"}, "EXTRACT ResPeriodNWDTABLE", JoinKind.LeftOuter),
    #"Expanded EXTRACT ResPeriodNWDTABLE" = Table.ExpandTableColumn(#"Merged Queries", "EXTRACT ResPeriodNWDTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Inserted MINAVAILABLEREDUCTION" = Table.AddColumn(#"Expanded EXTRACT ResPeriodNWDTABLE", "MinimumAvailable", each List.Min({[#"C##"], [#"A-D"]})),
    #"Sorted Rows" = Table.Sort(#"Inserted MINAVAILABLEREDUCTION",{{"Period", Order.Ascending}, {"Excess%", Order.Descending}, {"ResPeriodC##MAXReduction", Order.Descending}, {"Resource", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

shared #"PeriodIndexLimits-ve (C##-D,C##)" = let
    Source = #"PrioritiseReductionAvail-Setup",
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([IndexAllRows]), type number}, {"PeriodExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"PeriodAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: SubtractOverallatedPeriods
// Purpose: Apply period reduction limits in the existing indexed priority order.
shared SubtractOverallatedPeriods = let
    Source = Table.NestedJoin(#"PrioritiseReductionAvail-Setup", {"Period"}, #"PeriodIndexLimits-ve (C##-D,C##)", {"Period"}, "ResLimits-Subtract", JoinKind.LeftOuter),
    #"Expanded ResLimits-Subtract" = Table.ExpandTableColumn(Source, "ResLimits-Subtract", {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}, {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResLimits-Subtract",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartPeriodIndex]+1),
    #"Reordered Columns" = Table.ReorderColumns(#"Inserted Subtraction",{"Period", "Resource", "C##", "ExcessC##-D", "C##/D", "A-D", "ResPeriodC##MAXReduction", "Excess%", "NWDPriority", "IndexAllRows", "StartPeriodIndex", "PeriodExcessAvailable", "MinimumAvailable", "PeriodAvailableNumber", "Subtraction"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Subtraction", Int64.Type}}),




    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Changed Type", "MinimumAvailable", "IndexAllRows", "StartPeriodIndex", "Subtraction", "RunningPeriodTotal"),
    #"Added KEEP" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "KeepRunningTotal", each if [RunningPeriodTotal] > [#"ResPeriodC##MAXReduction"] then "Remove" else if [RunningPeriodTotal] <= [#"ResPeriodC##MAXReduction"] then "Keep" else "Split")
in
    #"Added KEEP";

// Query: ResIndex-Steup
// Purpose: Retain Resource reduction rankings and resolve ties by Period before indexing.
shared #"ResIndex-Steup" = let
    Source = SubtractOverallatedPeriods,
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "ResRosterAvailabilityC##CapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityC##CapReduction" = Table.ExpandTableColumn(#"Merged Queries", "ResRosterAvailabilityC##CapReduction", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded ResRosterAvailabilityC##CapReduction",{"Period", "Resource", "A-D", "ResPeriodC##MAXReduction", "MinimumAvailable", "RunningPeriodTotal", "KeepRunningTotal", "MaxC##Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"MaxC##Reduction", Order.Descending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type)
in
    #"Added Index";

shared ResIndexLimits = let
    Source = #"ResIndex-Steup",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRows]), type number}, {"ResExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"ResAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: SubtractOverallocatedResources
// Purpose: Retain reduction rows within both resource and period caps using the existing indexed order.
shared SubtractOverallocatedResources = let
    Source = Table.NestedJoin(#"ResIndex-Steup", {"Resource"}, ResIndexLimits, {"Resource"}, "ResIndexLimits", JoinKind.LeftOuter),
    #"Expanded ResIndexLimits" = Table.ExpandTableColumn(Source, "ResIndexLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResIndexLimits",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartResIndex]+1),
    #"Added Custom" = fnAddIndexedRunningTotal(#"Inserted Subtraction", "MinimumAvailable", "IndexAllRows", "StartResIndex", "Subtraction", "RunningResTotal"),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",null,0,Replacer.ReplaceValue,{"A-D"}),
    #"Added KEEP RES" = Table.AddColumn(#"Replaced Value", "KeepRes", each if [RunningResTotal] <= [#"MaxC##Reduction"] then "Keep" else "Remove"),
    #"Added Conditional Column" = Table.AddColumn(#"Added KEEP RES", "KeepPR", each if [KeepRunningTotal] <> "Keep" then null else if [KeepRes] <> "Keep" then null else "Keep"),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([KeepPR] = "Keep"))
in
    #"Filtered Rows";

shared ReduceCheck = let
    Source = SubtractOverallocatedResources,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"MaxReDistrib", each List.Max([RunningResTotal]), type number}, {"Available", each List.Sum([MinimumAvailable]), type number}})
in
    #"Grouped Rows";

// Query: C##TABLE
// Purpose: Preserve redistribution arithmetic and exclude spare when its optional calculation fails.
shared #"C##TABLE" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    RedistributionRows = if BRedistribution_RESULT[HasError] then #table(type table [Resource = nullable number, Period = nullable number, #"C#" = nullable number], {}) else BRedistribution_RESULT[Value],
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, RedistributionRows, {"Resource", "Period"}, "ReDistribAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReDistribAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReDistribAvailabilityTABLE", {"C#"}, {"C#.1"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ReDistribAvailabilityTABLE",null,0,Replacer.ReplaceValue,{"C#.1"}),
    #"CHANGE C" = Table.AddColumn(#"Replaced Value", "C##", each if BRedistribution_RESULT[HasError] and not [HasAllocationEvidence] then null else [#"C#"] - [#"C#.1"], type number),
    #"Removed Columns" = Table.RemoveColumns(#"CHANGE C",{"C#", "C#.1"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Removed Columns",0,null,Replacer.ReplaceValue,{"C##"}),
    WithStagePermission = Table.TransformColumns(#"Replaced Value1", {
        {"SpareCalculationAllowed", each _ = true and not BRedistribution_RESULT[HasError], type logical},
        {"SpareIssueReason", each if not BRedistribution_RESULT[HasError] then _ else Text.Combine(List.RemoveNulls({_, "Redistribution calculation failed; all spare excluded at C##"}), "; "), type nullable text}
    })
in
    WithStagePermission;

shared #"C##SUM" = let
    Source = #"C##TABLE",
    #"C##" = Source[#"C##"],
    #"Calculated Sum" = List.Sum(#"C##")
in
    #"Calculated Sum";

// Query: C##MATRIX
// Purpose: Preserve the legacy C## matrix grain without using Stage 1 metadata as pivot keys.
shared #"C##MATRIX" = let
    Source = #"C##TABLE",
    LegacyMatrixColumns = Table.SelectColumns(Source, {"Resource", "Period", "Role", "C##"}),
    #"Sorted Rows" = Table.Sort(LegacyMatrixColumns,{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column1" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "C##", List.Sum)
in
    #"Pivoted Column1";

shared #"PeriodC#-D" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Grouped PERIODC#" = Table.Group(Source, {"Period"}, {{"PeriodC#", each List.Sum([#"C#"]), type nullable number}}),
    #"Merged Queries" = Table.NestedJoin(#"Grouped PERIODC#", {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C#-D", each [#"PeriodC#"] - [D], type number),
    #"Inserted Division" = Table.AddColumn(#"Inserted Subtraction", "Division", each [#"PeriodC#"] / [D], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Division",{{"Division", "C#/D"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"D"})
in
    #"Removed Columns";

shared #"PeriodC#-DPos (Excess)TABLE" = let
    Source = #"PeriodC#-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C#-D"] > 0),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"C#-D", "ExcessC#-D"}})
in
    #"Renamed Columns";

shared PeriodCapacityTABLE = let
    Source = ResPeriodAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"PeriodCapacity", each List.Sum([#"C#"]), type nullable number}})
in
    #"Grouped Rows";

// Query: ResMaxAvailability
// Purpose: Retain original availability and usable cap audit measures without inventing a missing Resource cap.
// Query: ResMaxAvailability
// Purpose: Report the A.1 published availability total and its known contract limit at Resource grain.
// Notes: A.1 published Availability excludes issue-Resource spare; raw per-cell OriginalAvailability remains on the capacity tables.
shared ResMaxAvailability = let
    Source = BOriginalAvailability_Prepare,
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"ResAvailability", each List.Sum([Availability]), type nullable number}}),
    // ResMaxAvail is a Resource measure, so attach the contract cap after Resource aggregation.
    #"Merged Resource Contract" = Table.NestedJoin(#"Grouped Rows", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added RESMAXAVAILABILITY" = Table.AddColumn(#"Expanded Effective Shift Cap", "ResMaxAvail", each
        // A Resource with no original availability does not need a contract cap in this calculation.
        if [ResAvailability] = null then 0
        else if [ResAvailability] = 0 then 0
        // An unusable cap remains unknown in the audit output; it is never invented as zero or a Settings fallback.
        else if [Effective Shift Cap] = null then null
        else if [ResAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [ResAvailability], type number),
    #"Removed Effective Shift Cap" = Table.RemoveColumns(#"Added RESMAXAVAILABILITY", {"Effective Shift Cap"})
in
    #"Removed Effective Shift Cap";

shared ResourcesLIST = let
    Source = ResPeriodAvailabilityTABLE,
    Resource = Source[Resource],
    #"Removed Duplicates" = List.Distinct(Resource),
    #"Converted to Table" = Table.FromList(#"Removed Duplicates", Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Resource"}})
in
    #"Renamed Columns";

shared PeriodLIST = let
    Source = PeriodDemandTABLE,
    #"Removed Columns" = Table.RemoveColumns(Source,{"D"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates";

shared #"PeriodC##CapacityTABLE" = let
    Source = #"C##TABLE",
    #"Grouped PERIODC##" = Table.Group(Source, {"Period"}, {{"PeriodC##", each List.Sum([#"C##"]), type nullable number}})
in
    #"Grouped PERIODC##";

shared #"PeriodC##-D" = let
    Source = #"PeriodC##CapacityTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C##-D", each [#"PeriodC##"] - [D], type number),
    #"Inserted Division" = Table.AddColumn(#"Inserted Subtraction", "Division", each [#"PeriodC##"] / [D], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Division",{{"Division", "C##/D"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"D"})
in
    #"Removed Columns";

shared #"PeriodC##-DPos (Excess)TABLE" = let
    Source = #"PeriodC##-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C##-D"] > 0),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"C##-D", "ExcessC##-D"}})
in
    #"Renamed Columns";

shared #"PeriodMaxReduction - Empty Tables" = let
    Source = Table.NestedJoin(PeriodLIST, {"Period"}, #"PeriodC##-DPos (Excess)TABLE", {"Period"}, "PeriodC##-DPos (Excess)TABLE", JoinKind.LeftOuter),
    #"Expanded PeriodC##-DPos (Excess)TABLE" = Table.ExpandTableColumn(Source, "PeriodC##-DPos (Excess)TABLE", {"ExcessC##-D", "C##/D"}, {"ExcessC##-D", "C##/D"})
in
    #"Expanded PeriodC##-DPos (Excess)TABLE";

shared #"PeriodC##MaxReduction" = let
    Source = #"PeriodMaxReduction - Empty Tables",
    #"Filtered Rows" = Table.SelectRows(Source, each ([#"ExcessC##-D"] <> null)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period"}, #"PeriodA-DPos (Overallocation)", {"Period"}, "PeriodA-D", JoinKind.LeftOuter),
    #"Expanded PeriodA-D" = Table.ExpandTableColumn(#"Merged Queries", "PeriodA-D", {"A-D"}, {"A-D"}),
    #"Inserted Minimum" = Table.AddColumn(#"Expanded PeriodA-D", "PeriodMaxC##Reduction", each List.Min({ [#"A-D"]}))
in
    #"Inserted Minimum";

[ Description = "BUFFER" ]
// Query: ResPeriodMaxC##Reduction
// Purpose: Retain the existing final-reduction candidates only for Resources whose spare remains eligible.
shared #"ResPeriodMaxC##Reduction" = let
    Source = Table.NestedJoin(#"PeriodC##MaxReduction", {"Period"}, #"C##TABLE", {"Period"}, "C##TABLE", JoinKind.LeftOuter),
    #"Expanded C##TABLE" = Table.ExpandTableColumn(Source, "C##TABLE", {"Resource", "C##"}, {"Resource", "C##"}),
    AllowedResources = List.Buffer(Table.SelectRows(BResourceSpareEligibility, each [SpareCalculationAllowed])[Resource]),
    #"Filtered Rows" = Table.SelectRows(#"Expanded C##TABLE", each ([#"C##"] <> null) and not BRedistribution_RESULT[HasError] and List.Contains(AllowedResources, [Resource])),
    #"Merged Queries1" = Table.NestedJoin(#"Filtered Rows", {"Resource", "Period"}, BAllocation_CalculationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Filtered NOT ALLOCATED" = Table.SelectRows(#"Expanded ResPeriodAllocationTABLE", each ([Allocation] = null)),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered NOT ALLOCATED",{"Period", "Resource", "C##", "PeriodMaxC##Reduction","ExcessC##-D", "C##/D", "A-D"}),
    #"Merged Queries" = Table.NestedJoin(#"Reordered Columns", {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "ResPeriodOverallocationC##Reduction", JoinKind.LeftOuter),
    #"Expanded ResPeriodOverallocationC##Reduction" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodOverallocationC##Reduction", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    #"Filtered MAXREDUCTION C## <>0 or NULL" = Table.SelectRows(#"Expanded ResPeriodOverallocationC##Reduction", each ([#"MaxC##Reduction"] <> null) and ([#"C##"] <> 0)),
    // Period or res max reduction
    #"Inserted Minimum" = Table.AddColumn(#"Filtered MAXREDUCTION C## <>0 or NULL", "ResPeriodC##MAXReduction", each List.Min({[#"PeriodMaxC##Reduction"], [#"MaxC##Reduction"]}), type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Minimum",{"PeriodMaxC##Reduction", "MaxC##Reduction"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER;

// Query: C###TABLE B
// Purpose: Publish final B capacity with Resource-local spare exclusion and honest nonblocking diagnostics.
// Output: Preserve the existing final-capacity table interface used by C###SUM and C###MATRIX.
shared #"C###TABLE B" = let
    Source = #"C##TABLE",
    ReductionRows = if BFinalReduction_RESULT[HasError] then #table(type table [Resource = nullable number, Period = nullable number, MinimumAvailable = nullable number, KeepPR = nullable text], {}) else BFinalReduction_RESULT[Value],
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, ReductionRows, {"Resource", "Period"}, "SubtractAvailablilityTABLE", JoinKind.LeftOuter),
    #"Expanded SubtractAvailablilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "SubtractAvailablilityTABLE", {"MinimumAvailable", "KeepPR"}, {"MinimumAvailable", "KeepPR"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded SubtractAvailablilityTABLE",null,0,Replacer.ReplaceValue,{"MinimumAvailable"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Replaced Value", "C###", each if BFinalReduction_RESULT[HasError] and not [HasAllocationEvidence] then null else [#"C##"] - [MinimumAvailable], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"C##", "MinimumAvailable", "KeepPR"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Removed Columns",0,null,Replacer.ReplaceValue,{"C###"}),
    #"Merged Queries1" = Table.NestedJoin(#"Replaced Value1", {"Resource"}, ResMaxAvailability, {"Resource"}, "ResMaxAvailability", JoinKind.LeftOuter),
    #"Expanded ResMaxAvailability" = Table.ExpandTableColumn(#"Merged Queries1", "ResMaxAvailability", {"ResAvailability", "ResMaxAvail"}, {"ResAvailability", "ResMaxAvail"}),
    WithFinalPermission = Table.TransformColumns(#"Expanded ResMaxAvailability", {
        {"SpareCalculationAllowed", each _ = true and not BFinalReduction_RESULT[HasError], type logical},
        {"SpareIssueReason", each if not BFinalReduction_RESULT[HasError] then _ else Text.Combine(List.RemoveNulls({_, "Final reduction calculation failed; all spare excluded at C###"}), "; "), type nullable text}
    })
in
    WithFinalPermission;

shared #"C###SUM" = let
    Source = #"C###TABLE B",
    #"C###" = Source[#"C###"],
    #"Calculated Sum" = List.Sum(#"C###")
in
    #"Calculated Sum";

// Query: C###MATRIX
// Purpose: Preserve the legacy final-capacity matrix grain without using Stage 1 metadata as pivot keys.
shared #"C###MATRIX" = let
    Source = #"C###TABLE B",
    LegacyMatrixColumns = Table.SelectColumns(Source, {"Resource", "Period", "Role", "C###", "ResAvailability", "ResMaxAvail"}),
    #"Sorted Rows" = Table.Sort(LegacyMatrixColumns,{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "C###", List.Sum)
in
    #"Pivoted Column";

[ Description = "BUFFER" ]
shared #"ResourcePeriods-EmptyTABLE" = let
    Source = PeriodLIST,
    #"Added Custom" = Table.AddColumn(Source, "Custom", each ResourcesLIST),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Added Custom", "Custom", {"Resource"}, {"Resource"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Custom",{{"Period", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;

shared ResPeriodAllocationMATRIX = let
    Source = #"ResourcePeriods-EmptyTABLE",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Facility", "Date"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResPeriodAllocationTABLE",{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Allocation", List.Sum)
in
    #"Pivoted Column";

// Query: RolePathTABLE
// Purpose: Resolve this workbook's folder and dynamic role through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value rows used by RolePath and Role.
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
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "CapacityDistrib(B)-shifts.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate CapacityDistrib(B)-shifts.xlsx.",
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
