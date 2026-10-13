// Power Query from: CapacityDistrib(B)-shifts.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\AIN\CapacityDistrib(B)-shifts.xlsx
// Extracted: 2026-10-05T07:55:23.444Z

section Section1;

// Query: fnBFiniteNumber
// Purpose: Test finite numeric values without coercing text or null.
shared fnBFiniteNumber = (value as any) as logical =>
    try Value.Is(value, type number) and not Number.IsNaN(value) and value <> #infinity and value <> -#infinity otherwise false;

// Query: BSettingsMaximum_Status
// Purpose: Capture the configured roster shift maximum without rounding or stopping spare resolution.
shared BSettingsMaximum_Status = let
    ReadSetting = try let
        Setting = #"IMPORTSource Settings Data"{[Item = "MaxAvailability", Kind = "Table"]}[Data],
        HasColumn = Table.HasColumns(Setting, {"MaxAvailability"}),
        Value = if HasColumn and Table.RowCount(Setting) = 1 then Setting{0}[MaxAvailability] else null,
        Valid = fnBFiniteNumber(Value) and Value > 0 and Value = Number.RoundDown(Value)
    in
        if not HasColumn then [Status = "Error", Value = null, Reason = "Settings maximum column is missing"]
        else if Table.RowCount(Setting) <> 1 then [Status = "Error", Value = null, Reason = "Settings maximum must have exactly one row"]
        else if not Valid then [Status = "Error", Value = null, Reason = "Settings maximum must be a finite positive whole number"]
        else [Status = "Pass", Value = Value, Reason = null],
    Result = if ReadSetting[HasError] then [Status = "Error", Value = null, Reason = try ReadSetting[Error][Message] otherwise "Settings maximum could not be read"] else ReadSetting[Value]
in
    Result;

// Query: BAllocationSlots_Prepare
// Purpose: Count each positively allocated Resource/Period once, including allocations with zero availability.
// Notes: C# evidence includes positive facts below the legacy allocation import threshold. No capacity is invented.
shared BAllocationSlots_Prepare = let
    ImportedFacts = Table.SelectRows(BAllocation_CalculationTABLE, each [HasAllocationEvidence]),
    SavedEvidence = Table.SelectRows(BCappedAvailability_Prepare, each (try [HasAllocationEvidence] otherwise false) = true),
    Keys = Table.Combine({Table.SelectColumns(ImportedFacts, {"Resource", "Period"}), Table.SelectColumns(SavedEvidence, {"Resource", "Period"})}),
    UniqueSlots = Table.Distinct(Keys)
in
    Table.Buffer(UniqueSlots);

// Query: BAllocatedSlotTotals_Prepare
// Purpose: Count allocated slots once per Resource for reuse by every cap stage.
// Notes: Buffer only the narrow Resource/count result; unavailable allocated slots remain included.
shared BAllocatedSlotTotals_Prepare = let
    ResourceTotals = Table.Group(BAllocationSlots_Prepare, {"Resource"}, {{"AllocatedSlots", each Table.RowCount(_), Int64.Type}})
in
    Table.Buffer(ResourceTotals);

// Query: fnBCalendarIssues
// Purpose: Check stable planned clusters for five worked days and two full off days between clusters.
// Inputs: Distinct Resource/Day/PlannedCluster workday anchors plus retained optional days.
shared fnBCalendarIssues = (calendar as table, checkName as text) as table => let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text, Day = nullable number, PlannedCluster = nullable number, WorkedDays = nullable number, OffDays = nullable number],
    MakeIssue = (resource as number, cluster as number, day as number, worked as number, off as nullable number, reason as text) as record =>
        [Stage = "B", Check = checkName, Status = "Fail", Resource = resource, Role = Role, Name = null, Date = null, Period = null, Reason = reason,
         Action = "Retain allocation evidence; exclude affected Resource spare", Day = day, PlannedCluster = cluster, WorkedDays = worked, OffDays = off],
    UniqueDays = Table.Distinct(Table.SelectColumns(calendar, {"Resource", "Day", "PlannedCluster"})),
    Clusters = Table.Group(UniqueDays, {"Resource", "PlannedCluster"}, {{"Days", each List.Sort(List.Distinct([Day])), type list}}),
    WithStart = Table.AddColumn(Clusters, "StartDay", each List.Min([Days]), type number),
    WithEnd = Table.AddColumn(WithStart, "EndDay", each List.Max([Days]), type number),
    WithCount = Table.AddColumn(WithEnd, "WorkedDays", each List.Count([Days]), Int64.Type),
    // Reuse compact cluster facts for the internal and between-cluster checks.
    ClusterFacts = Table.Buffer(WithCount),
    ClusterRows = Table.ToRecords(ClusterFacts),
    InternalIssues = List.Combine(List.Transform(ClusterRows, (cluster) =>
        let
            BrokenInside = List.AnyTrue(List.Transform(List.Skip(List.Positions(cluster[Days])), (position) => cluster[Days]{position} - cluster[Days]{position - 1} > 2)),
            Reasons = (if cluster[WorkedDays] > 5 then {"Cluster exceeds five worked days"} else {}) & (if BrokenInside then {"A retained planned cluster was split by reduction"} else {})
        in List.Transform(Reasons, each MakeIssue(cluster[Resource], cluster[PlannedCluster], cluster[StartDay], cluster[WorkedDays], null, _)))),
    IndexedClusters = Table.AddIndexColumn(ClusterFacts, "ClusterOrder", 0, 1, Int64.Type),
    // Group once instead of filtering every Resource against all clusters; retain the original diagnostic order.
    ResourceClusters = Table.Group(IndexedClusters, {"Resource"}, {
        {"FirstClusterOrder", each List.Min([ClusterOrder]), Int64.Type},
        {"OrderedClusters", each Table.ToRecords(Table.Sort(_, {{"StartDay", Order.Ascending}, {"PlannedCluster", Order.Ascending}})), type list}
    }),
    OrderedResources = Table.Sort(ResourceClusters, {{"FirstClusterOrder", Order.Ascending}}),
    GapIssues = List.Combine(List.Transform(Table.ToRecords(OrderedResources), (resource) =>
        let
            Ordered = resource[OrderedClusters],
            Pairs = List.Skip(List.Positions(Ordered)),
            BadPairs = List.Select(Pairs, each Ordered{_}[StartDay] - Ordered{_ - 1}[EndDay] - 1 < 2)
        in List.Transform(BadPairs, (position) => MakeIssue(resource[Resource], Ordered{position}[PlannedCluster], Ordered{position}[StartDay], Ordered{position}[WorkedDays], Ordered{position}[StartDay] - Ordered{position - 1}[EndDay] - 1, "Fewer than two full off days between planned clusters")))),
    Issues = Table.FromRecords(InternalIssues & GapIssues, DetailType)
in
    Table.Buffer(Issues);

// Query: BSparePlanIssues_DETAILS
// Purpose: Validate saved Stage 3 day and cluster lineage before any B reduction; failures remain nonblocking.
shared BSparePlanIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    MakeIssue = (resource as nullable number, period as nullable number, reason as text) as record =>
        [Stage = "B", Check = "Saved spare plan", Status = "Error", Resource = resource, Role = Role, Name = null, Date = null, Period = period, Reason = reason,
         Action = if resource = null then "Exclude all spare calculations; retain allocation evidence" else "Exclude spare calculations for this Resource; retain allocation evidence"],
    Required = {"Day", "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday"},
    Raw = BCappedAvailability_Prepare,
    Missing = List.Difference(Required, Table.ColumnNames(Raw)),
    HasSchema = List.IsEmpty(Missing),
    Rows = if HasSchema then Table.ToRecords(Raw) else {},
    ValidRow = (row as record) as logical => try
        fnBFiniteNumber(row[Day]) and row[Day] > 0 and row[Day] = Number.RoundDown(row[Day])
        and Value.Is(row[SpareDayEligible], type logical) and Value.Is(row[SparePeriodSelected], type logical) and Value.Is(row[AllocationWorkday], type logical)
        and (row[SpareEligibilityReason] = null or Value.Is(row[SpareEligibilityReason], type text))
        and (row[ExtensionSide] = null or List.Contains({"Start", "Finish", "Spare-only"}, row[ExtensionSide]))
        and (row[PlannedCluster] = null or (fnBFiniteNumber(row[PlannedCluster]) and row[PlannedCluster] > 0 and row[PlannedCluster] = Number.RoundDown(row[PlannedCluster])))
        and (not row[AllocationWorkday] or row[PlannedCluster] <> null)
        and (if row[AvailabilityCapped] = null or row[AvailabilityCapped] = 0 or row[HasAllocationEvidence] then true
             else row[AvailabilityCapped] = 1 and row[SpareDayEligible] and row[SparePeriodSelected] and not row[AllocationWorkday] and row[PlannedCluster] <> null and row[ExtensionSide] <> null)
        otherwise false,
    BadRows = List.Select(Rows, each not ValidRow(_)),
    RowIssues = List.Transform(BadRows, each MakeIssue(try _[Resource] otherwise null, try _[Period] otherwise null, "Invalid Day, cluster lineage or whole selected spare cell")),
    ValidRows = if HasSchema then Table.FromRecords(List.Select(Rows, ValidRow), Value.Type(Raw)) else #table({"Resource", "Day", "Period", "PlannedCluster", "AllocationWorkday", "HasAllocationEvidence", "AvailabilityCapped"}, {}),
    DayGroups = Table.Group(ValidRows, {"Resource", "Day"}, {{"Rows", each _, type table}}),
    BadDays = Table.SelectRows(DayGroups, each
        let dayRows = [Rows], optional = Table.SelectRows(dayRows, each not [HasAllocationEvidence] and [AvailabilityCapped] <> null and [AvailabilityCapped] > 0)
        in Table.RowCount(optional) > 1 or (not Table.IsEmpty(optional) and List.Contains(dayRows[HasAllocationEvidence], true))
            or List.Count(List.Distinct(dayRows[PlannedCluster])) > 1 or List.Count(List.Distinct(dayRows[AllocationWorkday])) > 1),
    DayIssues = List.Transform(Table.ToRecords(BadDays), each MakeIssue(_[Resource], null, "Inconsistent day plan or optional spare on an allocated day")),
    BadResources = List.Distinct(List.RemoveNulls(List.Transform(RowIssues & DayIssues, each [Resource]))),
    CalendarRows = Table.SelectRows(ValidRows, each not List.Contains(BadResources, [Resource]) and ([AllocationWorkday] or (not [HasAllocationEvidence] and [AvailabilityCapped] <> null and [AvailabilityCapped] > 0))),
    CalendarIssues = fnBCalendarIssues(CalendarRows, "Saved spare plan"),
    SchemaIssues = if HasSchema then {} else {MakeIssue(null, null, "C# is missing Stage 3 metadata: " & Text.Combine(Missing, ", "))},
    Output = Table.Combine({Table.FromRecords(SchemaIssues & RowIssues & DayIssues, DetailType), CalendarIssues})
in
    Table.Buffer(Output);

// Query: fnBSelectReductions
// Purpose: Spend Resource and Period allowances only after a whole, exterior spare removal is accepted.
// Notes: Select the best current candidate in one pass; forced coverage still recalculates after every accepted removal.
shared fnBSelectReductions = (candidates as table, calendar as table, resourceLimits as table, periodLimits as table, coverage as table, forceCap as logical) as table => let
    ResultType = type table [Resource = number, Period = number, Day = number, PlannedCluster = number, ReductionAmount = number, RunningResTotal = number, RunningPeriodTotal = number, ForcedCap = logical, SurplusBefore = number, ShortageAfter = number],
    CandidateColumns = Table.SelectColumns(candidates, {"Resource", "Period", "Day", "PlannedCluster", "Amount", "NWDPriority", "PreferenceOrder"}),
    CandidateRows = List.Buffer(Table.ToRecords(CandidateColumns)),
    ClusterKey = (row as record) as text => Text.From(row[Resource]) & ":" & Text.From(row[PlannedCluster]),
    ClusterDays = Table.Group(calendar, {"Resource", "PlannedCluster"}, {{"Days", each List.Buffer(List.Distinct([Day])), type list}}),
    // Calendar lookup stays local to one cluster (at most five planned days), rather than rescanning the roster for each candidate.
    InitialCalendar = Record.FromList(ClusterDays[Days], List.Transform(Table.ToRecords(ClusterDays), ClusterKey)),
    ResourceBudget = Record.FromList(resourceLimits[Limit], List.Transform(resourceLimits[Resource], each Text.From(_))),
    PeriodBudget = Record.FromList(periodLimits[Limit], List.Transform(periodLimits[Period], each Text.From(_))),
    Capacity = Record.FromList(coverage[Capacity], List.Transform(coverage[Period], each Text.From(_))),
    Demand = Record.FromList(coverage[D], List.Transform(coverage[Period], each Text.From(_))),
    Lookup = (map as record, key as number, fallback as any) as any => Record.FieldOrDefault(map, Text.From(key), fallback),
    Put = (map as record, key as number, value as number) as record => Record.AddField(Record.RemoveFields(map, {Text.From(key)}, MissingField.Ignore), Text.From(key), value),
    CellKey = (row as record) as text => Text.From(row[Resource]) & ":" & Text.From(row[Period]),
    // Comparison signs reproduce the previous ascending/descending keys, including null NWD priorities.
    RankKeys = if forceCap then {{"SurplusBefore", -1}, {"ShortageAfter", 1}, {"NWDPriority", -1}, {"PreferenceOrder", 1}, {"Resource", 1}, {"Period", 1}}
        else {{"PreferenceOrder", 1}, {"Resource", 1}, {"Period", 1}},
    CompareChoices = (left as record, right as record) as number => let
        KeyComparisons = List.Transform(RankKeys, each Value.Compare(Record.Field(left, _{0}), Record.Field(right, _{0})) * _{1}),
        FirstDifference = List.First(List.Select(KeyComparisons, each _ <> 0), 0)
    in
        FirstDifference,
    NextAccepted = (state as record) as record => let
        Remaining = state[Remaining],
        // Keep only the best feasible record; avoid creating and sorting a scored table after each removal.
        Choice = List.Accumulate(Remaining, null, (best, row) =>
            let
                CurrentDays = Record.FieldOrDefault(state[Calendar], ClusterKey(row), {}),
                Exterior = not List.IsEmpty(CurrentDays) and (row[Day] = List.Min(CurrentDays) or row[Day] = List.Max(CurrentDays)),
                ResourceRemaining = Lookup(ResourceBudget, row[Resource], 0) - Lookup(state[ResourceSpend], row[Resource], 0),
                PeriodRemaining = Lookup(PeriodBudget, row[Period], 0) - Lookup(state[PeriodSpend], row[Period], 0),
                Feasible = Exterior and ResourceRemaining > 0 and (forceCap or (row[Amount] <= ResourceRemaining and row[Amount] <= PeriodRemaining)),
                CurrentCapacity = Lookup(Capacity, row[Period], 0) - Lookup(state[PeriodSpend], row[Period], 0),
                PeriodDemand = Lookup(Demand, row[Period], 0),
                ScoredChoice = Record.Combine({row, [SurplusBefore = List.Max({0, CurrentCapacity - PeriodDemand}), ShortageAfter = List.Max({0, PeriodDemand - (CurrentCapacity - row[Amount])})]})
            in
                if not Feasible then best
                else if best = null then ScoredChoice
                else if CompareChoices(ScoredChoice, best) < 0 then ScoredChoice else best),
        NewResourceSpend = if Choice = null then state[ResourceSpend] else Put(state[ResourceSpend], Choice[Resource], Lookup(state[ResourceSpend], Choice[Resource], 0) + Choice[Amount]),
        NewPeriodSpend = if Choice = null then state[PeriodSpend] else Put(state[PeriodSpend], Choice[Period], Lookup(state[PeriodSpend], Choice[Period], 0) + Choice[Amount]),
        Chosen = if Choice = null then null else [Resource = Choice[Resource], Period = Choice[Period], Day = Choice[Day], PlannedCluster = Choice[PlannedCluster], ReductionAmount = Choice[Amount],
            RunningResTotal = Lookup(NewResourceSpend, Choice[Resource], 0), RunningPeriodTotal = Lookup(NewPeriodSpend, Choice[Period], 0), ForcedCap = forceCap, SurplusBefore = Choice[SurplusBefore], ShortageAfter = Choice[ShortageAfter]],
        NewCalendar = if Choice = null then state[Calendar] else Record.AddField(Record.RemoveFields(state[Calendar], {ClusterKey(Choice)}), ClusterKey(Choice),
            List.RemoveItems(Record.Field(state[Calendar], ClusterKey(Choice)), {Choice[Day]})),
        // Buffer only surviving narrow candidate records so later passes do not replay a chain of prior filters.
        NextRemaining = if Choice = null then Remaining else List.Buffer(List.Select(Remaining, each CellKey(_) <> CellKey(Choice)))
    in [Active = Choice <> null, ResourceSpend = NewResourceSpend, PeriodSpend = NewPeriodSpend,
        Remaining = NextRemaining,
        Calendar = NewCalendar, Chosen = Chosen],
    Initial = [ResourceSpend = [], PeriodSpend = [], Remaining = CandidateRows, Calendar = InitialCalendar],
    Accepted = List.Generate(() => NextAccepted(Initial), each [Active], each NextAccepted(_), each [Chosen]),
    Output = Table.FromRecords(Accepted, ResultType)
in
    Table.Buffer(Output);

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

// Query: BDemandValueIssues_DETAILS
// Purpose: Report missing or invalid demand that prevents truthful surplus and shortage ranking.
shared BDemandValueIssues_DETAILS = let
    InvalidDemand = Table.SelectRows(BDemand_Prepare, each not fnBFiniteNumber([D]) or [D] < 0),
    Issues = List.Transform(Table.ToRecords(InvalidDemand), each [Stage = "B", Check = "Demand input", Status = "Error", Resource = null, Role = Role, Name = null,
        Date = _[Date], Period = _[Period], Reason = "Demand must be finite, numeric and nonnegative to rank cap reductions",
        Action = "Exclude all spare calculations; retain allocation evidence"]),
    CapacityPeriods = Table.Distinct(Table.SelectColumns(BCappedAvailability_Prepare, {"Period"})),
    MissingDemand = Table.NestedJoin(CapacityPeriods, {"Period"}, BDemand_Prepare, {"Period"}, "Demand", JoinKind.LeftAnti),
    MissingIssues = List.Transform(MissingDemand[Period], each [Stage = "B", Check = "Demand input", Status = "Error", Resource = null, Role = Role, Name = null,
        Date = null, Period = _, Reason = "Capacity Period has no demand row; omitted demand is not assumed to be zero",
        Action = "Exclude all spare calculations; retain allocation evidence"]),
    Output = Table.FromRecords(Issues & MissingIssues, type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text])
in
    Table.Buffer(Output);

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
    InvalidRows = Table.SelectRows(Contracts, each not fnBFiniteNumber([Effective Shift Cap]) or [Effective Shift Cap] <= 0
        or [Effective Shift Cap] <> Number.RoundDown([Effective Shift Cap]) or BSettingsMaximum_Status[Status] <> "Pass"
        or [Effective Shift Cap] > BSettingsMaximum_Status[Value] or [Limit Basis] = null or Text.Trim([Limit Basis]) = ""),
    InvalidIssues = List.Transform(Table.ToRecords(InvalidRows), each MakeIssue(_[Resource], "Effective Shift Cap or Limit Basis is invalid, exceeds Settings maximum, or cannot be checked against it")),
    SettingsIssues = if BSettingsMaximum_Status[Status] = "Pass" then {} else
        {Record.TransformFields(MakeIssue(null, BSettingsMaximum_Status[Reason]), {{"Status", each "Error"}, {"Check", each "Settings maximum"}})},
    PositiveOriginal = Table.SelectRows(BOriginalAvailability_Prepare, each [Availability] <> null and [Availability] > 0),
    PositiveCapped = Table.SelectRows(BCappedAvailability_Prepare, each [AvailabilityCapped] <> null and [AvailabilityCapped] > 0),
    PositiveResources = Table.Distinct(Table.Combine({Table.SelectColumns(PositiveOriginal, {"Resource"}), Table.SelectColumns(PositiveCapped, {"Resource"})})),
    MissingCaps = Table.NestedJoin(PositiveResources, {"Resource"}, Contracts, {"Resource"}, "Contracts", JoinKind.LeftAnti),
    MissingIssues = List.Transform(MissingCaps[Resource], each MakeIssue(_, "Positive availability has no Resource contract")),
    FallbackResources = Table.SelectRows(Contracts, each [Worker Record Status] = "Not found"),
    InvalidFallback = Table.NestedJoin(Table.Distinct(Table.SelectColumns(PositiveOriginal, {"Resource"})), {"Resource"}, FallbackResources, {"Resource"}, "Fallback", JoinKind.Inner),
    FallbackIssues = List.Transform(InvalidFallback[Resource], each MakeIssue(_, "Missing-worker fallback has positive AIN availability")),
    Output = Table.FromRecords(SettingsIssues & InvalidIssues & MissingIssues & FallbackIssues, DetailType)
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
shared BPreCalculationIssues_DETAILS = Table.Buffer(Table.Combine({BInputIssues_DETAILS, BDemandValueIssues_DETAILS, BContractIssues_DETAILS, BMaskIssues_DETAILS, BLineageIssues_DETAILS, BSparePlanIssues_DETAILS}));

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
shared CapacityDiagnostics_DETAILS = Table.Buffer(Table.Combine({BPreCalculationIssues_DETAILS, BOptionalCalculationIssues_DETAILS, BFinalCalendarIssues_DETAILS, BFinalCapReconciliationIssues_DETAILS, BFinalOutputIssues_DETAILS}));

// Query: CapacityDiagnostics_SUMMARY
// Purpose: Summarise real Pass, Fail and Error statuses without disguising failed inputs as successful calculations.
shared CapacityDiagnostics_SUMMARY = let
    CheckNames = {"Settings maximum", "Demand input", "Allocation input", "C# input", "Original availability input", "Priority input", "Resource contract keys", "Resource contracts", "Upstream spare permission", "Availability lineage", "Saved spare plan", "Redistribution calculation", "Final reduction calculation", "Final calendar", "Final cap reconciliation", "Published effective cap", "Published spare lineage", "Allocated capacity preserved"},
    // Summarise each check once rather than filtering the complete detail table for every check name.
    SummaryFacts = Table.Group(CapacityDiagnostics_DETAILS, {"Check"}, {
        {"Status", each if List.Contains([Status], "Error") then "Error" else if List.Contains([Status], "Fail") then "Fail"
            else if List.Contains([Status], "NotEvaluated") then "NotEvaluated" else "Fail", type text},
        {"Failures", each Table.RowCount(_), Int64.Type},
        {"Details", each Text.Combine(List.Distinct([Reason]), "; "), type nullable text}
    }),
    BufferedFacts = Table.Buffer(SummaryFacts),
    FactsByCheck = Record.FromList(Table.ToRecords(BufferedFacts), BufferedFacts[Check]),
    SummaryRows = List.Transform(CheckNames, (checkName) =>
        let Facts = Record.FieldOrDefault(FactsByCheck, checkName, [Status = "Pass", Failures = 0, Details = null])
        in [Stage = "B", Check = checkName, Status = Facts[Status], Failures = Facts[Failures], Details = Facts[Details]]),
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


// Query: fnBResourceCapTotals
// Purpose: Calculate remaining cap excess from allocated slots plus genuinely optional capacity, with no overlap.
shared fnBResourceCapTotals = (cells as table, amountColumn as text) as table => let
    Resources = Table.Distinct(Table.Combine({Table.SelectColumns(cells, {"Resource"}), Table.SelectColumns(BAllocationSlots_Prepare, {"Resource"})})),
    // Aggregate each input once instead of rescanning its complete rows separately for every Resource.
    OptionalColumns = Table.SelectColumns(cells, {"Resource", "HasAllocationEvidence", amountColumn}),
    OptionalCells = Table.SelectRows(OptionalColumns, each not [HasAllocationEvidence]),
    OptionalTotals = Table.Group(OptionalCells, {"Resource"}, {{"OptionalSlots", each List.Sum({0} & Table.Column(_, amountColumn)), type number}}),
    JoinedAllocated = Table.NestedJoin(Resources, {"Resource"}, BAllocatedSlotTotals_Prepare, {"Resource"}, "AllocatedTotals", JoinKind.LeftOuter),
    ExpandedAllocated = Table.ExpandTableColumn(JoinedAllocated, "AllocatedTotals", {"AllocatedSlots"}, {"AllocatedSlots"}),
    JoinedOptional = Table.NestedJoin(ExpandedAllocated, {"Resource"}, OptionalTotals, {"Resource"}, "OptionalTotals", JoinKind.LeftOuter),
    ExpandedOptional = Table.ExpandTableColumn(JoinedOptional, "OptionalTotals", {"OptionalSlots"}, {"OptionalSlots"}),
    WithOptional = Table.ReplaceValue(ExpandedOptional, null, 0, Replacer.ReplaceValue, {"AllocatedSlots", "OptionalSlots"}),
    WithTotal = Table.AddColumn(WithOptional, "TotalSlots", each [AllocatedSlots] + [OptionalSlots], type number),
    JoinedCap = Table.NestedJoin(WithTotal, {"Resource"}, BResourceContract, {"Resource"}, "Contract", JoinKind.LeftOuter),
    ExpandedCap = Table.ExpandTableColumn(JoinedCap, "Contract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    WithReduction = Table.AddColumn(ExpandedCap, "RequiredReduction", each if [Effective Shift Cap] = null then 0 else List.Max({0, [TotalSlots] - [Effective Shift Cap]}), type number)
in
    Table.Buffer(WithReduction);

// Query: BCapTotals_Capped
// Purpose: Reuse the C# allocation-plus-spare totals in redistribution and its priority preparation.
shared BCapTotals_Capped = fnBResourceCapTotals(#"ResPeriodAvailabilityCapped(C#)TABLE", "C#");

// Query: BCapTotals_Redistributed
// Purpose: Reuse the C## allocation-plus-spare totals in final reduction and its priority preparation.
shared BCapTotals_Redistributed = fnBResourceCapTotals(#"C##TABLE", "C##");

// Query: fnBWorkCalendar
// Purpose: Preserve original flagged workday anchors and add only retained optional workdays.
shared fnBWorkCalendar = (cells as table, amountColumn as text) as table => let
    WorkedRows = Table.SelectRows(cells, each fnBFiniteNumber([Day]) and [Day] > 0 and [Day] = Number.RoundDown([Day])
        and fnBFiniteNumber([PlannedCluster]) and [PlannedCluster] > 0 and [PlannedCluster] = Number.RoundDown([PlannedCluster]) and ([AllocationWorkday] = true
        or (not [HasAllocationEvidence] and Record.Field(_, amountColumn) <> null and Record.Field(_, amountColumn) > 0))),
    Calendar = Table.Distinct(Table.SelectColumns(WorkedRows, {"Resource", "Day", "PlannedCluster"}))
in
    Table.Buffer(Calendar);

// Query: fnBPeriodCoverage
// Purpose: Prepare current capacity and unchanged demand for ranking each accepted whole-shift removal.
shared fnBPeriodCoverage = (cells as table, amountColumn as text) as table => let
    Totals = Table.Group(cells, {"Period"}, {{"Capacity", each List.Sum({0} & Table.Column(_, amountColumn)), type number}}),
    JoinedDemand = Table.NestedJoin(Totals, {"Period"}, BDemand_Prepare, {"Period"}, "Demand", JoinKind.LeftOuter),
    // Missing demand is diagnosed before optional selection; keep it unknown here rather than inventing zero.
    WithDemand = Table.AddColumn(JoinedDemand, "D", each if Table.IsEmpty([Demand]) then null else [Demand]{0}[D], type nullable number),
    Output = Table.RemoveColumns(WithDemand, {"Demand"})
in
    Table.Buffer(Output);

// Query: fnBOptionalCandidates
// Purpose: Select whole chosen optional cells, retaining the saved Day and cluster ownership.
shared fnBOptionalCandidates = (cells as table, amountColumn as text) as table => let
    OptionalRows = Table.SelectRows(cells, each [SpareCalculationAllowed] = true and not [HasAllocationEvidence] and [SparePeriodSelected]
        and [SpareDayEligible] and not [AllocationWorkday] and [Day] <> null and [PlannedCluster] <> null and Record.Field(_, amountColumn) = 1),
    WithAmount = Table.AddColumn(OptionalRows, "Amount", each Record.Field(_, amountColumn), type number),
    Selected = Table.SelectColumns(WithAmount, {"Resource", "Period", "Day", "PlannedCluster", "Amount"}),
    JoinedPriority = Table.NestedJoin(Selected, {"Resource", "Period"}, BPriority_Prepare, {"Resource", "Period"}, "Priority", JoinKind.LeftOuter),
    WithPriority = Table.AddColumn(JoinedPriority, "NWDPriority", each if Table.IsEmpty([Priority]) then 0 else [Priority]{0}[NWDPriority], type number),
    Output = Table.RemoveColumns(WithPriority, {"Priority"})
in
    Table.Buffer(Output);

// Query: BRedistributionSelection
// Purpose: Jointly accept redistribution removals within the Resource cap and Period demand allowances.
shared BRedistributionSelection = let
    Cells = #"ResPeriodAvailabilityCapped(C#)TABLE",
    OptionalCandidates = fnBOptionalCandidates(Cells, "C#"),
    OriginalOrder = Table.SelectColumns(#"PrioritiseRedistribAvail-Distrib", {"Resource", "Period", "IndexAllRowPrioritySort"}),
    JoinedOrder = Table.NestedJoin(OptionalCandidates, {"Resource", "Period"}, OriginalOrder, {"Resource", "Period"}, "Preference", JoinKind.Inner),
    WithOrder = Table.ExpandTableColumn(JoinedOrder, "Preference", {"IndexAllRowPrioritySort"}, {"PreferenceOrder"}),
    ResourceLimits = Table.RenameColumns(Table.SelectColumns(BCapTotals_Capped, {"Resource", "RequiredReduction"}), {{"RequiredReduction", "Limit"}}),
    PeriodLimits = Table.RenameColumns(Table.SelectColumns(#"PeriodC#-DPos (Excess)TABLE", {"Period", "ExcessC#-D"}), {{"ExcessC#-D", "Limit"}}),
    Selected = fnBSelectReductions(WithOrder, fnBWorkCalendar(Cells, "C#"), ResourceLimits, PeriodLimits, fnBPeriodCoverage(Cells, "C#"), false)
in
    Selected;

// Query: BNormalFinalReductionSelection
// Purpose: Apply the final demand-surplus allowances with accepted-only joint Resource and Period state.
shared BNormalFinalReductionSelection = let
    Cells = #"C##TABLE",
    OptionalCandidates = fnBOptionalCandidates(Cells, "C##"),
    // Preserve the existing Period / allocation-excess / cap-reduction / Resource ordering; only budget acceptance changes.
    OriginalOrder = Table.SelectColumns(#"PrioritiseReductionAvail-Setup", {"Resource", "Period", "IndexAllRows"}),
    JoinedOrder = Table.NestedJoin(OptionalCandidates, {"Resource", "Period"}, OriginalOrder, {"Resource", "Period"}, "Preference", JoinKind.Inner),
    WithOrder = Table.ExpandTableColumn(JoinedOrder, "Preference", {"IndexAllRows"}, {"PreferenceOrder"}),
    ResourceLimits = Table.RenameColumns(Table.SelectColumns(BCapTotals_Redistributed, {"Resource", "RequiredReduction"}), {{"RequiredReduction", "Limit"}}),
    RawPeriodLimits = Table.SelectRows(#"PeriodC##MaxReduction", each fnBFiniteNumber([#"PeriodMaxC##Reduction"]) and [#"PeriodMaxC##Reduction"] > 0),
    BoundedPeriodLimits = Table.AddColumn(RawPeriodLimits, "Limit", each List.Min({[#"PeriodMaxC##Reduction"], [#"ExcessC##-D"]}), type number),
    PeriodLimits = Table.SelectColumns(BoundedPeriodLimits, {"Period", "Limit"}),
    Selected = fnBSelectReductions(WithOrder, fnBWorkCalendar(Cells, "C##"), ResourceLimits, PeriodLimits, fnBPeriodCoverage(Cells, "C##"), false)
in
    Selected;

// Query: BPostDemandReduction_Prepare
// Purpose: Prepare current capacity after normal accepted reductions for the independent hard-cap selection.
shared BPostDemandReduction_Prepare = let
    Joined = Table.NestedJoin(#"C##TABLE", {"Resource", "Period"}, BNormalFinalReductionSelection, {"Resource", "Period"}, "Reduction", JoinKind.LeftOuter),
    WithRemaining = Table.AddColumn(Joined, "RemainingCapacity", each [#"C##"] - (if Table.IsEmpty([Reduction]) then 0 else [Reduction]{0}[ReductionAmount]), type nullable number),
    Output = Table.RemoveColumns(WithRemaining, {"Reduction"})
in
    Table.Buffer(Output);

// Query: BHardCapReductionSelection
// Purpose: Remove remaining excess spare even without Period surplus, preferring surplus then least resulting shortage.
// Notes: Coverage is recalculated after each accepted exterior trim; allocations are never candidates.
shared BHardCapReductionSelection = let
    Cells = BPostDemandReduction_Prepare,
    OptionalCandidates = fnBOptionalCandidates(Cells, "RemainingCapacity"),
    RankedCandidates = Table.Sort(OptionalCandidates, {{"NWDPriority", Order.Descending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    WithOrder = Table.AddIndexColumn(RankedCandidates, "PreferenceOrder", 1, 1, Int64.Type),
    ResourceLimits = Table.RenameColumns(Table.SelectColumns(fnBResourceCapTotals(Cells, "RemainingCapacity"), {"Resource", "RequiredReduction"}), {{"RequiredReduction", "Limit"}}),
    EmptyPeriodLimits = #table(type table [Period = number, Limit = number], {}),
    Selected = fnBSelectReductions(WithOrder, fnBWorkCalendar(Cells, "RemainingCapacity"), ResourceLimits, EmptyPeriodLimits, fnBPeriodCoverage(Cells, "RemainingCapacity"), true)
in
    Selected;

// Query: BFinalReductionSelection
// Purpose: Publish one accepted whole-cell reduction per Resource/Period across normal and forced cap stages.
shared BFinalReductionSelection = Table.Buffer(Table.Combine({BNormalFinalReductionSelection, BHardCapReductionSelection}));

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
// Purpose: Show accepted joint-budget redistribution decisions in the existing candidate priority order.
shared ReDistributeResAvailability = let
    CandidateOrder = #"PrioritiseRedistribAvail-Distrib",
    JoinedLimits = Table.NestedJoin(CandidateOrder, {"Resource"}, #"ResIndexLimits (ResMaxReduction)", {"Resource"}, "Limits", JoinKind.LeftOuter),
    ExpandedLimits = Table.ExpandTableColumn(JoinedLimits, "Limits", {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}),
    WithAmount = Table.AddColumn(ExpandedLimits, "C#", each [Unassigned], type number),
    JoinedSelection = Table.NestedJoin(WithAmount, {"Resource", "Period"}, BRedistributionSelection, {"Resource", "Period"}, "Accepted", JoinKind.LeftOuter),
    WithDecision = Table.AddColumn(JoinedSelection, "KeepRunningTotal", each if Table.IsEmpty([Accepted]) then "Remove" else "Keep", type text),
    ExpandedSpend = Table.ExpandTableColumn(WithDecision, "Accepted", {"RunningResTotal", "RunningPeriodTotal"}, {"RunningResTotal", "RunningPeriodTotal"}),
    WithPosition = Table.AddColumn(ExpandedSpend, "Subtraction", each [IndexAllRowPrioritySort] - [StartResIndex] + 1, Int64.Type)
in
    WithPosition;

[ Description = "BUFFER" ]
// Query: ReDistribResAvailabilityTABLE
// Purpose: Preserve the indexed audit interface with totals from accepted joint decisions.
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
// Purpose: Publish only jointly accepted whole-cell redistribution removals.
shared #"ReDistribPeriodAvailbility TABLE" = let
    AcceptedRows = Table.SelectRows(ReDistribResAvailabilityTABLE, each [KeepRunningTotal] = "Keep"),
    NamedPeriodSpend = Table.RenameColumns(AcceptedRows, {{"RunningPeriodTotal", "PeriodRunningTotal"}}),
    WithDecision = Table.AddColumn(NamedPeriodSpend, "Keep RunningTotal", each "Keep", type text)
in
    Table.Buffer(WithDecision);

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
// Purpose: Calculate allocated slots plus optional C## cap excess without counting overlap twice.
shared #"ResRosterAvailabilityC##CapReduction" = let
    ResourceTotals = BCapTotals_Redistributed,
    Output = Table.RenameColumns(Table.SelectColumns(ResourceTotals, {"Resource", "RequiredReduction"}), {{"RequiredReduction", "MaxC##Reduction"}})
in
    Output;

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
    PeriodShiftLimit = BSettingsMaximum_Status[Value]
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
    ValidCaps = Table.SelectRows(PreparedContracts, each BSettingsMaximum_Status[Status] = "Pass" and fnBFiniteNumber([Effective Shift Cap]) and [Effective Shift Cap] > 0
        and [Effective Shift Cap] = Number.RoundDown([Effective Shift Cap]) and [Effective Shift Cap] <= BSettingsMaximum_Status[Value] and [Limit Basis] <> null and Text.Trim([Limit Basis]) <> ""),
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
// Purpose: Calculate allocated slots plus optional C# cap excess without counting overlap twice.
shared ResRosterAvailabilityCapReduction = let
    ResourceTotals = BCapTotals_Capped,
    Output = Table.RenameColumns(Table.SelectColumns(ResourceTotals, {"Resource", "RequiredReduction"}), {{"RequiredReduction", "AvailabilityReduction"}})
in
    Output;

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
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAllocationTABLE", {"Allocation", "HasAllocationEvidence"}, {"Allocation", "AllocationEvidenceInFacts"}),
    #"Merged Queries5" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource"}, ResRosterAvailabilityCapReduction, {"Resource"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction1" = Table.ExpandTableColumn(#"Merged Queries5", "ResRosterAvailabilityCapReduction", {"AvailabilityReduction"}, {"ResAvailabilityReduction"}),
    #"Merged Queries4" = Table.NestedJoin(#"Expanded ResRosterAvailabilityCapReduction1", {"Resource", "Period"}, ResPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction" = Table.ExpandTableColumn(#"Merged Queries4", "ResRosterAvailabilityCapReduction", {"C#", "HasAllocationEvidence"}, {"C#", "HasAllocationEvidence"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResRosterAvailabilityCapReduction",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    // Neither imported facts nor lower-threshold upstream allocation evidence can become optional reductions.
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each
        ([ResAvailabilityReduction] <> null 
        and [ResAvailabilityReduction] <> 0) 
        and ([#"C#/D"] <> null 
        and [#"C#/D"] <> 0) 
        and ([#"C#"] <> 0) 
        and ([#"C#"] <> null)
        and [HasAllocationEvidence] <> true
        and [AllocationEvidenceInFacts] <> true
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
    UsableOrdinal = (values as list) as nullable number => let
        Value = try List.First(values) otherwise null,
        Valid = fnBFiniteNumber(Value) and Value > 0 and Value = Number.RoundDown(Value)
    in if Valid then Value else null,
    Source = BCappedAvailability_Prepare,
    RoleRows = Table.SelectRows(Source, each [Role] = Role and [AvailabilityCapped] <> null),
    // A duplicate input key is already a Resource issue. Preserve one existing allocated-cell amount, never multiply it through joins.
    GroupedCells = Table.Group(RoleRows, {"Resource", "Period"}, {
        {"Role", each List.First([Role]), type nullable text},
        {"C#", each List.Max([AvailabilityCapped]), type nullable number},
        {"UpstreamAllocationEvidence", each List.AnyTrue(List.Transform(Table.ToRecords(_), each (try _[HasAllocationEvidence] otherwise false) = true)), type logical},
        {"OriginalAvailability", each List.Max(List.Transform(Table.ToRecords(_), each try _[OriginalAvailability] otherwise null)), type nullable number},
        {"Day", each try UsableOrdinal([Day]) otherwise null, type nullable number},
        {"PlannedCluster", each try UsableOrdinal([PlannedCluster]) otherwise null, type nullable number},
        {"AllocationWorkday", each (try List.First([AllocationWorkday]) otherwise false) = true, type logical},
        {"SpareDayEligible", each (try List.First([SpareDayEligible]) otherwise false) = true, type logical},
        {"SparePeriodSelected", each (try List.First([SparePeriodSelected]) otherwise false) = true, type logical},
        {"ExtensionSide", each let Value = try List.First([ExtensionSide]) otherwise null in if Value.Is(Value, type text) then Value else null, type nullable text},
        {"SpareEligibilityReason", each let Value = try List.First([SpareEligibilityReason]) otherwise null in if Value.Is(Value, type text) then Value else null, type nullable text}
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
// Purpose: Expose normal final reductions accepted jointly against Resource and Period allowances.
shared SubtractOverallatedPeriods = let
    SelectedRows = BNormalFinalReductionSelection,
    WithAmount = Table.RenameColumns(SelectedRows, {{"ReductionAmount", "MinimumAvailable"}}),
    WithDecision = Table.AddColumn(WithAmount, "KeepRunningTotal", each "Keep", type text)
in
    WithDecision;

// Query: ResIndex-Steup
// Purpose: Preserve the indexed audit interface for accepted normal and forced cap removals.
shared #"ResIndex-Steup" = let
    SelectedRows = BFinalReductionSelection,
    WithAmount = Table.RenameColumns(SelectedRows, {{"ReductionAmount", "MinimumAvailable"}}),
    JoinedLimit = Table.NestedJoin(WithAmount, {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "Limit", JoinKind.LeftOuter),
    ExpandedLimit = Table.ExpandTableColumn(JoinedLimit, "Limit", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    WithIndex = Table.AddIndexColumn(ExpandedLimit, "IndexAllRows", 2, 1, Int64.Type)
in
    WithIndex;

// Query: ResIndexLimits
// Purpose: Retain Resource-grain audit counts and cap limits for accepted final reductions.
shared ResIndexLimits = let
    AcceptedRows = #"ResIndex-Steup",
    Limits = Table.Group(AcceptedRows, {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRows]), type number}, {"ResExcessAvailable", each List.Max([#"MaxC##Reduction"]), type number}, {"ResAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    Limits;

// Query: SubtractOverallocatedResources
// Purpose: Publish accepted normal and forced cap removals for final C### subtraction.
shared SubtractOverallocatedResources = let
    AcceptedRows = BFinalReductionSelection,
    WithAmount = Table.RenameColumns(AcceptedRows, {{"ReductionAmount", "MinimumAvailable"}}),
    WithDecision = Table.AddColumn(WithAmount, "KeepPR", each "Keep", type text)
in
    Table.Buffer(WithDecision);

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

// Query: BFinalCapacity_Prepare
// Purpose: Subtract accepted normal and forced whole-cell reductions while preserving allocated capacity.
shared BFinalCapacity_Prepare = let
    Source = #"C##TABLE",
    ReductionRows = if BFinalReduction_RESULT[HasError] then #table(type table [Resource = nullable number, Period = nullable number, MinimumAvailable = nullable number], {}) else BFinalReduction_RESULT[Value],
    JoinedReduction = Table.NestedJoin(Source, {"Resource", "Period"}, ReductionRows, {"Resource", "Period"}, "Reduction", JoinKind.LeftOuter),
    WithFinalCapacity = Table.AddColumn(JoinedReduction, "C###", each
        if [HasAllocationEvidence] then [#"C##"]
        else if BFinalReduction_RESULT[HasError] then null
        else [#"C##"] - (if Table.IsEmpty([Reduction]) then 0 else [Reduction]{0}[MinimumAvailable]), type nullable number),
    RemovedCalculationColumns = Table.RemoveColumns(WithFinalCapacity, {"C##", "Reduction"}),
    ZeroToNull = Table.ReplaceValue(RemovedCalculationColumns, 0, null, Replacer.ReplaceValue, {"C###"})
in
    Table.Buffer(ZeroToNull);

// Query: BFinalCalendarIssues_DETAILS
// Purpose: Check the retained calendar before publication; unsafe optional capacity is excluded per Resource.
shared BFinalCalendarIssues_DETAILS = let
    RuleIssues = fnBCalendarIssues(fnBWorkCalendar(BFinalCapacity_Prepare, "C###"), "Final calendar"),
    IncompletePlan = Table.Combine({Table.SelectRows(BSparePlanIssues_DETAILS, each [Status] = "Error"),
        Table.SelectRows(BInputIssues_DETAILS, each [Check] = "C# input"), Table.SelectRows(BMaskIssues_DETAILS, each [Status] = "Error")}),
    CoverageIssues = Table.TransformColumns(IncompletePlan, {{"Check", each "Final calendar", type text}, {"Status", each "NotEvaluated", type text},
        {"Reason", each "Final calendar cannot be fully checked: " & _, type text}}),
    Output = Table.Combine({RuleIssues, CoverageIssues})
in
    Table.Buffer(Output);

// Query: BFinalCapReconciliationIssues_DETAILS
// Purpose: Flag unresolved cap excess with spare still present before the Resource-local publication fallback.
shared BFinalCapReconciliationIssues_DETAILS = let
    Totals = fnBResourceCapTotals(BFinalCapacity_Prepare, "C###"),
    Failures = Table.SelectRows(Totals, each [Effective Shift Cap] <> null and [TotalSlots] > [Effective Shift Cap] and [OptionalSlots] > 0),
    Details = List.Transform(Table.ToRecords(Failures), each [Stage = "B", Check = "Final cap reconciliation", Status = "Error", Resource = _[Resource], Role = Role,
        Name = null, Date = null, Period = null, Reason = "Legal whole-cell trimming did not resolve cap excess; remaining Resource spare excluded",
        Action = "Exclude remaining Resource spare; retain allocation evidence", AllocatedSlots = _[AllocatedSlots], OptionalSlots = _[OptionalSlots], EffectiveCap = _[Effective Shift Cap], ExcessSlots = _[RequiredReduction]]),
    Output = Table.FromRecords(Details, type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text, AllocatedSlots = number, OptionalSlots = number, EffectiveCap = number, ExcessSlots = number])
in
    Table.Buffer(Output);

// Query: BFinalCap_CHECK
// Purpose: Show the published allocation-plus-spare footprint against each usable effective cap.
// Notes: Preexisting allocated slots above cap remain visible; they are never removed to produce a Pass.
shared BFinalCap_CHECK = let
    Totals = fnBResourceCapTotals(#"C###TABLE B", "C###"),
    EvidenceInputIssues = Table.SelectRows(BInputIssues_DETAILS, each [Check] = "Allocation input" or [Check] = "C# input"),
    EvidenceMaskIssues = Table.SelectRows(BMaskIssues_DETAILS, each [Status] = "Error"),
    IncompleteEvidence = Table.Combine({EvidenceInputIssues, EvidenceMaskIssues}),
    GlobalIncomplete = not Table.IsEmpty(Table.SelectRows(IncompleteEvidence, each [Resource] = null and [Action] <> "Allocation skipped; evidence retained; other Resources continue")),
    IncompleteResources = List.Distinct(List.RemoveNulls(IncompleteEvidence[Resource])),
    WithStatus = Table.AddColumn(Totals, "Status", each if [Effective Shift Cap] = null or GlobalIncomplete or List.Contains(IncompleteResources, [Resource]) then "NotEvaluated"
        else if [TotalSlots] > [Effective Shift Cap] then "Fail" else "Pass", type text),
    WithOrigin = Table.AddColumn(WithStatus, "Origin", each if [Effective Shift Cap] = null then "Unknown cap" else if [AllocatedSlots] > [Effective Shift Cap] then "Preexisting allocations" else if [RequiredReduction] > 0 then "Optional spare" else null, type nullable text)
in
    Table.Buffer(WithOrigin);

// Query: BFinalOutputIssues_DETAILS
// Purpose: Publish cap, allocation preservation and spare-lineage checks without stopping other Resources.
shared BFinalOutputIssues_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Status = text, Resource = nullable number, Role = nullable text, Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
    MakeIssue = (check as text, status as text, resource as nullable number, period as nullable number, reason as text) as record =>
        [Stage = "B", Check = check, Status = status, Resource = resource, Role = Role, Name = null, Date = null, Period = period, Reason = reason, Action = "Manual checking; allocation evidence retained; other Resources continue"],
    CapFailures = Table.SelectRows(BFinalCap_CHECK, each [Status] <> "Pass"),
    CapIssues = List.Transform(Table.ToRecords(CapFailures), each MakeIssue("Published effective cap", _[Status], _[Resource], null,
        if _[Status] = "NotEvaluated" then "Published cap cannot be fully checked: usable contract or complete allocation evidence is missing"
        else "Allocated plus optional slots exceed effective cap; origin: " & _[Origin] & "; total=" & Text.From(_[TotalSlots]) & "; cap=" & Text.From(_[Effective Shift Cap]))),
    Output = #"C###TABLE B",
    BadSpare = Table.SelectRows(Output, each not [HasAllocationEvidence] and [#"C###"] <> null and [#"C###"] > 0 and
        (not [SpareCalculationAllowed] or not [SpareDayEligible] or not [SparePeriodSelected] or [#"C###"] <> 1 or [OriginalAvailability] = null or [OriginalAvailability] <= 0)),
    SpareIssues = List.Transform(Table.ToRecords(BadSpare), each MakeIssue("Published spare lineage", "Fail", _[Resource], _[Period], "Retained spare is not a whole eligible selected original-availability cell")),
    Protected = Table.SelectRows(#"ResPeriodAvailabilityCapped(C#)TABLE", each [HasAllocationEvidence]),
    JoinedProtected = Table.NestedJoin(Protected, {"Resource", "Period"}, Output, {"Resource", "Period"}, "Published", JoinKind.LeftOuter),
    LostAllocation = Table.SelectRows(JoinedProtected, each Table.RowCount([Published]) <> 1 or not ([#"C#"] = [Published]{0}[#"C###"])),
    AllocationIssues = List.Transform(Table.ToRecords(LostAllocation), each MakeIssue("Allocated capacity preserved", "Fail", _[Resource], _[Period], "Allocated C# capacity was changed or lost before publication")),
    MissingCapped = BCappedAvailabilityInput_RESULT[HasError] or not Table.HasColumns(BCappedAvailability_Prepare, {"HasAllocationEvidence", "OriginalAvailability"}),
    MissingOriginal = BOriginalAvailabilityInput_RESULT[HasError] or not Table.HasColumns(BOriginalAvailability_Prepare, {"HasAllocationEvidence", "OriginalAvailability"}),
    EvaluationIssues = (if MissingCapped then {MakeIssue("Allocated capacity preserved", "NotEvaluated", null, null, "Required C# allocation evidence could not be read")}
        else {}) & (if MissingCapped or MissingOriginal then {MakeIssue("Published spare lineage", "NotEvaluated", null, null, "Required source availability or allocation evidence could not be read")} else {}),
    CappedCoverageFailures = Table.Combine({Table.SelectRows(BInputIssues_DETAILS, each [Check] = "C# input"), Table.SelectRows(BMaskIssues_DETAILS, each [Status] = "Error")}),
    OriginalCoverageFailures = Table.SelectRows(BInputIssues_DETAILS, each [Check] = "Original availability input"),
    CapCoverageFailures = Table.Combine({CappedCoverageFailures,
        Table.SelectRows(BInputIssues_DETAILS, each ([Check] = "Allocation input" and [Action] <> "Allocation skipped; evidence retained; other Resources continue") or [Check] = "Resource contract keys"),
        Table.SelectRows(BContractIssues_DETAILS, each [Status] = "Error")}),
    CoverageIssues = List.Transform(Table.ToRecords(CappedCoverageFailures), each MakeIssue("Allocated capacity preserved", "NotEvaluated", _[Resource], _[Period], "C# allocation evidence cannot be completely reconciled: " & _[Reason]))
        & List.Transform(Table.ToRecords(Table.Combine({CappedCoverageFailures, OriginalCoverageFailures})), each MakeIssue("Published spare lineage", "NotEvaluated", _[Resource], _[Period], "Source lineage cannot be completely reconciled: " & _[Reason]))
        & List.Transform(Table.ToRecords(CapCoverageFailures), each MakeIssue("Published effective cap", "NotEvaluated", _[Resource], _[Period], "Allocation-plus-spare cap footprint cannot be completely checked: " & _[Reason])),
    Issues = Table.Distinct(Table.FromRecords(CapIssues & SpareIssues & AllocationIssues & EvaluationIssues & CoverageIssues, DetailType))
in
    Table.Buffer(Issues);

// Query: BPublicationResourceIssues_Prepare
// Purpose: Prepare publication exclusions and ordered reasons once per Resource for the final capacity join.
// Notes: A null Resource retains its diagnostic reasons but does not become a Resource-local exclusion.
shared BPublicationResourceIssues_Prepare = let
    PublicationIssues = Table.Combine({BFinalCalendarIssues_DETAILS, BFinalCapReconciliationIssues_DETAILS}),
    IndexedIssues = Table.AddIndexColumn(PublicationIssues, "PublicationIssueOrder", 0, 1, Int64.Type),
    GroupedIssues = Table.Group(IndexedIssues, {"Resource"}, {
        {"PublicationReasons", each let
            OrderedIssues = Table.Sort(_, {{"PublicationIssueOrder", Order.Ascending}}),
            // These short lists are reused by every period row for the Resource.
            Reasons = List.Buffer(List.Distinct(List.RemoveNulls(OrderedIssues[Reason])))
        in Reasons, type list}
    }),
    WithExclusion = Table.AddColumn(GroupedIssues, "PublicationSpareExcluded", each [Resource] <> null, type logical)
in
    Table.Buffer(WithExclusion);

// Query: C###TABLE B
// Purpose: Publish final B capacity with Resource-local spare exclusion and honest nonblocking diagnostics.
// Output: Preserve the existing final-capacity table interface used by C###SUM and C###MATRIX.
shared #"C###TABLE B" = let
    Source = BFinalCapacity_Prepare,
    JoinedPublicationIssues = Table.NestedJoin(Source, {"Resource"}, BPublicationResourceIssues_Prepare, {"Resource"}, "PublicationIssues", JoinKind.LeftOuter),
    ExpandedPublicationIssues = Table.ExpandTableColumn(JoinedPublicationIssues, "PublicationIssues", {"PublicationReasons", "PublicationSpareExcluded"}, {"PublicationReasons", "PublicationSpareExcluded"}),
    WithPublicationPermission = Table.ReplaceValue(ExpandedPublicationIssues, null, false, Replacer.ReplaceValue, {"PublicationSpareExcluded"}),
    WithSafeCapacity = Table.AddColumn(WithPublicationPermission, "PublishedCapacity", each if [HasAllocationEvidence] then [#"C###"]
        else if [PublicationSpareExcluded] then null else [#"C###"], type nullable number),
    RemovedProvisionalCapacity = Table.RemoveColumns(WithSafeCapacity, {"C###"}),
    NamedPublishedCapacity = Table.RenameColumns(RemovedProvisionalCapacity, {{"PublishedCapacity", "C###"}}),
    WithPermission = Table.AddColumn(NamedPublishedCapacity, "FinalSparePermission", each [SpareCalculationAllowed] and not BFinalReduction_RESULT[HasError]
        and not [PublicationSpareExcluded], type logical),
    WithReason = Table.AddColumn(WithPermission, "FinalSpareReason", (row) => Text.Combine(List.Distinct(List.RemoveNulls({row[SpareIssueReason]}
        & (if BFinalReduction_RESULT[HasError] then {"Final reduction calculation failed; all spare excluded at C###"} else {})
        & (if row[PublicationReasons] = null then {} else row[PublicationReasons]))), "; "), type text),
    RemovedPublicationColumns = Table.RemoveColumns(WithReason, {"PublicationReasons", "PublicationSpareExcluded"}),
    RemovedOldPermission = Table.RemoveColumns(RemovedPublicationColumns, {"SpareCalculationAllowed", "SpareIssueReason"}),
    NamedFinalPermission = Table.RenameColumns(RemovedOldPermission, {{"FinalSparePermission", "SpareCalculationAllowed"}, {"FinalSpareReason", "SpareIssueReason"}}),
    JoinedAvailabilityAudit = Table.NestedJoin(NamedFinalPermission, {"Resource"}, ResMaxAvailability, {"Resource"}, "AvailabilityAudit", JoinKind.LeftOuter),
    Output = Table.ExpandTableColumn(JoinedAvailabilityAudit, "AvailabilityAudit", {"ResAvailability", "ResMaxAvail"}, {"ResAvailability", "ResMaxAvail"})
in
    Table.Buffer(Output);

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
