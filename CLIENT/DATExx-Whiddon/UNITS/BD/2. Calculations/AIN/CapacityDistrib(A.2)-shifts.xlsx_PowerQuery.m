// Power Query from: CapacityDistrib(A.2)-shifts.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\AIN\CapacityDistrib(A.2)-shifts.xlsx
// Extracted: 2026-10-05T07:55:20.402Z

section Section1;

// Query: A2OptionalInputAssessment
// Purpose: Capture an optional input failure and identify Resources whose optional spare cannot be calculated.
// Notes: Invalid or ambiguous keys are excluded only with a reported Resource exclusion; an unidentified issue excludes all optional spare.
shared A2OptionalInputAssessment = (evaluate as function, columns as list, periodColumn as text) as record =>
let
    InputAttempt = try let
        Source = evaluate(),
        // Validate required fields while retaining the source's existing optional handoff columns.
        InputColumns = List.Union({Table.ColumnNames(Source), columns}),
        ValidatedInput = Table.SelectColumns(Source, InputColumns),
        BufferedInput = Table.Buffer(ValidatedInput)
    in BufferedInput,
    EmptyInput = #table(columns, {}),
    Input = if InputAttempt[HasError] then EmptyInput else InputAttempt[Value],
    ErrorRows = Table.SelectRowsWithErrors(Input, columns),
    ErrorCount = Table.RowCount(ErrorRows),
    ErrorResourceKeys = Table.SelectColumns(ErrorRows, {"Resource"}),
    ErrorResourceKeysWithoutErrors = Table.RemoveRowsWithErrors(ErrorResourceKeys, {"Resource"}),
    ErrorResourceRows = Table.SelectRows(ErrorResourceKeysWithoutErrors, each [Resource] <> null),
    RowsWithoutErrors = Table.RemoveRowsWithErrors(Input, columns),
    MissingKeyRows = Table.SelectRows(RowsWithoutErrors, each [Resource] = null or Record.Field(_, periodColumn) = null),
    MissingKeyCount = Table.RowCount(MissingKeyRows),
    MissingKeyResourceKeys = Table.SelectColumns(MissingKeyRows, {"Resource"}),
    MissingKeyResourceRows = Table.SelectRows(MissingKeyResourceKeys, each [Resource] <> null),
    CompleteKeyRows = Table.SelectRows(RowsWithoutErrors, each [Resource] <> null and Record.Field(_, periodColumn) <> null),
    KeyCounts = Table.Group(CompleteKeyRows, {"Resource", periodColumn}, {{"Rows", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.SelectRows(KeyCounts, each [Rows] > 1),
    DuplicateKeyCount = Table.RowCount(DuplicateKeys),
    AffectedResourceRows = Table.Combine({ErrorResourceRows, MissingKeyResourceRows, Table.SelectColumns(DuplicateKeys, {"Resource"})}),
    AffectedResources = List.Buffer(List.Distinct(AffectedResourceRows[Resource])),
    HasGlobalIssue = InputAttempt[HasError]
        or Table.RowCount(ErrorResourceRows) < ErrorCount
        or Table.RowCount(MissingKeyResourceRows) < MissingKeyCount,
    UsableRows = if HasGlobalIssue then EmptyInput else Table.SelectRows(CompleteKeyRows, each not List.Contains(AffectedResources, [Resource])),
    IssueCount = ErrorCount + MissingKeyCount + DuplicateKeyCount,
    Status = if InputAttempt[HasError] or ErrorCount > 0 then "Error" else if IssueCount > 0 then "Fail" else "Pass",
    Details = if InputAttempt[HasError] then (try InputAttempt[Error][Message] otherwise "Optional input evaluation failed")
        else if IssueCount > 0 then Number.ToText(ErrorCount) & " error row(s), " & Number.ToText(MissingKeyCount)
            & " missing-key row(s), " & Number.ToText(DuplicateKeyCount) & " duplicate key group(s)."
        else null
in
    [Data = Table.Buffer(UsableRows), Status = Status, Failures = if InputAttempt[HasError] then null else IssueCount,
     Details = Details, AffectedResources = AffectedResources, HasGlobalIssue = HasGlobalIssue];

// Query: A2CapPriorityInput_Prepare
// Purpose: Preserve the existing priority types while retaining fractional Resource and Period reduction budgets.
// Notes: Demand surplus below one cannot be rounded up to authorize removal of a whole optional shift.
shared A2CapPriorityInput_Prepare = let
    Source = #"IMPORT ResPeriodCapPrioritised",
    TypedPriorities = Table.TransformColumnTypes(Source, {{"Resource", Int64.Type}, {"Period", Int64.Type},
        {"Allocation", type any}, {"PotentialAvailability", type any}, {"PRCell", type text}, {"UnassignedAvail", Int64.Type},
        {"ResAv-C'", type number}, {"PeriodA-CNeg", type number}, {"C'-D", type number}, {"C'/D", type number},
        {"ClusterAllocation", type number}, {"NWDPriority", Int64.Type}, {"NWDType", type text}})
in
    TypedPriorities;

// Query: A2CapPriorityInputAssessment
// Purpose: Validate the saved A.1 priority schema and Resource/Period keys before sorting or reductions.
// Notes: ClusterAllocation may be null for days outside allocated clusters; key issues exclude only identifiable affected Resources.
shared A2CapPriorityInputAssessment = A2OptionalInputAssessment(() => A2CapPriorityInput_Prepare,
    {"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "ClusterAllocation", "NWDPriority", "NWDType"}, "Period");


// Query: PeriodCapPrioritisedTABLE
// Purpose: Order the existing period reduction priorities with Resource as the final tie-break.
shared PeriodCapPrioritisedTABLE = let
    Source = A2CapPriorityInputAssessment[Data],
    #"Sorted NWD,C-D,PA-Cn,PERIOD" = Table.Sort(Source,{{"Period", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"ClusterAllocation", Order.Descending}, {"C'-D", Order.Ascending}, {"PeriodA-CNeg", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted NWD,C-D,PA-Cn,PERIOD", "PeriodIndex",  2, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "PeriodIndexX", 1, 1, Int64.Type)
in
    #"Added Index1";

// Query: PeriodCapIndexLimits
// Purpose: Preserve period priority-index limits used to derive each candidate's ordinal within its period.
shared PeriodCapIndexLimits = let
    Source = PeriodCapPrioritisedTABLE,
    #"Grouped INDEX+COUNT" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([PeriodIndexX]), type number}, {"PeriodAvailNumber", each Table.RowCount(Table.Distinct(_)), Int64.Type}, {"PeriodResAv-C'", each List.Max([#"ResAv-C'"]), type nullable number}})
in
    #"Grouped INDEX+COUNT";

// Query: A2WorkdayInputAssessment
// Purpose: Assess the legacy roster-day statuses used by the existing C# calculation.
shared A2WorkdayInputAssessment = A2OptionalInputAssessment(() => #"IMPORT ResPeriodWDTABLE", {"Resource", "Period", "RosteredPeriodStatus"}, "Period");

// Query: A2ShiftChoiceInputAssessment
// Purpose: Assess the optional multiple-shift rejection interface without choosing a replacement shift.
shared A2ShiftChoiceInputAssessment = A2OptionalInputAssessment(() => #"IMPORT MultiPeriod - Remove", {"Resource", "AvailablePeriod", "NoAllocationKeep"}, "AvailablePeriod");

// Query: A2 Availability Publication Prepare
// Purpose: Preserve the mandatory legacy availability columns while taking Stage1 guards from the complete Resource-period skeleton.
shared #"A2 Availability Publication Prepare" = let
    Source = #"IMPORT ResPeriodAvailabilityTABLE",
    LegacyAvailabilityColumns = Table.SelectColumns(Source, {"Role", "Resource", "Period", "Availability"})
in
    LegacyAvailabilityColumns;

// Query: A2EligibilityMetadata_Prepare
// Purpose: Validate the authoritative A.1 spare-day plan at Resource/Period grain without inventing eligibility.
// Output: The existing skeleton with normalized eligibility fields and an explicit metadata status.
// Notes: Missing or invalid fields are reported as Error; normalization retains no implied permission.
shared A2EligibilityMetadata_Prepare = let
    Source = #"IMPORT ResourcePeriodTABLE_empty",
    // Preserve the legacy identifier types in staging, while validating the saved Day before any numeric coercion.
    TypedCoreIdentifiers = Table.TransformColumnTypes(Source, {{"Name", type text}, {"Role", type text},
        {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Shift", type any}}),
    EligibilityColumns = {"Day", "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday"},
    MissingColumns = List.Difference(EligibilityColumns, Table.ColumnNames(TypedCoreIdentifiers)),
    WithMissingColumns = List.Accumulate(MissingColumns, TypedCoreIdentifiers, (state, column) => Table.AddColumn(state, column, each null)),
    WithMetadataStatus = Table.AddColumn(WithMissingColumns, "A2EligibilityMetadataStatus", each
        let
            ValuesAttempt = try [Day = [Day], Eligible = [SpareDayEligible], Selected = [SparePeriodSelected],
                Reason = [SpareEligibilityReason], Cluster = [PlannedCluster], Side = [ExtensionSide], Workday = [AllocationWorkday]],
            IsCompleteAttempt = try if ValuesAttempt[HasError] then false else
                let
                    Values = ValuesAttempt[Value],
                    IsClusterNumber = if not Value.Is(Values[Cluster], type number) then false else
                        not Number.IsNaN(Values[Cluster]) and Number.Abs(Values[Cluster]) <> #infinity
                            and Values[Cluster] > 0 and Values[Cluster] = Number.RoundDown(Values[Cluster]),
                    IsDayNumber = if not Value.Is(Values[Day], type number) then false else
                        not Number.IsNaN(Values[Day]) and Number.Abs(Values[Day]) <> #infinity
                            and Values[Day] > 0 and Values[Day] = Number.RoundDown(Values[Day]),
                    HasValidTypes = IsDayNumber
                        and Value.Is(Values[Eligible], type logical) and Value.Is(Values[Selected], type logical)
                        and Value.Is(Values[Workday], type logical)
                        and (Values[Reason] = null or Value.Is(Values[Reason], type text))
                        and (Values[Side] = null or Value.Is(Values[Side], type text))
                        and (Values[Cluster] = null or IsClusterNumber),
                    HasValidPlan = if not HasValidTypes then false else
                        (not Values[Selected] or Values[Eligible])
                            and (not Values[Workday] or (not Values[Eligible] and not Values[Selected]))
                            and (not (Values[Workday] or Values[Selected]) or IsClusterNumber),
                    SelectedAvailability = try [OriginalAvailability] otherwise null,
                    HasAvailableSelection = if not HasValidPlan then false else if not Values[Selected] then true
                        else if not Value.Is(SelectedAvailability, type number) then false
                        else not Number.IsNaN(SelectedAvailability) and Number.Abs(SelectedAvailability) <> #infinity and SelectedAvailability > 0
                in
                    HasValidPlan and HasAvailableSelection,
            IsComplete = if IsCompleteAttempt[HasError] then false else IsCompleteAttempt[Value]
        in if List.IsEmpty(MissingColumns) and IsComplete then "Available" else "Error", type text),
    // Cell errors remain diagnosed before nullable normalization makes downstream exclusions evaluable.
    NormalizedFields = Table.TransformColumns(WithMetadataStatus, {
        {"Day", each let Value = try _ otherwise null in if Value.Is(Value, type number) then Value else null, type nullable number},
        {"SpareDayEligible", each let Value = try _ otherwise null in if Value.Is(Value, type logical) then Value else null, type nullable logical},
        {"SparePeriodSelected", each let Value = try _ otherwise null in if Value.Is(Value, type logical) then Value else null, type nullable logical},
        {"AllocationWorkday", each let Value = try _ otherwise null in if Value.Is(Value, type logical) then Value else null, type nullable logical},
        {"PlannedCluster", each let Value = try _ otherwise null in if Value.Is(Value, type number) then Value else null, type nullable number},
        {"SpareEligibilityReason", each let Value = try _ otherwise null in if Value.Is(Value, type text) then Value else null, type nullable text},
        {"ExtensionSide", each let Value = try _ otherwise null in if Value.Is(Value, type text) then Value else null, type nullable text}
    })
in
    Table.Buffer(NormalizedFields);

// Query: A2EligibilityDay_CHECK
// Purpose: Validate consistent day-level planning and exactly one selected period on each eligible spare day.
// Output: One Resource/day metadata status and reason; failures exclude optional spare for the affected Resource.
shared A2EligibilityDay_CHECK = let
    Source = A2EligibilityMetadata_Prepare,
    ResourceDays = Table.Group(Source, {"Resource", "Day"}, {{"Rows", each _, type table}}),
    WithReason = Table.AddColumn(ResourceDays, "A2EligibilityDayReason", each
        let
            Rows = [Rows],
            SelectedPeriods = List.Count(List.Select(Rows[SparePeriodSelected], each _ = true)),
            HasAllocation = List.AnyTrue(List.Transform(Table.ToRecords(Rows), each (try _[HasAllocationEvidence] otherwise false) = true)),
            Reasons = List.RemoveNulls({
                if List.Contains(Rows[A2EligibilityMetadataStatus], "Error") then "Eligibility fields are missing, invalid or inconsistent with the selected cell" else null,
                if List.Count(List.Distinct(Rows[SpareDayEligible])) > 1 then "Spare-day eligibility differs between periods on the same day" else null,
                if List.Count(List.Distinct(Rows[AllocationWorkday])) > 1 then "Original allocation workday differs between periods on the same day" else null,
                if List.Count(List.Distinct(Rows[PlannedCluster])) > 1 then "Planned cluster differs between periods on the same day" else null,
                if SelectedPeriods > 1 then "More than one spare period is selected on the same day" else null,
                if List.Contains(Rows[SpareDayEligible], true) and SelectedPeriods <> 1 then "Eligible spare day does not have exactly one selected period" else null,
                if HasAllocation and List.Contains(Rows[SpareDayEligible], true) then "Spare day has positive allocation evidence" else null
            })
        in if List.IsEmpty(Reasons) then null else Text.Combine(Reasons, "; "), type nullable text),
    WithStatus = Table.AddColumn(WithReason, "A2EligibilityDayStatus", each if [A2EligibilityDayReason] = null then "Available" else "Error", type text),
    Output = Table.RemoveColumns(WithStatus, {"Rows"})
in
    Table.Buffer(Output);

// Query: A2ClusterSchedule
// Purpose: Count worked days per planned cluster and calculate chronological full-day gaps within each Resource.
// Inputs: One row per Resource/day with Workday, AllocationWorkday, OptionalSpareWorkday and PlanAvailable.
// Notes: Existing internal one-day gaps remain within their planned cluster. WorkedDays counts days, never elapsed span.
shared A2ClusterSchedule = (days as table) as table =>
let
    WorkedDays = Table.SelectRows(days, each [Workday] = true and [PlanAvailable] = true and [PlannedCluster] <> null and [Day] <> null),
    ClusterTotals = Table.Group(WorkedDays, {"Resource", "Role", "PlannedCluster"}, {
        {"StartDay", each List.Min([Day]), type number},
        {"EndDay", each List.Max([Day]), type number},
        {"WorkedDays", each Table.RowCount(_), Int64.Type},
        {"OriginalWorkedDays", each List.Count(List.Select([AllocationWorkday], each _ = true)), Int64.Type},
        {"OptionalSpareDays", each List.Count(List.Select([OptionalSpareWorkday], each _ = true)), Int64.Type}
    }),
    ResourceGroups = Table.Group(ClusterTotals, {"Resource", "Role"}, {{"Clusters", each
        let
            OrderedClusters = Table.Sort(_, {{"StartDay", Order.Ascending}, {"EndDay", Order.Ascending}, {"PlannedCluster", Order.Ascending}}),
            NumberedClusters = Table.AddIndexColumn(OrderedClusters, "ClusterOrder", 0, 1, Int64.Type),
            ClusterIds = List.Buffer(NumberedClusters[PlannedCluster]),
            EndDays = List.Buffer(NumberedClusters[EndDay]),
            SpareDays = List.Buffer(NumberedClusters[OptionalSpareDays]),
            WithPreviousCluster = Table.AddColumn(NumberedClusters, "PreviousCluster", each
                if [ClusterOrder] = 0 then null else ClusterIds{[ClusterOrder] - 1}, type nullable number),
            WithPreviousEnd = Table.AddColumn(WithPreviousCluster, "PreviousEndDay", each
                if [ClusterOrder] = 0 then null else EndDays{[ClusterOrder] - 1}, type nullable number),
            WithPreviousSpare = Table.AddColumn(WithPreviousEnd, "PreviousOptionalSpareDays", each
                if [ClusterOrder] = 0 then 0 else SpareDays{[ClusterOrder] - 1}, Int64.Type),
            WithOffDays = Table.AddColumn(WithPreviousSpare, "OffDays", each
                if [PreviousEndDay] = null then null else [StartDay] - [PreviousEndDay] - 1, type nullable number)
        in
            WithOffDays, type table}}),
    ExpandedClusters = Table.ExpandTableColumn(ResourceGroups, "Clusters",
        {"PlannedCluster", "StartDay", "EndDay", "WorkedDays", "OriginalWorkedDays", "OptionalSpareDays", "PreviousCluster", "PreviousEndDay", "PreviousOptionalSpareDays", "OffDays"},
        {"PlannedCluster", "StartDay", "EndDay", "WorkedDays", "OriginalWorkedDays", "OptionalSpareDays", "PreviousCluster", "PreviousEndDay", "PreviousOptionalSpareDays", "OffDays"}),
    SortedClusters = Table.Sort(ExpandedClusters, {{"Resource", Order.Ascending}, {"Role", Order.Ascending}, {"StartDay", Order.Ascending}})
in
    Table.Buffer(SortedClusters);

// Query: A2PlannedDaySchedule_Prepare
// Purpose: Collapse the saved A.1 eligibility plan to one day before any C# calculation or optional reduction.
// Notes: Plan validation is independent of C# publication and cannot create a diagnostic dependency cycle.
shared A2PlannedDaySchedule_Prepare = let
    Source = A2EligibilityMetadata_Prepare,
    ResourceDays = Table.Group(Source, {"Resource", "Role", "Day"}, {
        {"AllocationWorkday", each List.Contains([AllocationWorkday], true), type logical},
        {"OptionalSpareWorkday", each List.AnyTrue(List.Transform(Table.ToRecords(_), each _[SpareDayEligible] = true and _[SparePeriodSelected] = true)), type logical},
        {"PlannedCluster", each let Values = List.Distinct([PlannedCluster]) in if List.Count(Values) = 1 then Values{0} else null, type nullable number},
        {"MetadataAvailable", each not List.Contains([A2EligibilityMetadataStatus], "Error"), type logical}
    }),
    JoinedDayStatus = Table.NestedJoin(ResourceDays, {"Resource", "Day"}, A2EligibilityDay_CHECK,
        {"Resource", "Day"}, "DayStatus", JoinKind.LeftOuter),
    ExpandedDayStatus = Table.ExpandTableColumn(JoinedDayStatus, "DayStatus", {"A2EligibilityDayStatus"}, {"A2EligibilityDayStatus"}),
    WithPlanAvailability = Table.AddColumn(ExpandedDayStatus, "PlanAvailable", each [MetadataAvailable] and [A2EligibilityDayStatus] = "Available", type logical),
    WithWorkday = Table.AddColumn(WithPlanAvailability, "Workday", each [AllocationWorkday] or [OptionalSpareWorkday], type logical)
in
    Table.Buffer(WithWorkday);

// Query: A2PlannedClusterSchedule_Prepare
// Purpose: Assess the imported eligibility calendar independently before C# publication.
shared A2PlannedClusterSchedule_Prepare = A2ClusterSchedule(A2PlannedDaySchedule_Prepare);

// Query: A2PlannedCalendarIssues_Prepare
// Purpose: Exclude Resource spare when a saved eligibility plan exceeds five worked days or leaves fewer than two full off days.
// Notes: Existing allocation breaches are reported and retained; these checks do not alter work flags or cap arithmetic.
shared A2PlannedCalendarIssues_Prepare = let
    Source = A2PlannedClusterSchedule_Prepare,
    OversizedClusters = Table.SelectRows(Source, each [WorkedDays] > 5),
    ShortRestGaps = Table.SelectRows(Source, each [OffDays] <> null and [OffDays] < 2),
    OversizedReasons = Table.AddColumn(OversizedClusters, "Reason", each "Saved eligibility cluster has " & Number.ToText([WorkedDays]) & " worked days; maximum is 5", type text),
    GapReasons = Table.AddColumn(ShortRestGaps, "Reason", each "Saved eligibility clusters have " & Number.ToText([OffDays]) & " full off days between them; minimum is 2", type text),
    IssueRows = Table.Combine({OversizedReasons, GapReasons}),
    Output = Table.SelectColumns(IssueRows, {"Resource", "Role", "PlannedCluster", "StartDay", "EndDay", "WorkedDays", "OffDays", "Reason"})
in
    Table.Buffer(Output);

// Query: A2PreReductionResourcePeriods_Prepare
// Purpose: Apply upstream eligibility and independent input exclusions before the joint reduction calculation.
// Output: Existing skeleton columns plus audit availability, spare permission, allocation evidence and validated spare eligibility.
// Notes: Missing Stage1 or eligibility metadata is Error, never Pass. Allocation evidence remains nullable until the legacy status fallback is available.
shared A2PreReductionResourcePeriods_Prepare = let
    Source = A2EligibilityMetadata_Prepare,
    Stage1Columns = {"OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence"},
    MissingColumns = List.Difference(Stage1Columns, Table.ColumnNames(Source)),
    WithMissingColumns = List.Accumulate(MissingColumns, Source, (state, column) => Table.AddColumn(state, column, each null)),
    // Reduction evaluation is assessed afterward, so candidate preparation cannot depend on its own result.
    Assessments = {A2CapPriorityInputAssessment, A2WorkdayInputAssessment, A2ShiftChoiceInputAssessment},
    HasGlobalOptionalIssue = List.AnyTrue(List.Transform(Assessments, each [HasGlobalIssue])),
    OptionalIssueResources = List.Buffer(List.Distinct(List.Combine(List.Transform(Assessments, each [AffectedResources])))),
    WithMetadataStatus = Table.AddColumn(WithMissingColumns, "A2Stage1MetadataStatus", each
        let
            AvailabilityAttempt = try [OriginalAvailability],
            PermissionAttempt = try [SpareCalculationAllowed],
            ReasonAttempt = try [SpareIssueReason],
            EvidenceAttempt = try [HasAllocationEvidence],
            HasValidValues = not AvailabilityAttempt[HasError] and not PermissionAttempt[HasError]
                and not ReasonAttempt[HasError] and not EvidenceAttempt[HasError],
            IsComplete = if not HasValidValues then false else
                (AvailabilityAttempt[Value] = null or Value.Is(AvailabilityAttempt[Value], type number))
                    and Value.Is(PermissionAttempt[Value], type logical)
                    and Value.Is(EvidenceAttempt[Value], type logical)
                    and (ReasonAttempt[Value] = null or Value.Is(ReasonAttempt[Value], type text))
        in if List.IsEmpty(MissingColumns) and IsComplete then "Available" else "Error", type text),
    MetadataIssueRows = Table.SelectRows(WithMetadataStatus, each [A2Stage1MetadataStatus] <> "Available"),
    MetadataIssueResources = List.Buffer(List.Distinct(MetadataIssueRows[Resource])),
    EligibilityIssueDays = Table.SelectRows(A2EligibilityDay_CHECK, each [A2EligibilityDayStatus] <> "Available"),
    EligibilityIssueResources = List.Buffer(List.Distinct(EligibilityIssueDays[Resource])),
    CalendarIssueResources = List.Buffer(List.Distinct(A2PlannedCalendarIssues_Prepare[Resource])),
    JoinedEligibilityDay = Table.NestedJoin(WithMetadataStatus, {"Resource", "Day"}, A2EligibilityDay_CHECK,
        {"Resource", "Day"}, "EligibilityDay", JoinKind.LeftOuter),
    ExpandedEligibilityDay = Table.ExpandTableColumn(JoinedEligibilityDay, "EligibilityDay",
        {"A2EligibilityDayStatus", "A2EligibilityDayReason"}, {"A2EligibilityDayStatus", "A2EligibilityDayReason"}),
    UpstreamExcludedRows = Table.SelectRows(WithMetadataStatus, each (try [SpareCalculationAllowed] otherwise null) = false),
    UpstreamExcludedResources = List.Buffer(List.Distinct(UpstreamExcludedRows[Resource])),
    ExcludedResources = List.Buffer(List.Distinct(List.Combine({MetadataIssueResources, EligibilityIssueResources, CalendarIssueResources, UpstreamExcludedResources, OptionalIssueResources}))),
    WithSparePermission = Table.AddColumn(ExpandedEligibilityDay, "A2SpareCalculationAllowed", each
        [A2Stage1MetadataStatus] = "Available"
            and [A2EligibilityMetadataStatus] = "Available" and [A2EligibilityDayStatus] = "Available"
            and (try [SpareCalculationAllowed] otherwise false) = true
            and not HasGlobalOptionalIssue
            and not List.Contains(ExcludedResources, [Resource]), type logical),
    WithIssueReason = Table.AddColumn(WithSparePermission, "A2SpareIssueReason", each
        let
            UpstreamReason = try [SpareIssueReason] otherwise null,
            Reasons = List.RemoveNulls({
                if Value.Is(UpstreamReason, type text) and Text.Trim(UpstreamReason) <> "" then UpstreamReason else null,
                if List.Contains(MetadataIssueResources, [Resource]) then "upstream Stage1 diagnostics unavailable for this Resource" else null,
                if List.Contains(EligibilityIssueResources, [Resource]) then "upstream spare eligibility is unavailable or inconsistent for this Resource" else null,
                if List.Contains(CalendarIssueResources, [Resource]) then "upstream eligibility plan breaks the worked-day or full-off-day rule for this Resource" else null,
                if List.Contains(UpstreamExcludedResources, [Resource]) and UpstreamReason = null then "upstream Stage1 excludes optional spare for this Resource" else null,
                if HasGlobalOptionalIssue then "A.2 optional input issue cannot be assigned to a Resource" else null,
                if List.Contains(OptionalIssueResources, [Resource]) then "A.2 optional input issue affects this Resource" else null
            })
        in if List.IsEmpty(Reasons) then null else Text.Combine(List.Distinct(Reasons), "; "), type nullable text),
    NormalizedAuditValues = Table.TransformColumns(WithIssueReason, {
        {"HasAllocationEvidence", each let Evidence = try _ otherwise null in if Value.Is(Evidence, type logical) then Evidence else null, type nullable logical},
        {"OriginalAvailability", each let Availability = try _ otherwise null in if Value.Is(Availability, type number) then Availability else null, type nullable number}
    }),
    RemovedInputGuards = Table.RemoveColumns(NormalizedAuditValues, {"SpareCalculationAllowed", "SpareIssueReason"}),
    RenamedGuards = Table.RenameColumns(RemovedInputGuards, {{"A2SpareCalculationAllowed", "SpareCalculationAllowed"}, {"A2SpareIssueReason", "SpareIssueReason"}})
in
    Table.Buffer(RenamedGuards);

// Query: A2CapacityBeforeReduction_Prepare
// Purpose: Pair the complete guarded skeleton with current availability before any cap reduction.
// Notes: Candidate preparation uses independent guards; allocated cells do not become optional reduction candidates.
shared A2CapacityBeforeReduction_Prepare = let
    Source = A2PreReductionResourcePeriods_Prepare,
    JoinedAvailability = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"A2 Availability Publication Prepare",
        {"Role", "Resource", "Period"}, "CurrentAvailability", JoinKind.LeftOuter),
    ExpandedAvailability = Table.ExpandTableColumn(JoinedAvailability, "CurrentAvailability", {"Availability"}, {"Availability"})
in
    Table.Buffer(ExpandedAvailability);

// Query: A2ReductionWorkdaySchedule_Prepare
// Purpose: Retain fixed original workdays and positive selected spare days for outermost-trim decisions.
// Output: One Resource/planned-cluster/day row with an immutable original flag and a stable day key.
shared A2ReductionWorkdaySchedule_Prepare = let
    Source = A2CapacityBeforeReduction_Prepare,
    // Already excluded Resources cannot create unusable trim keys or suppress other workers through the reducer.
    PermittedRows = Table.SelectRows(Source, each [SpareCalculationAllowed] = true),
    WithOptionalWorkday = Table.AddColumn(PermittedRows, "OptionalSpareWorkday", each
        let Amount = try [Availability] otherwise null in
        [SpareCalculationAllowed] = true and [SpareDayEligible] = true and [SparePeriodSelected] = true
            and [HasAllocationEvidence] = false and Value.Is(Amount, type number)
            and not Number.IsNaN(Amount) and Number.Abs(Amount) <> #infinity and Amount > 0, type logical),
    ResourceDays = Table.Group(WithOptionalWorkday, {"Resource", "PlannedCluster", "Day"}, {
        {"AllocationWorkday", each List.Contains([AllocationWorkday], true), type logical},
        {"OptionalSpareWorkday", each List.Contains([OptionalSpareWorkday], true), type logical}
    }),
    WorkedDays = Table.SelectRows(ResourceDays, each [AllocationWorkday] = true or [OptionalSpareWorkday] = true),
    WithDayKey = Table.AddColumn(WorkedDays, "DayKey", each "R" & Number.ToText([Resource], "0", "en-AU")
        & "-D" & Number.ToText([Day], "0", "en-AU"), type text),
    SortedDays = Table.Sort(WithDayKey, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}})
in
    Table.Buffer(SortedDays);

// Query: A2ReductionCandidates_Prepare
// Purpose: Join saved priorities to authoritative eligible, selected and evidence-free optional availability cells.
// Output: One Resource/Period candidate with complete day/plan keys and fractional-preserving budget values.
shared A2ReductionCandidates_Prepare = let
    Source = PeriodCapPrioritisedTABLE,
    PriorityColumns = Table.SelectColumns(Source, {"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'",
        "PeriodA-CNeg", "C'-D", "ClusterAllocation", "NWDPriority", "NWDType", "PeriodIndex", "PeriodIndexX"}),
    JoinedCapacity = Table.NestedJoin(PriorityColumns, {"Resource", "Period"}, A2CapacityBeforeReduction_Prepare,
        {"Resource", "Period"}, "Capacity", JoinKind.Inner),
    ExpandedCapacity = Table.ExpandTableColumn(JoinedCapacity, "Capacity",
        {"Role", "Day", "PlannedCluster", "AllocationWorkday", "HasAllocationEvidence", "SpareCalculationAllowed",
            "SpareDayEligible", "SparePeriodSelected", "Availability"},
        {"Role", "Day", "PlannedCluster", "AllocationWorkday", "HasAllocationEvidence", "SpareCalculationAllowed",
            "SpareDayEligible", "SparePeriodSelected", "Availability"}),
    OptionalCandidates = Table.SelectRows(ExpandedCapacity, each [SpareCalculationAllowed] = true
        and [SpareDayEligible] = true and [SparePeriodSelected] = true and [HasAllocationEvidence] = false),
    WithAmount = Table.AddColumn(OptionalCandidates, "CellAmount", each try [Availability] otherwise null, type nullable number),
    WithPeriodBudget = Table.AddColumn(WithAmount, "MaxPeriodReduction", each
        try List.Min({-[#"PeriodA-CNeg"], [#"C'-D"]}) otherwise null, type nullable number),
    WithDayKey = Table.AddColumn(WithPeriodBudget, "DayKey", each "R" & Number.ToText([Resource], "0", "en-AU")
        & "-D" & Number.ToText([Day], "0", "en-AU"), type text),
    WithCellKey = Table.AddColumn(WithDayKey, "CellKey", each "R" & Number.ToText([Resource], "0", "en-AU")
        & "-P" & Number.ToText([Period], "0", "en-AU"), type text),
    WithClusterKey = Table.AddColumn(WithCellKey, "ClusterKey", each "R" & Number.ToText([Resource], "0", "en-AU")
        & "-C" & Number.ToText([PlannedCluster], "0", "en-AU"), type text),
    WithBudgetReason = Table.AddColumn(WithClusterKey, "BudgetIssueReason", each
        let
            IsFinite = (value as any) as logical => Value.Is(value, type number)
                and not Number.IsNaN(value) and Number.Abs(value) <> #infinity,
            Reasons = List.RemoveNulls({
                if not IsFinite([CellAmount]) or [CellAmount] <= 0 then "Optional cell amount is not finite and positive" else null,
                if not IsFinite([#"ResAv-C'"]) or [#"ResAv-C'"] < 0 then "Resource reduction budget is not finite and nonnegative" else null,
                if not IsFinite([#"PeriodA-CNeg"]) or [#"PeriodA-CNeg"] > 0 then "Allocation-to-capacity surplus is unusable" else null,
                if not IsFinite([#"C'-D"]) or [#"C'-D"] < 0 then "Demand-surplus reduction budget is not finite and nonnegative" else null,
                if not IsFinite([MaxPeriodReduction]) or [MaxPeriodReduction] < 0 then "Joint period reduction budget is unusable" else null
            })
        in if List.IsEmpty(Reasons) then null else Text.Combine(Reasons, "; "), type nullable text)
in
    Table.Buffer(WithBudgetReason);

// Query: A2ReductionCandidateIssues_Prepare
// Purpose: Identify unusable cell amounts or inconsistent shared Resource/Period budgets without spending any budget.
// Notes: Every identifiable affected Resource is excluded from optional spare; a Period inconsistency affects all its candidate Resources.
shared A2ReductionCandidateIssues_Prepare = let
    Source = A2ReductionCandidates_Prepare,
    ResourceBudgets = Table.Group(Source, {"Resource"}, {{"BudgetValues", each List.Count(List.Distinct([#"ResAv-C'"])), Int64.Type}}),
    InconsistentResources = Table.SelectRows(ResourceBudgets, each [BudgetValues] <> 1)[Resource],
    PeriodBudgets = Table.Group(Source, {"Period"}, {{"BudgetValues", each List.Count(List.Distinct([MaxPeriodReduction])), Int64.Type}}),
    InconsistentPeriods = Table.SelectRows(PeriodBudgets, each [BudgetValues] <> 1)[Period],
    WithReason = Table.AddColumn(Source, "Reason", each
        let Reasons = List.RemoveNulls({[BudgetIssueReason],
            if List.Contains(InconsistentResources, [Resource]) then "Resource reduction budget differs between candidates" else null,
            if List.Contains(InconsistentPeriods, [Period]) then "Period reduction budget differs between candidates" else null})
        in if List.IsEmpty(Reasons) then null else Text.Combine(Reasons, "; "), type nullable text),
    IssueRows = Table.SelectRows(WithReason, each [Reason] <> null),
    Output = Table.SelectColumns(IssueRows, {"Resource", "Role", "Period", "Day", "PlannedCluster", "Reason"})
in
    Table.Buffer(Output);

[ Description = "BUFFER. Cap rs. availability based on roster" ]
// Query: Period Running Total
// Purpose: Expose initial period budgets and period priority positions for the joint reducer, without consuming prefix budgets.
// Notes: The legacy PeriodRunningTotal column now contains a priority ordinal; accepted budget use is AcceptedPeriodRunningTotal.
shared #"Period Running Total" = let
    Source = A2ReductionCandidates_Prepare,
    JoinedPeriodLimits = Table.NestedJoin(Source, {"Period"}, PeriodCapIndexLimits, {"Period"}, "PeriodLimits", JoinKind.LeftOuter),
    ExpandedPeriodLimits = Table.ExpandTableColumn(JoinedPeriodLimits, "PeriodLimits", {"StartPeriodIndex", "PeriodAvailNumber"}, {"StartPeriodIndex", "PeriodAvailNumber"}),
    WithPriorityPosition = Table.AddColumn(ExpandedPeriodLimits, "PeriodRunningTotal", each [PeriodIndexX] - [StartPeriodIndex] + 1, Int64.Type),
    WithWholeAmount = Table.AddColumn(WithPriorityPosition, "Reduction", each [CellAmount], type nullable number),
    WithPendingDecision = Table.AddColumn(WithWholeAmount, "ReducePeriod", each "Joint decision pending", type text),
    BUFFER = Table.Buffer(WithPendingDecision)
in
    BUFFER;

[ Description = "BUFFER" ]
// Query: ResCapPrioritisedTABLE
// Purpose: Order whole optional-cell candidates by the existing Resource and workday priorities with deterministic period ties.
shared ResCapPrioritisedTABLE = let
    Source = #"Period Running Total",
    #"Sorted RES,NWDPRI,MAXPRT,PRT" = Table.Sort(Source,{{"Resource", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"MaxPeriodReduction", Order.Descending}, {"PeriodRunningTotal", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES,NWDPRI,MAXPRT,PRT", "ResIndex", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

// Query: ResCapIndexLimits
// Purpose: Retain the legacy Resource priority-index summary for inspection.
shared ResCapIndexLimits = let
    Source = ResCapPrioritisedTABLE,
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "ResIndexX"}}),
    #"Grouped Rows" = Table.Group(#"Renamed Columns", {"Resource"}, {{"StartResIndex", each List.Min([ResIndexX]), type number}, {"ResUnassignedAvailable", each List.Max([#"ResAv-C'"]), type nullable number}, {"ResAvailNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: A2OuterSpareTrimAllowed
// Purpose: Permit removal only at a current planned cluster's outside edge, retaining all original workday anchors.
// Inputs: Day records for one Resource/plan, previously removed day keys, and the optional candidate day.
// Notes: This preserves extension prefixes/suffixes. A spare-only cluster can shrink from either end and disappear.
shared A2OuterSpareTrimAllowed = (clusterDays as list, removedDayKeys as list, candidateDay as number) as logical =>
let
    RemainingRecords = List.Select(clusterDays, each _[AllocationWorkday] = true or not List.Contains(removedDayKeys, _[DayKey])),
    RemainingDays = List.Transform(RemainingRecords, each _[Day]),
    CandidateRecords = List.Select(RemainingRecords, each _[Day] = candidateDay and _[AllocationWorkday] = false)
in
    not List.IsEmpty(CandidateRecords) and not List.IsEmpty(RemainingDays)
        and (candidateDay = List.Min(RemainingDays) or candidateDay = List.Max(RemainingDays));

// Query: A2SetBudgetSpent
// Purpose: Update one accepted-removal budget total in an immutable keyed state record.
shared A2SetBudgetSpent = (totals as record, key as text, amount as number) as record =>
let
    ExistingTotal = Record.FieldOrDefault(totals, key, 0),
    WithoutKey = Record.RemoveFields(totals, {key}, MissingField.Ignore),
    UpdatedTotal = Record.AddField(WithoutKey, key, ExistingTotal + amount)
in
    UpdatedTotal;

// Query: A2JointReductionPass
// Purpose: Scan candidates in business-priority order and spend Resource and Period budgets only on accepted outermost whole-cell removals.
// Notes: Rejected candidates do not change the state. A later pass can revisit an inner day exposed by an accepted outer trim.
shared A2JointReductionPass = (candidates as list, clusterDaysByKey as record, initialState as record) as record =>
    List.Accumulate(candidates, initialState, (state, candidate) =>
        let
            ResourceKey = Number.ToText(candidate[Resource], "0", "en-AU"),
            PeriodKey = Number.ToText(candidate[Period], "0", "en-AU"),
            ResourceSpent = Record.FieldOrDefault(state[ResourceSpent], ResourceKey, 0),
            PeriodSpent = Record.FieldOrDefault(state[PeriodSpent], PeriodKey, 0),
            AlreadyRemoved = List.Contains(state[RemovedDayKeys], candidate[DayKey]),
            ClusterDays = Record.Field(clusterDaysByKey, candidate[ClusterKey]),
            CanTrim = if AlreadyRemoved then false else A2OuterSpareTrimAllowed(ClusterDays, state[RemovedDayKeys], candidate[Day]),
            CanSpend = candidate[CellAmount] <= candidate[#"ResAv-C'"] - ResourceSpent
                and candidate[CellAmount] <= candidate[MaxPeriodReduction] - PeriodSpent,
            Accepted = not AlreadyRemoved and CanTrim and CanSpend,
            Decision = [CellKey = candidate[CellKey], ResourceAcceptedRemoval = ResourceSpent + candidate[CellAmount],
                PeriodAcceptedRemoval = PeriodSpent + candidate[CellAmount], AcceptanceOrder = state[AcceptedCount] + 1]
        in if not Accepted then state else
            [ResourceSpent = A2SetBudgetSpent(state[ResourceSpent], ResourceKey, candidate[CellAmount]),
             PeriodSpent = A2SetBudgetSpent(state[PeriodSpent], PeriodKey, candidate[CellAmount]),
             RemovedDayKeys = state[RemovedDayKeys] & {candidate[DayKey]},
             AcceptedByCell = Record.AddField(state[AcceptedByCell], candidate[CellKey], Decision),
             AcceptedCount = state[AcceptedCount] + 1]);

// Query: A2JointReductionDecisions
// Purpose: Complete accepted-only joint-budget reduction decisions, rescanning deferred inner candidates until no further trim is accepted.
// Output: One candidate decision with accepted totals, signed whole-cell reduction and an explicit reason for rejection.
shared A2JointReductionDecisions = let
    IssueResources = List.Buffer(List.Distinct(A2ReductionCandidateIssues_Prepare[Resource])),
    CandidateRows = Table.SelectRows(ResCapPrioritisedTABLE, each not List.Contains(IssueResources, [Resource])),
    CandidateRecords = List.Buffer(Table.ToRecords(CandidateRows)),
    DayGroups = Table.Group(A2ReductionWorkdaySchedule_Prepare, {"Resource", "PlannedCluster"}, {{"ClusterDays", each Table.ToRecords(_), type list}}),
    WithClusterKey = Table.AddColumn(DayGroups, "Name", each "R" & Number.ToText([Resource], "0", "en-AU")
        & "-C" & Number.ToText([PlannedCluster], "0", "en-AU"), type text),
    KeyedClusterDays = Table.RenameColumns(Table.SelectColumns(WithClusterKey, {"Name", "ClusterDays"}), {{"ClusterDays", "Value"}}),
    ClusterDaysByKey = Record.FromTable(KeyedClusterDays),
    InitialState = [ResourceSpent = [], PeriodSpent = [], RemovedDayKeys = {}, AcceptedByCell = [], AcceptedCount = 0],
    PassStates = List.Generate(() => [State = InitialState, Continue = true], each [Continue], each
        let NextState = A2JointReductionPass(CandidateRecords, ClusterDaysByKey, [State])
        in [State = NextState, Continue = NextState[AcceptedCount] > [State][AcceptedCount]], each [State]),
    FinalState = List.Last(PassStates),
    WithDecision = Table.AddColumn(CandidateRows, "JointDecision", each
        let
            AcceptedRecord = Record.FieldOrDefault(FinalState[AcceptedByCell], [CellKey], null),
            ResourceSpent = Record.FieldOrDefault(FinalState[ResourceSpent], Number.ToText([Resource], "0", "en-AU"), 0),
            PeriodSpent = Record.FieldOrDefault(FinalState[PeriodSpent], Number.ToText([Period], "0", "en-AU"), 0),
            CanTrim = A2OuterSpareTrimAllowed(Record.Field(ClusterDaysByKey, [ClusterKey]), FinalState[RemovedDayKeys], [Day]),
            Reason = if AcceptedRecord <> null then "Accepted whole optional cell"
                else if not CanTrim then "Deferred inner spare day remains protected by an outer worked day"
                else if [CellAmount] > [#"ResAv-C'"] - ResourceSpent then "Remaining Resource reduction budget is smaller than this whole optional cell"
                else if [CellAmount] > [MaxPeriodReduction] - PeriodSpent then "Remaining Period reduction budget is smaller than this whole optional cell"
                else "No additional removal was accepted"
        in [Accepted = AcceptedRecord <> null, Reduction = if AcceptedRecord = null then 0 else -[CellAmount],
            ResRunningTotal = if AcceptedRecord = null then ResourceSpent else AcceptedRecord[ResourceAcceptedRemoval],
            AcceptedPeriodRunningTotal = if AcceptedRecord = null then PeriodSpent else AcceptedRecord[PeriodAcceptedRemoval],
            AcceptanceOrder = if AcceptedRecord = null then null else AcceptedRecord[AcceptanceOrder], Reason = Reason], type record),
    ExpandedDecision = Table.ExpandRecordColumn(WithDecision, "JointDecision",
        {"Accepted", "Reduction", "ResRunningTotal", "AcceptedPeriodRunningTotal", "AcceptanceOrder", "Reason"},
        {"Accepted", "SignedReduction", "ResRunningTotal", "AcceptedPeriodRunningTotal", "AcceptanceOrder", "Reason"})
in
    Table.Buffer(ExpandedDecision);

[ Description = "BUFFER. Cap rs. availability based on roster" ]
// Query: Resource Running total
// Purpose: Publish accepted whole-cell reductions with Resource totals accumulated only from accepted joint decisions.
// Output: The existing Resource/Period/PRCell/signed Reduction/ResRunningTotal interface.
shared #"Resource Running total" = let
    Source = A2JointReductionDecisions,
    AcceptedDecisions = Table.SelectRows(Source, each [Accepted] = true),
    OrderedAcceptedDecisions = Table.Sort(AcceptedDecisions, {{"Resource", Order.Ascending}, {"AcceptanceOrder", Order.Ascending}}),
    SelectedColumns = Table.SelectColumns(OrderedAcceptedDecisions, {"Resource", "Period", "PRCell", "SignedReduction", "ResRunningTotal"}),
    RenamedReduction = Table.RenameColumns(SelectedColumns, {{"SignedReduction", "Reduction"}}),
    BUFFER = Table.Buffer(RenamedReduction)
in
    BUFFER;

// Query: CheckCap
// Purpose: Retain the legacy accepted Resource-reduction totals for inspection; final cap compliance is checked separately.
shared CheckCap = let
    Source = #"Resource Running total",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"MAXResRunningTotal", each List.Max([ResRunningTotal]), type number}, {"AvailabilityReduction", each List.Sum([Reduction]), type number}})
in
    #"Grouped Rows";

// Query: A2ReductionInputAssessment
// Purpose: Assess optional signed cap reductions without hiding a calculation or interface error.
shared A2ReductionInputAssessment = A2OptionalInputAssessment(() => #"Resource Running total", {"Resource", "Period", "Reduction"}, "Period");

// Query: A2JointBudgetIssues_Prepare
// Purpose: Reconcile accepted whole-cell removals against the shared Resource and Period budgets.
// Notes: This audit uses accepted rows only; rejected candidates cannot contribute to a budget total.
shared A2JointBudgetIssues_Prepare = let
    Source = Table.SelectRows(A2JointReductionDecisions, each [Accepted] = true),
    ResourceTotals = Table.Group(Source, {"Resource"}, {
        {"AcceptedRemoval", each List.Sum([CellAmount]), type number},
        {"Budget", each List.Min([#"ResAv-C'"]), type number}
    }),
    PeriodTotals = Table.Group(Source, {"Period"}, {
        {"AcceptedRemoval", each List.Sum([CellAmount]), type number},
        {"Budget", each List.Min([MaxPeriodReduction]), type number}
    }),
    ResourceFailures = Table.SelectRows(ResourceTotals, each [AcceptedRemoval] < 0 or [AcceptedRemoval] > [Budget]),
    PeriodFailures = Table.SelectRows(PeriodTotals, each [AcceptedRemoval] < 0 or [AcceptedRemoval] > [Budget]),
    ResourceReasons = Table.AddColumn(ResourceFailures, "Reason", each "Accepted Resource removal "
        & Number.ToText([AcceptedRemoval]) & " exceeds its budget " & Number.ToText([Budget]), type text),
    ResourceProof = Table.AddColumn(ResourceReasons, "Period", each null, type nullable number),
    PeriodReasons = Table.AddColumn(PeriodFailures, "Reason", each "Accepted Period removal "
        & Number.ToText([AcceptedRemoval]) & " exceeds its budget " & Number.ToText([Budget]), type text),
    PeriodProof = Table.AddColumn(PeriodReasons, "Resource", each null, type nullable number),
    Output = Table.Combine({Table.SelectColumns(ResourceProof, {"Resource", "Period", "AcceptedRemoval", "Budget", "Reason"}),
        Table.SelectColumns(PeriodProof, {"Resource", "Period", "AcceptedRemoval", "Budget", "Reason"})})
in
    Table.Buffer(Output);

// Query: A2 Resource Period Stage1 Prepare
// Purpose: Add diagnosed joint-reduction exclusions to the independently prepared Resource-period guards.
// Notes: A calculation error excludes optional spare without changing allocation evidence or introducing a dependency cycle.
shared #"A2 Resource Period Stage1 Prepare" = let
    Source = A2PreReductionResourcePeriods_Prepare,
    Assessment = A2ReductionInputAssessment,
    BudgetIssueResources = List.Buffer(List.Distinct(A2ReductionCandidateIssues_Prepare[Resource])),
    AffectedResources = List.Buffer(List.Distinct(List.Combine({Assessment[AffectedResources], BudgetIssueResources}))),
    WithPermission = Table.AddColumn(Source, "A2ReductionSpareAllowed", each [SpareCalculationAllowed] = true
        and not Assessment[HasGlobalIssue] and not List.Contains(AffectedResources, [Resource]), type logical),
    WithReason = Table.AddColumn(WithPermission, "A2ReductionIssueReason", each
        let Reasons = List.RemoveNulls({[SpareIssueReason],
            if Assessment[HasGlobalIssue] then "A.2 joint reduction evaluation is unavailable" else null,
            if List.Contains(Assessment[AffectedResources], [Resource]) then "A.2 joint reduction interface affects this Resource" else null,
            if List.Contains(BudgetIssueResources, [Resource]) then "A.2 reduction budgets or optional cell amounts are unusable for this Resource" else null})
        in if List.IsEmpty(Reasons) then null else Text.Combine(List.Distinct(Reasons), "; "), type nullable text),
    RemovedGuards = Table.RemoveColumns(WithReason, {"SpareCalculationAllowed", "SpareIssueReason"}),
    RenamedGuards = Table.RenameColumns(RemovedGuards, {{"A2ReductionSpareAllowed", "SpareCalculationAllowed"}, {"A2ReductionIssueReason", "SpareIssueReason"}})
in
    Table.Buffer(RenamedGuards);

// Query: A2 Selected Reductions Prepare
// Purpose: Use optional signed reductions only for Resources whose spare calculation remains allowed.
// Notes: Diagnosed input failure or Resource exclusion retains allocation evidence and cannot become an unreported empty passing input.
shared #"A2 Selected Reductions Prepare" = let
    Source = A2ReductionInputAssessment[Data],
    ResourcePermissions = Table.Group(#"A2 Resource Period Stage1 Prepare", {"Resource"}, {{"Allowed", each List.AllTrue([SpareCalculationAllowed]), type logical}}),
    AllowedResources = Table.SelectRows(ResourcePermissions, each [Allowed] = true),
    JoinedPermissions = Table.NestedJoin(Source, {"Resource"}, AllowedResources, {"Resource"}, "Permission", JoinKind.Inner),
    SelectedColumns = Table.SelectColumns(JoinedPermissions, {"Resource", "Period", "Reduction"})
in
    Table.Buffer(SelectedColumns);

// Query: A2 Workday Status Prepare
// Purpose: Return the assessed legacy statuses, including allocated-period evidence needed when upstream Stage1 metadata is absent.
shared #"A2 Workday Status Prepare" = A2WorkdayInputAssessment[Data];

// Query: A2 Shift Choices Prepare
// Purpose: Return the assessed existing shift choices without changing their selection rules.
shared #"A2 Shift Choices Prepare" = A2ShiftChoiceInputAssessment[Data];

// Query: ResPeriodCappedAvailability(C#)TABLE
// Purpose: Calculate C# from authoritative selected spare while retaining available cells with allocation evidence.
// Notes: Legacy workday/shift-choice labels cannot reject selected legal spare. Mandatory availability/skeleton errors remain errors.
shared #"ResPeriodCappedAvailability(C#)TABLE" = let
    Source = #"A2 Availability Publication Prepare",

    #"Merged PREMPTY" = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"A2 Resource Period Stage1 Prepare", {"Role", "Resource", "Period"}, "ResourcePeriodTABLE_empty", JoinKind.RightOuter),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged PREMPTY",{"Role", "Resource", "Period"}),
    #"Expanded ResourcePeriodTABLE_empty" = Table.ExpandTableColumn(#"Removed Columns1", "ResourcePeriodTABLE_empty",
        {"Role", "Resource", "Period", "Day", "OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence",
            "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday",
            "A2Stage1MetadataStatus", "A2EligibilityMetadataStatus", "A2EligibilityDayStatus"},
        {"Role", "Resource", "Period", "Day", "OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence",
            "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday",
            "A2Stage1MetadataStatus", "A2EligibilityMetadataStatus", "A2EligibilityDayStatus"}),

    #"Merged Queries" = Table.NestedJoin(#"Expanded ResourcePeriodTABLE_empty", {"Resource", "Period"}, #"A2 Selected Reductions Prepare", {"Resource", "Period"}, "ReCapAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReCapAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReCapAvailabilityTABLE", {"Reduction"}, {"Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ReCapAvailabilityTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),

    // Optional status evidence enriches the authoritative skeleton; it cannot introduce extra capacity cells.
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Resource", "Period"}, #"A2 Workday Status Prepare", {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodWDTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodWDTABLE", {"RosteredPeriodStatus"}, {"RosteredPeriodStatus"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ResPeriodWDTABLE",null,0,Replacer.ReplaceValue,{"Reduction"}),

    #"Merged Queries2" = Table.NestedJoin(#"Replaced Value", {"Resource", "Period"}, #"A2 Shift Choices Prepare", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-Remove" = Table.ExpandTableColumn(#"Merged Queries2", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    WithAllocationEvidence = Table.AddColumn(#"Expanded MultiDayPeriod-Remove", "A2HasAllocationEvidence", each
        if Value.Is([HasAllocationEvidence], type logical) then [HasAllocationEvidence]
        else List.Contains({"AllocatedPeriod", "UnavailableAllocatedPeriod"}, [RosteredPeriodStatus]), type logical),
    WithSparePermission = Table.AddColumn(WithAllocationEvidence, "A2SpareCalculationAllowed", each (try [SpareCalculationAllowed] otherwise false) = true, type logical),
    WithIssueReason = Table.AddColumn(WithSparePermission, "A2SpareIssueReason", each
        if [SpareCalculationAllowed] = null then "upstream Stage1 diagnostics unavailable" else [SpareIssueReason], type nullable text),
    RemovedInputGuards = Table.RemoveColumns(WithIssueReason, {"HasAllocationEvidence", "SpareCalculationAllowed", "SpareIssueReason"}),
    PreparedAllocationEvidence = Table.RenameColumns(RemovedInputGuards, {{"A2HasAllocationEvidence", "HasAllocationEvidence"}, {"A2SpareCalculationAllowed", "SpareCalculationAllowed"}, {"A2SpareIssueReason", "SpareIssueReason"}}),
    // Allocation evidence is retained independently. Optional spare needs both authoritative eligibility flags before any reduction.
    // Missing original availability still contributes zero; no allocation availability is fabricated.
    #"Inserted REDUCTION" = Table.AddColumn(PreparedAllocationEvidence, "AvailabilityCapped", each
        if [Availability] = null then 0
        else if [HasAllocationEvidence] = true then [Availability]
        else if [SpareCalculationAllowed] = true and [SpareDayEligible] = true and [SparePeriodSelected] = true
            then [Availability] + [Reduction]
        else 0),
        #"Removed Columns" = Table.RemoveColumns(#"Inserted REDUCTION",{"Availability", "Reduction", "NoAllocationKeep"})
    in
        #"Removed Columns";

[ Description = "BUFFER-All cells capped availability" ]
// Query: ResPeriodAvailabilityCapped(C#)TABLE
// Purpose: Publish C# with Resource/day eligibility, selected-period and allocation/audit evidence on the existing interface.
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    Source = #"ResPeriodCappedAvailability(C#)TABLE"
in
    Source;

shared #"ResPeriodAvailabilityCapped(C#)SUM" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    AvailabilityCapped = Source[AvailabilityCapped],
    #"Calculated Sum" = List.Sum(AvailabilityCapped)
in
    #"Calculated Sum";

// Query: ResPeriodAvailabilityCappedMATRIX
// Purpose: Preserve the legacy Resource/Role pivot grain while diagnostic metadata stays on the C# table.
shared ResPeriodAvailabilityCappedMATRIX = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Removed Columns" = Table.SelectColumns(Source,{"Role", "Resource", "Period", "AvailabilityCapped"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "AvailabilityCapped", List.Sum)
in
    #"Pivoted Column";

// Query: A2EligibilityResourceStatus_Prepare
// Purpose: Identify complete Resource calendars so partial or missing plan metadata cannot receive a passing calendar check.
shared A2EligibilityResourceStatus_Prepare = let
    Source = #"A2 Resource Period Stage1 Prepare",
    ResourceStatus = Table.Group(Source, {"Resource"}, {{"PlanAvailable", each
        List.AllTrue(List.Transform(Table.ToRecords(_), each _[A2Stage1MetadataStatus] = "Available"
            and _[A2EligibilityMetadataStatus] = "Available" and _[A2EligibilityDayStatus] = "Available")), type logical}})
in
    Table.Buffer(ResourceStatus);

// Query: A2PublishedDaySchedule_Prepare
// Purpose: Count actual C# workdays once per Resource/day while retaining original allocation workday flags.
// Notes: A positive optional cell contributes a workday only without allocation evidence. AllocationWorkday is the original daily flag.
shared A2PublishedDaySchedule_Prepare = let
    Source = #"ResPeriodCappedAvailability(C#)TABLE",
    ResourceDays = Table.Group(Source, {"Resource", "Role", "Day"}, {
        {"AllocationWorkday", each List.Contains([AllocationWorkday], true), type logical},
        {"OptionalSpareWorkday", each List.AnyTrue(List.Transform(Table.ToRecords(_), each
            _[HasAllocationEvidence] = false and _[AvailabilityCapped] <> null and _[AvailabilityCapped] > 0)), type logical},
        {"PlannedCluster", each let Values = List.Distinct([PlannedCluster]) in if List.Count(Values) = 1 then Values{0} else null, type nullable number}
    }),
    JoinedResourceStatus = Table.NestedJoin(ResourceDays, {"Resource"}, A2EligibilityResourceStatus_Prepare,
        {"Resource"}, "ResourceStatus", JoinKind.LeftOuter),
    ExpandedResourceStatus = Table.ExpandTableColumn(JoinedResourceStatus, "ResourceStatus", {"PlanAvailable"}, {"PlanAvailable"}),
    WithWorkday = Table.AddColumn(ExpandedResourceStatus, "Workday", each [AllocationWorkday] or [OptionalSpareWorkday], type logical)
in
    Table.Buffer(WithWorkday);

// Query: A2PublishedClusterSchedule_Prepare
// Purpose: Calculate worked-day totals and chronological gaps from actual C# plus original allocation flags.
shared A2PublishedClusterSchedule_Prepare = A2ClusterSchedule(A2PublishedDaySchedule_Prepare);

// Query: A2OriginalDaySchedule_Prepare
// Purpose: Isolate the original allocation workdays without counting optional capacity as allocation work.
shared A2OriginalDaySchedule_Prepare = let
    Source = A2PublishedDaySchedule_Prepare,
    OriginalFlags = Table.TransformColumns(Source, {{"OptionalSpareWorkday", each false, type logical}}),
    RemovedPublishedWorkday = Table.RemoveColumns(OriginalFlags, {"Workday"}),
    WithOriginalWorkday = Table.AddColumn(RemovedPublishedWorkday, "Workday", each [AllocationWorkday], type logical)
in
    Table.Buffer(WithOriginalWorkday);

// Query: A2OriginalClusterSchedule_Prepare
// Purpose: Expose the allocation-only cluster calendar for separate preexisting-breach diagnostics.
shared A2OriginalClusterSchedule_Prepare = A2ClusterSchedule(A2OriginalDaySchedule_Prepare);

// Query: A2CalendarRuleIssues
// Purpose: Return one issue per breached worked-day or full-off-day rule on an already prepared cluster calendar.
// Notes: This helper reports failures; it does not raise errors, choose spare days or change availability.
shared A2CalendarRuleIssues = (clusters as table) as table =>
let
    OversizedClusters = Table.SelectRows(clusters, each [WorkedDays] > 5),
    ShortRestGaps = Table.SelectRows(clusters, each [OffDays] <> null and [OffDays] < 2),
    ClusterChecks = Table.AddColumn(OversizedClusters, "Check", each "Maximum five worked days per planned cluster", type text),
    ClusterReasons = Table.AddColumn(ClusterChecks, "Reason", each "Planned cluster " & Number.ToText([PlannedCluster])
        & " has " & Number.ToText([WorkedDays]) & " worked days; maximum is 5", type text),
    GapChecks = Table.AddColumn(ShortRestGaps, "Check", each "Minimum two full off days between planned clusters", type text),
    GapReasons = Table.AddColumn(GapChecks, "Reason", each "Planned clusters " & Number.ToText([PreviousCluster])
        & " and " & Number.ToText([PlannedCluster]) & " have " & Number.ToText([OffDays])
        & " full off days between them; minimum is 2", type text),
    Issues = Table.Combine({ClusterReasons, GapReasons})
in
    Table.Buffer(Issues);

// Query: A2OriginalCalendarIssues_Prepare
// Purpose: Report preexisting allocation-only breaches separately from any optional-spare publication failure.
shared A2OriginalCalendarIssues_Prepare = let
    Source = A2CalendarRuleIssues(A2OriginalClusterSchedule_Prepare),
    WithOrigin = Table.AddColumn(Source, "Origin", each "Preexisting allocation", type text)
in
    Table.Buffer(WithOrigin);

// Query: A2PublishedCalendarIssues_Prepare
// Purpose: Report actual C# calendar breaches and identify whether the same breach exists in the allocation-only calendar.
// Notes: Plan IDs may interleave chronologically; gap comparison uses adjacent dates and both actual plan IDs.
shared A2PublishedCalendarIssues_Prepare = let
    Source = A2CalendarRuleIssues(A2PublishedClusterSchedule_Prepare),
    JoinedOriginalCalendar = Table.NestedJoin(Source, {"Resource", "Role", "PreviousCluster", "PlannedCluster"},
        A2OriginalClusterSchedule_Prepare, {"Resource", "Role", "PreviousCluster", "PlannedCluster"}, "OriginalCalendar", JoinKind.LeftOuter),
    ExpandedOriginalCalendar = Table.ExpandTableColumn(JoinedOriginalCalendar, "OriginalCalendar", {"OffDays"}, {"OriginalOffDays"}),
    WithOrigin = Table.AddColumn(ExpandedOriginalCalendar, "Origin", each
        if [Check] = "Maximum five worked days per planned cluster" and [OriginalWorkedDays] > 5 then "Preexisting allocation"
        else if [Check] = "Minimum two full off days between planned clusters" and [OriginalOffDays] <> null and [OriginalOffDays] < 2 then "Preexisting allocation"
        else "Published optional spare", type text)
in
    Table.Buffer(WithOrigin);

// Query: A2SparePublicationIssues_Prepare
// Purpose: Expose any positive optional C# cell retained despite its Resource exclusion or eligibility flags.
// Notes: This is an audit of the actual output, never an additional publication gate.
shared A2SparePublicationIssues_Prepare = let
    Source = #"ResPeriodCappedAvailability(C#)TABLE",
    PositiveOptionalCells = Table.SelectRows(Source, each [HasAllocationEvidence] = false
        and [AvailabilityCapped] <> null and [AvailabilityCapped] > 0),
    ExcludedCells = Table.SelectRows(PositiveOptionalCells, each [SpareCalculationAllowed] <> true),
    IneligibleCells = Table.SelectRows(PositiveOptionalCells, each not ([SpareDayEligible] = true and [SparePeriodSelected] = true)),
    ExclusionChecks = Table.AddColumn(ExcludedCells, "Check", each "Excluded optional spare is not retained in C#", type text),
    ExclusionReasons = Table.AddColumn(ExclusionChecks, "Reason", each "Positive optional C# remains on a Resource excluded from spare calculation", type text),
    EligibilityChecks = Table.AddColumn(IneligibleCells, "Check", each "Unselected or ineligible optional spare is not retained in C#", type text),
    EligibilityReasons = Table.AddColumn(EligibilityChecks, "Reason", each "Positive optional C# remains without both authoritative eligibility and selected-period flags", type text),
    IssueRows = Table.Combine({ExclusionReasons, EligibilityReasons}),
    Output = Table.SelectColumns(IssueRows, {"Resource", "Role", "Period", "Day", "PlannedCluster", "Check", "Reason"})
in
    Table.Buffer(Output);

// Query: A2WholeOptionalCellIssues_Prepare
// Purpose: Audit that actual optional C# is either its complete original current-cell amount or zero.
// Notes: Allocated cells are excluded from this whole-optional-cell test and remain independently preserved.
shared A2WholeOptionalCellIssues_Prepare = let
    Source = Table.SelectRows(#"ResPeriodCappedAvailability(C#)TABLE", each [HasAllocationEvidence] = false),
    JoinedOriginalCells = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"A2 Availability Publication Prepare",
        {"Role", "Resource", "Period"}, "OriginalCells", JoinKind.LeftOuter),
    ExpandedOriginalCells = Table.ExpandTableColumn(JoinedOriginalCells, "OriginalCells", {"Availability"}, {"CurrentCellAmount"}),
    SplitCells = Table.SelectRows(ExpandedOriginalCells, each [AvailabilityCapped] <> 0 and [AvailabilityCapped] <> [CurrentCellAmount]),
    WithReason = Table.AddColumn(SplitCells, "Reason", each "Optional C# is neither zero nor its whole current-cell availability amount", type text),
    Output = Table.SelectColumns(WithReason, {"Resource", "Role", "Period", "Day", "PlannedCluster", "AvailabilityCapped", "CurrentCellAmount", "Reason"})
in
    Table.Buffer(Output);




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

shared RolePath = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    Folder = #"Renamed Columns"{0}[Folder]
in
    Folder;

// Query: IMPORT ResPeriodCapPrioritised
// Purpose: Read the saved A.1 reduction-priority table unchanged for downstream staging.
shared #"IMPORT ResPeriodCapPrioritised" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodCapPrioritised_Table = Source{[Item="ResPeriodCapPrioritised",Kind="Table"]}[Data]
in
    ResPeriodCapPrioritised_Table;

shared #"IMPORT MultiPeriod - Remove" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    MultiDayPeriod_Remove_Table = Source{[Item="MultiDayPeriod_Remove",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(MultiDayPeriod_Remove_Table,{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type2",{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}})
in
    #"Changed Type1";

shared #"IMPORT ResPeriodWDTABLE" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodWDTABLE_Table = Source{[Item="ResPeriodWDTABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodWDTABLE_Table,{{"Resource", Int64.Type}, {"Day", Int64.Type}, {"PotentialAvailability", type text}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type1";

shared #"IMPORT ResPeriodAvailabilityTABLE" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORT ResourcePeriodTABLE_empty
// Purpose: Read the saved A.1 complete Resource-period table unchanged for downstream metadata validation.
// Notes: Day remains raw so staging can reject invalid, fractional or coerced-looking values rather than rounding them first.
shared #"IMPORT ResourcePeriodTABLE_empty" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResourcePeriodTABLE_empty_Table = Source{[Item="ResourcePeriodTABLE_empty",Kind="Table"]}[Data]
in
    ResourcePeriodTABLE_empty_Table;

// Query: A2DiagnosticCount
// Purpose: Return an honest diagnostic count without raising a business-check publication gate.
shared A2DiagnosticCount = (check as text, evaluate as function) as record =>
let
    Attempt = try evaluate()
in
    [Stage = "A.2", Check = check,
     Status = if Attempt[HasError] then "Error" else if Attempt[Value] = 0 then "Pass" else "Fail",
     Failures = if Attempt[HasError] then null else Attempt[Value],
     Details = if Attempt[HasError] then (try Attempt[Error][Message] otherwise "Diagnostic evaluation failed") else null];

// Query: A2CalendarDiagnosticCount
// Purpose: Report known calendar violations without treating an incomplete Resource plan as a passing calendar.
// Notes: A known breach is Fail; otherwise any unevaluated Resource makes the check NotEvaluated.
shared A2CalendarDiagnosticCount = (check as text, evaluate as function) as record =>
let
    FailureAttempt = try evaluate(),
    UnknownAttempt = try Table.RowCount(Table.SelectRows(A2EligibilityResourceStatus_Prepare, each [PlanAvailable] <> true)),
    HasError = FailureAttempt[HasError] or UnknownAttempt[HasError],
    ErrorRecord = if FailureAttempt[HasError] then FailureAttempt[Error] else if UnknownAttempt[HasError] then UnknownAttempt[Error] else null
in
    [Stage = "A.2", Check = check,
     Status = if HasError then "Error" else if FailureAttempt[Value] > 0 then "Fail"
        else if UnknownAttempt[Value] > 0 then "NotEvaluated" else "Pass",
     Failures = if HasError then null else FailureAttempt[Value],
     Details = if HasError then (try ErrorRecord[Message] otherwise "Calendar diagnostic evaluation failed")
        else if UnknownAttempt[Value] > 0 then Number.ToText(UnknownAttempt[Value])
            & " Resource calendar(s) were not evaluated because their eligibility or allocation-evidence metadata is incomplete."
        else null];

// Query: CapacityDiagnostics_SUMMARY
// Purpose: Publish nonblocking input, eligibility and actual calendar checks for manual review.
// Output: One row per check; missing metadata is Error and incomplete calendars cannot receive Pass.
// Notes: Legacy CheckCap remains available; its totals are not treated as proof of final contract-cap compliance.
shared CapacityDiagnostics_SUMMARY = let
    OptionalInputs = {
        [Check = "Upstream cap priority interface", Assessment = A2CapPriorityInputAssessment],
        [Check = "Optional cap reduction input", Assessment = A2ReductionInputAssessment],
        [Check = "Legacy workday status input", Assessment = A2WorkdayInputAssessment],
        [Check = "Multiple-shift choice input", Assessment = A2ShiftChoiceInputAssessment]
    },
    OptionalChecks = List.Transform(OptionalInputs, each
        let Assessment = [Assessment] in
        [Stage = "A.2", Check = [Check], Status = Assessment[Status], Failures = Assessment[Failures], Details = Assessment[Details]]),
    MetadataSchemaAttempt = try List.Difference({"OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence"}, Table.ColumnNames(#"IMPORT ResourcePeriodTABLE_empty")),
    MetadataAttempt = try Table.RowCount(Table.SelectRows(#"A2 Resource Period Stage1 Prepare", each [A2Stage1MetadataStatus] <> "Available")),
    MetadataCheck = [Stage = "A.2", Check = "Upstream Stage1 diagnostic metadata",
        Status = if MetadataSchemaAttempt[HasError] or MetadataAttempt[HasError] then "Error"
            else if not List.IsEmpty(MetadataSchemaAttempt[Value]) or MetadataAttempt[Value] > 0 then "Error" else "Pass",
        Failures = if MetadataSchemaAttempt[HasError] or MetadataAttempt[HasError] then null else MetadataAttempt[Value],
        Details = if MetadataSchemaAttempt[HasError] then (try MetadataSchemaAttempt[Error][Message] otherwise "Mandatory Resource-period schema evaluation failed")
            else if MetadataAttempt[HasError] then (try MetadataAttempt[Error][Message] otherwise "Mandatory Resource-period skeleton evaluation failed")
            else if not List.IsEmpty(MetadataSchemaAttempt[Value]) then "Missing upstream Stage1 fields: " & Text.Combine(MetadataSchemaAttempt[Value], ", ")
            else if MetadataAttempt[Value] > 0 then "Metadata is absent or unusable on these skeleton rows. Optional spare is excluded; allocation evidence uses the legacy status only where available."
            else null],
    EligibilitySchemaAttempt = try List.Difference({"Day", "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday"},
        Table.ColumnNames(#"IMPORT ResourcePeriodTABLE_empty")),
    EligibilityAttempt = try Table.RowCount(Table.SelectRows(A2EligibilityDay_CHECK, each [A2EligibilityDayStatus] <> "Available")),
    EligibilityCheck = [Stage = "A.2", Check = "Upstream spare eligibility interface",
        Status = if EligibilitySchemaAttempt[HasError] or EligibilityAttempt[HasError] then "Error"
            else if not List.IsEmpty(EligibilitySchemaAttempt[Value]) or EligibilityAttempt[Value] > 0 then "Error" else "Pass",
        Failures = if EligibilitySchemaAttempt[HasError] or EligibilityAttempt[HasError] then null else EligibilityAttempt[Value],
        Details = if EligibilitySchemaAttempt[HasError] then (try EligibilitySchemaAttempt[Error][Message] otherwise "Eligibility schema evaluation failed")
            else if EligibilityAttempt[HasError] then (try EligibilityAttempt[Error][Message] otherwise "Eligibility interface evaluation failed")
            else if not List.IsEmpty(EligibilitySchemaAttempt[Value]) then "Missing upstream eligibility fields: " & Text.Combine(EligibilitySchemaAttempt[Value], ", ")
            else if EligibilityAttempt[Value] > 0 then "Failure count is Resource/day groups with missing, invalid or inconsistent eligibility. Their Resources' optional spare is excluded."
            else null],
    PlannedCalendarCheck = A2CalendarDiagnosticCount("Saved eligibility plan calendar", () => Table.RowCount(A2PlannedCalendarIssues_Prepare)),
    ReductionCandidateCheck = A2DiagnosticCount("Joint reduction candidate budgets and cell amounts", () => Table.RowCount(A2ReductionCandidateIssues_Prepare)),
    JointReductionCheck = A2DiagnosticCount("Joint reduction decisions evaluate", () =>
        Table.RowCount(Table.SelectRowsWithErrors(A2JointReductionDecisions))),
    JointBudgetCheck = A2DiagnosticCount("Accepted joint removals stay within Resource and Period budgets", () => Table.RowCount(A2JointBudgetIssues_Prepare)),
    WholeOptionalCheck = A2DiagnosticCount("Optional C# retains or removes whole cells", () => Table.RowCount(A2WholeOptionalCellIssues_Prepare)),
    ResourceExclusionCheck = A2DiagnosticCount("Resources excluded from optional spare", () => let
        ExcludedRows = Table.SelectRows(#"A2 Resource Period Stage1 Prepare", each [SpareCalculationAllowed] = false),
        ResourceKeys = Table.SelectColumns(ExcludedRows, {"Resource"}),
        DistinctResources = Table.Distinct(ResourceKeys),
        ResourceCount = Table.RowCount(DistinctResources)
    in ResourceCount),
    PublicationErrorAttempt = try Table.RowCount(Table.SelectRowsWithErrors(#"ResPeriodCappedAvailability(C#)TABLE")),
    PublicationCheck = [Stage = "A.2", Check = "C# staging evaluation",
        Status = if PublicationErrorAttempt[HasError] then "Error" else if PublicationErrorAttempt[Value] = 0 then "Pass" else "Error",
        Failures = if PublicationErrorAttempt[HasError] then null else PublicationErrorAttempt[Value],
        Details = if PublicationErrorAttempt[HasError] then (try PublicationErrorAttempt[Error][Message] otherwise "Mandatory capacity input evaluation failed")
            else if PublicationErrorAttempt[Value] > 0 then "C# staging contains cell evaluation errors. No passing empty capacity result was substituted."
            else null],
    SpareExclusionCheck = A2DiagnosticCount("Excluded optional spare is not retained in C#", () =>
        Table.RowCount(Table.SelectRows(A2SparePublicationIssues_Prepare, each [Check] = "Excluded optional spare is not retained in C#"))),
    EligibilityExclusionCheck = A2DiagnosticCount("Unselected or ineligible optional spare is not retained in C#", () =>
        Table.RowCount(Table.SelectRows(A2SparePublicationIssues_Prepare, each [Check] = "Unselected or ineligible optional spare is not retained in C#"))),
    MaximumClusterCheck = A2CalendarDiagnosticCount("C# maximum five worked days per planned cluster", () =>
        Table.RowCount(Table.SelectRows(A2PublishedCalendarIssues_Prepare, each [Check] = "Maximum five worked days per planned cluster"))),
    MinimumGapCheck = A2CalendarDiagnosticCount("C# minimum two full off days between planned clusters", () =>
        Table.RowCount(Table.SelectRows(A2PublishedCalendarIssues_Prepare, each [Check] = "Minimum two full off days between planned clusters"))),
    OriginalCalendarCheck = A2CalendarDiagnosticCount("Preexisting allocation calendar breaches", () => Table.RowCount(A2OriginalCalendarIssues_Prepare)),
    CheckRows = List.Combine({OptionalChecks, {MetadataCheck, EligibilityCheck, PlannedCalendarCheck, ReductionCandidateCheck,
        JointReductionCheck, JointBudgetCheck, WholeOptionalCheck, ResourceExclusionCheck,
        PublicationCheck, SpareExclusionCheck, EligibilityExclusionCheck, MaximumClusterCheck, MinimumGapCheck, OriginalCalendarCheck}}),
    Output = Table.FromRecords(CheckRows, type table [Stage = text, Check = text, Status = text, Failures = nullable number, Details = nullable text])
in
    Table.Buffer(Output);

// Query: CapacityDiagnostics_DETAILS
// Purpose: Publish affected Resources, invalid plans and calendar breaches without stopping computable capacity publication.
// Output: Resource exclusions, eligibility issues and original/actual calendar evidence with numeric day/cluster proof columns.
shared CapacityDiagnostics_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Resource = nullable number, Role = nullable text,
        Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text,
        Day = nullable number, PlannedCluster = nullable number, PreviousCluster = nullable number,
        WorkedDays = nullable number, OffDays = nullable number, Origin = nullable text,
        CellAmount = nullable number, ResourceBudget = nullable number, PeriodBudget = nullable number,
        ResourceAcceptedRemoval = nullable number, PeriodAcceptedRemoval = nullable number, AcceptanceOrder = nullable number],
    OptionalInputs = {
        [Check = "Upstream cap priority interface", Assessment = A2CapPriorityInputAssessment],
        [Check = "Optional cap reduction input", Assessment = A2ReductionInputAssessment],
        [Check = "Legacy workday status input", Assessment = A2WorkdayInputAssessment],
        [Check = "Multiple-shift choice input", Assessment = A2ShiftChoiceInputAssessment]
    },
    OptionalIssueRows = List.Combine(List.Transform(OptionalInputs, each
        let
            CheckName = [Check],
            Assessment = [Assessment],
            Resources = if Assessment[HasGlobalIssue] then {null} else Assessment[AffectedResources]
        in if Assessment[Status] = "Pass" then {} else List.Transform(Resources, (resource) =>
            [Stage = "A.2", Check = CheckName, Resource = resource, Role = null, Name = null, Date = null, Period = null,
             Reason = if Assessment[Details] = null then "Optional input is unusable" else Assessment[Details],
             Action = if Assessment[HasGlobalIssue] then "Exclude all optional spare because the affected Resource cannot be identified; retain available allocation evidence."
                 else "Exclude all optional spare for this Resource; retain allocation evidence."]))),
    MetadataSchemaAttempt = try List.Difference({"OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence"}, Table.ColumnNames(#"IMPORT ResourcePeriodTABLE_empty")),
    SchemaIssueRows = if MetadataSchemaAttempt[HasError] then {
        [Stage = "A.2", Check = "Mandatory Resource-period schema", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try MetadataSchemaAttempt[Error][Message] otherwise "Mandatory Resource-period schema evaluation failed",
         Action = "Report capacity as unavailable; do not substitute a passing empty skeleton."]
    } else if not List.IsEmpty(MetadataSchemaAttempt[Value]) then {
        [Stage = "A.2", Check = "Upstream Stage1 diagnostic metadata", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = "Missing upstream Stage1 fields: " & Text.Combine(MetadataSchemaAttempt[Value], ", "),
         Action = "Exclude optional spare; use legacy allocated-period evidence only where available."]
    } else {},
    EligibilitySchemaAttempt = try List.Difference({"Day", "SpareDayEligible", "SparePeriodSelected", "SpareEligibilityReason", "PlannedCluster", "ExtensionSide", "AllocationWorkday"},
        Table.ColumnNames(#"IMPORT ResourcePeriodTABLE_empty")),
    EligibilitySchemaRows = if EligibilitySchemaAttempt[HasError] then {
        [Stage = "A.2", Check = "Upstream spare eligibility interface", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try EligibilitySchemaAttempt[Error][Message] otherwise "Eligibility schema evaluation failed",
         Action = "Report the unavailable interface; do not invent a passing eligibility plan."]
    } else if not List.IsEmpty(EligibilitySchemaAttempt[Value]) then {
        [Stage = "A.2", Check = "Upstream spare eligibility interface", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = "Missing upstream eligibility fields: " & Text.Combine(EligibilitySchemaAttempt[Value], ", "),
         Action = "Exclude all optional spare for affected Resources; retain available allocation evidence."]
    } else {},
    EligibilityDaysAttempt = try Table.Buffer(Table.SelectRows(A2EligibilityDay_CHECK, each [A2EligibilityDayStatus] <> "Available")),
    EligibilityDayRows = if EligibilityDaysAttempt[HasError] then {
        [Stage = "A.2", Check = "Upstream spare eligibility interface", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try EligibilityDaysAttempt[Error][Message] otherwise "Eligibility day validation failed",
         Action = "Report the unavailable interface; no passing eligibility replacement was invented."]
    } else List.Transform(Table.ToRecords(EligibilityDaysAttempt[Value]), each
        [Stage = "A.2", Check = "Upstream spare eligibility interface", Resource = [Resource], Role = null, Name = null, Date = null, Period = null,
         Day = [Day], Reason = [A2EligibilityDayReason],
         Action = "Exclude all optional spare for this Resource; retain available allocation evidence."]),
    PlannedCalendarAttempt = try Table.Buffer(A2PlannedCalendarIssues_Prepare),
    PlannedCalendarRows = if PlannedCalendarAttempt[HasError] then {
        [Stage = "A.2", Check = "Saved eligibility plan calendar", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try PlannedCalendarAttempt[Error][Message] otherwise "Saved eligibility calendar evaluation failed",
         Action = "Report the unavailable calendar; do not invent a passing plan."]
    } else List.Transform(Table.ToRecords(PlannedCalendarAttempt[Value]), each
        [Stage = "A.2", Check = "Saved eligibility plan calendar", Resource = [Resource], Role = [Role], Name = null, Date = null, Period = null,
         Day = [StartDay], PlannedCluster = [PlannedCluster], WorkedDays = [WorkedDays], OffDays = [OffDays],
         Reason = [Reason] & "; cluster day range " & Number.ToText([StartDay]) & " to " & Number.ToText([EndDay]),
         Action = "Exclude all optional spare for this Resource before C#; retain allocation evidence."]),
    ReductionCandidateAttempt = try let
        Rows = Table.Buffer(A2ReductionCandidateIssues_Prepare),
        ErrorRows = Table.SelectRowsWithErrors(Rows),
        CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else
            error Error.Record("CapacityDistribA2.ReductionCandidateDiagnosticErrors", "Reduction candidate diagnostics contain cell errors.", ErrorRows)
    in CheckedRows,
    ReductionCandidateRows = if ReductionCandidateAttempt[HasError] then {
        [Stage = "A.2", Check = "Joint reduction candidate budgets and cell amounts", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try ReductionCandidateAttempt[Error][Message] otherwise "Reduction candidate validation failed",
         Action = "Report the unavailable validation; do not invent passing budget values."]
    } else List.Transform(Table.ToRecords(ReductionCandidateAttempt[Value]), each
        [Stage = "A.2", Check = "Joint reduction candidate budgets and cell amounts", Resource = [Resource], Role = [Role], Name = null, Date = null, Period = [Period],
         Day = [Day], PlannedCluster = [PlannedCluster], Reason = [Reason],
         Action = "Exclude this Resource's optional spare and preserve allocation evidence."]),
    JointDecisionAttempt = try let
        Rows = Table.Buffer(A2JointReductionDecisions),
        ErrorRows = Table.SelectRowsWithErrors(Rows),
        CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else
            error Error.Record("CapacityDistribA2.JointDecisionDiagnosticErrors", "Joint reduction decisions contain cell errors.", ErrorRows)
    in CheckedRows,
    JointDecisionRows = if JointDecisionAttempt[HasError] then {
        [Stage = "A.2", Check = "Joint reduction decisions evaluate", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try JointDecisionAttempt[Error][Message] otherwise "Joint reduction evaluation failed",
         Action = "Exclude optional spare affected by this unavailable calculation; preserve allocation evidence."]
    } else List.Transform(Table.ToRecords(Table.SelectRows(JointDecisionAttempt[Value], each [Accepted] = false)), each
        [Stage = "A.2", Check = "Whole-cell joint reduction not accepted", Resource = [Resource], Role = [Role], Name = null, Date = null, Period = [Period],
         Day = [Day], PlannedCluster = [PlannedCluster], CellAmount = [CellAmount], ResourceBudget = [#"ResAv-C'"], PeriodBudget = [MaxPeriodReduction],
         ResourceAcceptedRemoval = [ResRunningTotal], PeriodAcceptedRemoval = [AcceptedPeriodRunningTotal], Reason = [Reason],
         Action = "Retain this whole optional cell at A.2; any later Resource-cap trim has its own decision and diagnostics."]),
    JointBudgetAttempt = try let
        Rows = Table.Buffer(A2JointBudgetIssues_Prepare),
        ErrorRows = Table.SelectRowsWithErrors(Rows),
        CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else
            error Error.Record("CapacityDistribA2.JointBudgetDiagnosticErrors", "Accepted budget diagnostics contain cell errors.", ErrorRows)
    in CheckedRows,
    JointBudgetRows = if JointBudgetAttempt[HasError] then {
        [Stage = "A.2", Check = "Accepted joint removals stay within Resource and Period budgets", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try JointBudgetAttempt[Error][Message] otherwise "Accepted budget audit failed",
         Action = "Report the unavailable audit; no passing budget reconciliation was substituted."]
    } else List.Transform(Table.ToRecords(JointBudgetAttempt[Value]), each
        [Stage = "A.2", Check = "Accepted joint removals stay within Resource and Period budgets", Resource = [Resource], Role = null, Name = null, Date = null, Period = [Period],
         Reason = [Reason], Action = "Review the accepted removal totals; preserve allocation evidence and continue other computable results."]),
    WholeCellAttempt = try let
        Rows = Table.Buffer(A2WholeOptionalCellIssues_Prepare),
        ErrorRows = Table.SelectRowsWithErrors(Rows),
        CheckedRows = if Table.IsEmpty(ErrorRows) then Rows else
            error Error.Record("CapacityDistribA2.WholeCellDiagnosticErrors", "Whole optional-cell diagnostics contain cell errors.", ErrorRows)
    in CheckedRows,
    WholeCellRows = if WholeCellAttempt[HasError] then {
        [Stage = "A.2", Check = "Optional C# retains or removes whole cells", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try WholeCellAttempt[Error][Message] otherwise "Whole optional-cell audit failed",
         Action = "Report the unavailable audit; no passing whole-cell result was substituted."]
    } else List.Transform(Table.ToRecords(WholeCellAttempt[Value]), each
        [Stage = "A.2", Check = "Optional C# retains or removes whole cells", Resource = [Resource], Role = [Role], Name = null, Date = null, Period = [Period],
         Day = [Day], PlannedCluster = [PlannedCluster], CellAmount = [CurrentCellAmount], Reason = [Reason],
         Action = "Review the partially retained optional cell; allocated cells remain independently preserved."]),
    SkeletonAttempt = try Table.Buffer(#"A2 Resource Period Stage1 Prepare"),
    ResourceIssueRows = if SkeletonAttempt[HasError] then {
        [Stage = "A.2", Check = "Mandatory Resource-period skeleton", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try SkeletonAttempt[Error][Message] otherwise "Mandatory Resource-period skeleton evaluation failed",
         Action = "Report capacity as unavailable; do not substitute a passing empty skeleton."]
    } else
        let
            ExcludedRows = Table.SelectRows(SkeletonAttempt[Value], each [SpareCalculationAllowed] = false),
            SelectedReasons = Table.SelectColumns(ExcludedRows, {"Resource", "Role", "Name", "SpareIssueReason", "A2Stage1MetadataStatus"}),
            ResourceReasons = Table.Distinct(SelectedReasons)
        in List.Transform(Table.ToRecords(ResourceReasons), each
            [Stage = "A.2", Check = if [A2Stage1MetadataStatus] = "Available" then "Resource optional spare exclusion" else "Upstream Stage1 diagnostic metadata",
             Resource = [Resource], Role = [Role], Name = [Name], Date = null, Period = null,
             Reason = if [SpareIssueReason] = null then "Resource is excluded from optional spare" else [SpareIssueReason],
             Action = "Exclude all optional spare for this Resource; retain allocation evidence."]),
    PublicationAttempt = try Table.Buffer(#"ResPeriodCappedAvailability(C#)TABLE"),
    PublicationIssueRows = if PublicationAttempt[HasError] then {
        [Stage = "A.2", Check = "C# staging evaluation", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try PublicationAttempt[Error][Message] otherwise "Mandatory capacity input evaluation failed",
         Action = "Report capacity as unavailable; no passing empty capacity result was substituted."]
    } else
        let ErrorRows = Table.SelectRowsWithErrors(PublicationAttempt[Value]) in
        List.Transform(Table.ToRecords(ErrorRows), each
            [Stage = "A.2", Check = "C# staging evaluation", Resource = try [Resource] otherwise null,
             Role = try [Role] otherwise null, Name = null, Date = null, Period = try [Period] otherwise null,
             Reason = "C# staging row contains evaluation errors.", Action = "Retain the issue for manual review; no passing replacement value was invented."]),
    SparePublicationAttempt = try Table.Buffer(A2SparePublicationIssues_Prepare),
    SparePublicationRows = if SparePublicationAttempt[HasError] then {
        [Stage = "A.2", Check = "Optional spare publication audit", Resource = null, Role = null, Name = null, Date = null, Period = null,
         Reason = try SparePublicationAttempt[Error][Message] otherwise "Optional spare publication audit failed",
         Action = "Report the unavailable audit; no passing replacement was invented."]
    } else List.Transform(Table.ToRecords(SparePublicationAttempt[Value]), each
        [Stage = "A.2", Check = [Check], Resource = [Resource], Role = [Role], Name = null, Date = null, Period = [Period],
         Day = [Day], PlannedCluster = [PlannedCluster], Reason = [Reason],
         Action = "Review the retained optional cell; preserve allocation evidence and the diagnostic record."]),
    ActualCalendarAttempt = try Table.Buffer(Table.SelectRows(A2PublishedCalendarIssues_Prepare, each [Origin] <> "Preexisting allocation")),
    OriginalCalendarAttempt = try Table.Buffer(A2OriginalCalendarIssues_Prepare),
    CalendarIssueRows = List.Combine(List.Transform({
        [Check = "Actual C# calendar", Attempt = ActualCalendarAttempt],
        [Check = "Preexisting allocation calendar", Attempt = OriginalCalendarAttempt]
    }, each let CheckName = [Check], Attempt = [Attempt] in
        if Attempt[HasError] then {
            [Stage = "A.2", Check = CheckName, Resource = null, Role = null, Name = null, Date = null, Period = null,
             Reason = try Attempt[Error][Message] otherwise "Calendar evaluation failed",
             Action = "Report the unavailable calendar; do not substitute a passing empty calendar."]
        } else List.Transform(Table.ToRecords(Attempt[Value]), (issue) =>
            [Stage = "A.2", Check = CheckName & ": " & issue[Check], Resource = issue[Resource], Role = issue[Role], Name = null, Date = null, Period = null,
             Day = issue[StartDay], PlannedCluster = issue[PlannedCluster], PreviousCluster = issue[PreviousCluster],
             WorkedDays = issue[WorkedDays], OffDays = issue[OffDays], Origin = issue[Origin],
             Reason = issue[Reason] & "; cluster day range " & Number.ToText(issue[StartDay]) & " to " & Number.ToText(issue[EndDay]),
             Action = if issue[Origin] = "Preexisting allocation" then "Retain allocation evidence and review this preexisting breach; optional spare is excluded."
                 else "Review the published optional-spare breach; diagnostics do not remove allocation evidence or stop other computable results."]))),
    AllIssueRows = List.Combine({OptionalIssueRows, SchemaIssueRows, EligibilitySchemaRows, EligibilityDayRows,
        PlannedCalendarRows, ReductionCandidateRows, JointDecisionRows, JointBudgetRows, WholeCellRows,
        ResourceIssueRows, PublicationIssueRows, SparePublicationRows, CalendarIssueRows}),
    // Fill only nullable proof fields explicitly; Table.FromRecords requires every declared nonnullable field in every record.
    ProofDefaults = [Day = null, PlannedCluster = null, PreviousCluster = null, WorkedDays = null, OffDays = null, Origin = null,
        CellAmount = null, ResourceBudget = null, PeriodBudget = null, ResourceAcceptedRemoval = null, PeriodAcceptedRemoval = null, AcceptanceOrder = null],
    CompleteIssueRows = List.Transform(AllIssueRows, each Record.Combine({ProofDefaults, _})),
    Output = Table.FromRecords(CompleteIssueRows, DetailType)
in
    Table.Buffer(Output);
