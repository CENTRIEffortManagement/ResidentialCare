// Power Query from: CapacityDistrib(A.2)-shifts.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\AIN\CapacityDistrib(A.2)-shifts.xlsx
// Extracted: 2026-10-05T07:55:20.402Z

section Section1;

[ Description = "BUFFER" ]
// Query: ResCapPrioritisedTABLE
// Purpose: Order the existing Resource reduction priorities with Period as the final tie-break.
shared ResCapPrioritisedTABLE = let
    Source = #"Period Running Total",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "PRCell", "Reduction", "ResAv-C'", "ClusterAllocation", "NWDPriority", "NWDType", "MaxPeriodReduction", "PeriodRunningTotal", "ReducePeriod"}),
    #"Sorted RES,NWDPRI,MAXPRT,PRT" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"MaxPeriodReduction", Order.Descending}, {"PeriodRunningTotal", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES,NWDPRI,MAXPRT,PRT", "ResIndex", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

shared ResCapIndexLimits = let
    Source = ResCapPrioritisedTABLE,
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "ResIndexX"}}),
    #"Grouped Rows" = Table.Group(#"Renamed Columns", {"Resource"}, {{"StartResIndex", each List.Min([ResIndexX]), type number}, {"ResUnassignedAvailable", each List.Max([#"ResAv-C'"]), type nullable number}, {"ResAvailNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
shared #"Resource Running total" = let
    Source = Table.NestedJoin(ResCapPrioritisedTABLE, {"Resource"}, ResCapIndexLimits, {"Resource"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartResIndex", "ResAvailNumber"}, {"StartResIndex", "ResAvailNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded REsLimits",{{"ResIndex", Order.Ascending}, {"ClusterAllocation", Order.Ascending}, {"PeriodRunningTotal", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction.1", each [ResIndex]- [StartResIndex]),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction.1", Int64.Type}}),
    LIST.BUFFER = List.Buffer(#"Changed Type"[Reduction]),
    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Changed Type", "ResRunningTotal", each List.Sum(List.Range(Source [Reduction], 
[StartResIndex]-1

, [Subtraction.1]))),
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "MaxResReduction", each List.Min({ [#"ResAv-C'"]})),
    #"Added KEEP" = Table.AddColumn(#"Inserted MAXRESTOTAL", "KeepRunningTotal", each if [ResRunningTotal] > [MaxResReduction] then "Remove for Res" else if [PeriodRunningTotal] > [MaxPeriodReduction] then "Remove for Perod" else if [ResRunningTotal] <= [MaxResReduction] then "Keep" else "Split"),
    #"Filtered KEEP" = Table.SelectRows(#"Added KEEP", each ([KeepRunningTotal] = "Keep")),
    #"Multiplied Column" = Table.TransformColumns(#"Filtered KEEP", {{"Reduction", each _ * -1, type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Multiplied Column",{"Resource", "Period", "PRCell", "Reduction", "ResRunningTotal"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

shared CheckCap = let
    Source = #"Resource Running total",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"MAXResRunningTotal", each List.Max([ResRunningTotal]), type number}, {"AvailabilityReduction", each List.Sum([Reduction]), type number}})
in
    #"Grouped Rows";

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

// Query: A2CapPriorityInputAssessment
// Purpose: Validate the saved A.1 priority schema and Resource/Period keys before sorting or reductions.
// Notes: ClusterAllocation may be null for days outside allocated clusters; key issues exclude only identifiable affected Resources.
shared A2CapPriorityInputAssessment = A2OptionalInputAssessment(() => #"IMPORT ResPeriodCapPrioritised",
    {"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "ClusterAllocation", "NWDPriority", "NWDType"}, "Period");

// Query: A2ReductionInputAssessment
// Purpose: Assess optional signed cap reductions without hiding a calculation or interface error.
shared A2ReductionInputAssessment = A2OptionalInputAssessment(() => #"Resource Running total", {"Resource", "Period", "Reduction"}, "Period");

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

// Query: A2 Resource Period Stage1 Prepare
// Purpose: Apply the authoritative A.1 Resource spare exclusions and diagnosed A.2 optional-input exclusions to every period.
// Output: Existing skeleton columns plus audit availability, spare permission, issue reason and allocation evidence.
// Notes: Missing Stage1 metadata is Error, never Pass. Allocation evidence remains nullable until the legacy status fallback is available.
shared #"A2 Resource Period Stage1 Prepare" = let
    Source = #"IMPORT ResourcePeriodTABLE_empty",
    Stage1Columns = {"OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence"},
    MissingColumns = List.Difference(Stage1Columns, Table.ColumnNames(Source)),
    WithMissingColumns = List.Accumulate(MissingColumns, Source, (state, column) => Table.AddColumn(state, column, each null)),
    Assessments = {A2CapPriorityInputAssessment, A2ReductionInputAssessment, A2WorkdayInputAssessment, A2ShiftChoiceInputAssessment},
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
    UpstreamExcludedRows = Table.SelectRows(WithMetadataStatus, each (try [SpareCalculationAllowed] otherwise null) = false),
    UpstreamExcludedResources = List.Buffer(List.Distinct(UpstreamExcludedRows[Resource])),
    ExcludedResources = List.Buffer(List.Distinct(List.Combine({MetadataIssueResources, UpstreamExcludedResources, OptionalIssueResources}))),
    WithSparePermission = Table.AddColumn(WithMetadataStatus, "A2SpareCalculationAllowed", each
        [A2Stage1MetadataStatus] = "Available"
            and (try [SpareCalculationAllowed] otherwise false) = true
            and not HasGlobalOptionalIssue
            and not List.Contains(ExcludedResources, [Resource]), type logical),
    WithIssueReason = Table.AddColumn(WithSparePermission, "A2SpareIssueReason", each
        let
            UpstreamReason = try [SpareIssueReason] otherwise null,
            Reasons = List.RemoveNulls({
                if Value.Is(UpstreamReason, type text) and Text.Trim(UpstreamReason) <> "" then UpstreamReason else null,
                if List.Contains(MetadataIssueResources, [Resource]) then "upstream Stage1 diagnostics unavailable for this Resource" else null,
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
// Purpose: Calculate C# while excluding all optional spare for issue Resources and preserving allocation evidence for review.
// Notes: Mandatory availability/skeleton evaluation errors remain errors; business diagnostics do not gate publication.
shared #"ResPeriodCappedAvailability(C#)TABLE" = let
    Source = #"A2 Availability Publication Prepare",
    
    #"Merged PREMPTY" = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"A2 Resource Period Stage1 Prepare", {"Role", "Resource", "Period"}, "ResourcePeriodTABLE_empty", JoinKind.RightOuter),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged PREMPTY",{"Resource", "Period"}),
    #"Expanded ResourcePeriodTABLE_empty" = Table.ExpandTableColumn(#"Removed Columns1", "ResourcePeriodTABLE_empty", {"Resource", "Period", "OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence", "A2Stage1MetadataStatus"}, {"Resource", "Period", "OriginalAvailability", "SpareCalculationAllowed", "SpareIssueReason", "HasAllocationEvidence", "A2Stage1MetadataStatus"}),
    
    #"Merged Queries" = Table.NestedJoin(#"Expanded ResourcePeriodTABLE_empty", {"Resource", "Period"}, #"A2 Selected Reductions Prepare", {"Resource", "Period"}, "ReCapAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReCapAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReCapAvailabilityTABLE", {"Reduction"}, {"Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ReCapAvailabilityTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Resource", "Period"}, #"A2 Workday Status Prepare", {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.FullOuter),
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
    // On an excluded Resource, preserve the existing available allocated-cell amount before optional shift choices or reductions.
    // Missing original availability still contributes zero; no allocation availability is fabricated.
    #"Inserted REDUCTION" = Table.AddColumn(PreparedAllocationEvidence, "AvailabilityCapped", each
        if [Availability] = null then 0
        else if [SpareCalculationAllowed] = false then
            if [HasAllocationEvidence] = true then [Availability] else 0
        else if [RosteredPeriodStatus] = "RosteredDay"
            or [RosteredPeriodStatus] = "UnavailableAllocatedPeriod"
            or [NoAllocationKeep] = false then 0
        else [Availability] + [Reduction]),
        #"Removed Columns" = Table.RemoveColumns(#"Inserted REDUCTION",{"Availability", "Reduction", "NoAllocationKeep", "A2Stage1MetadataStatus"}),
    #"Filled Down" = Table.FillDown(#"Removed Columns",{"Role"})
    in
        #"Filled Down";

[ Description = "BUFFER-All cells capped availability" ]
// Query: ResPeriodAvailabilityCapped(C#)TABLE
// Purpose: Publish C# with Stage1 spare exclusions and allocation/audit evidence on the existing saved interface.
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

// Query: PeriodCapPrioritisedTABLE
// Purpose: Order the existing period reduction priorities with Resource as the final tie-break.
shared PeriodCapPrioritisedTABLE = let
    Source = A2CapPriorityInputAssessment[Data],
    #"Sorted NWD,C-D,PA-Cn,PERIOD" = Table.Sort(Source,{{"Period", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"ClusterAllocation", Order.Descending}, {"C'-D", Order.Ascending}, {"PeriodA-CNeg", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted NWD,C-D,PA-Cn,PERIOD", "PeriodIndex",  2, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "PeriodIndexX", 1, 1, Int64.Type)
in
    #"Added Index1";

shared PeriodCapIndexLimits = let
    Source = PeriodCapPrioritisedTABLE,
    #"Grouped INDEX+COUNT" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([PeriodIndexX]), type number}, {"PeriodAvailNumber", each Table.RowCount(Table.Distinct(_)), Int64.Type}, {"PeriodResAv-C'", each List.Max([#"ResAv-C'"]), type nullable number}})
in
    #"Grouped INDEX+COUNT";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
shared #"Period Running Total" = let
    Source = Table.NestedJoin(PeriodCapPrioritisedTABLE, {"Period"}, PeriodCapIndexLimits, {"Period"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded PeriodLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartPeriodIndex", "PeriodAvailNumber", "PeriodResAv-C'"}, {"StartPeriodIndex", "PeriodAvailNumber", "PeriodResAv-C'"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PeriodLimits",{{"PeriodIndex", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction.1", each [PeriodIndex]-[StartPeriodIndex]),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction.1", Int64.Type}}),
    // Min reduction: PeriodA-C, PeriodC-D
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Changed Type", "MaxPeriodReduction", each List.Min({-1*[#"PeriodA-CNeg"], [#"C'-D"]})),
    #"BUFFER LIST" = List.Buffer(#"Inserted MAXRESTOTAL"[UnassignedAvail]),
    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Inserted MAXRESTOTAL", "PeriodRunningTotal", each List.Sum(List.Range(Source [UnassignedAvail], [StartPeriodIndex]-1

, [Subtraction.1]
))),
    #"Renamed Columns" = Table.RenameColumns(#"Added RUNNINGRESTOTAL",{{"UnassignedAvail", "Reduction"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ReducePeriod", each if [PeriodRunningTotal] <= [MaxPeriodReduction] then "RemovableByPeriod" else "UnremovableByPeriod"),
    BUFFER = Table.Buffer(#"Added Conditional Column")
in
    BUFFER;

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

shared #"IMPORT ResPeriodCapPrioritised" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodCapPrioritised_Table = Source{[Item="ResPeriodCapPrioritised",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodCapPrioritised_Table,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Allocation", type any}, {"PotentialAvailability", type any}, {"PRCell", type text}, {"UnassignedAvail", Int64.Type}, {"ResAv-C'", Int64.Type}, {"PeriodA-CNeg", type number}, {"C'-D", Int64.Type}, {"C'/D", type number}, {"ClusterAllocation", type number}, {"NWDPriority", Int64.Type}, {"NWDType", type text}})
in
    #"Changed Type";

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

shared #"IMPORT ResourcePeriodTABLE_empty" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResourcePeriodTABLE_empty_Table = Source{[Item="ResourcePeriodTABLE_empty",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResourcePeriodTABLE_empty_Table,{{"Name", type text},  {"Role", type text},     {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Shift", type any}, {"Day", Int64.Type}})
in
    #"Changed Type";

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

// Query: CapacityDiagnostics_SUMMARY
// Purpose: Publish nonblocking Stage1 check results for manual review; no result here gates C# publication.
// Output: One row per check, including Error when an input or required Stage1 metadata is unavailable.
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
        Table.RowCount(Table.SelectRows(#"ResPeriodCappedAvailability(C#)TABLE", each
            [SpareCalculationAllowed] = false and [HasAllocationEvidence] = false
                and [AvailabilityCapped] <> null and [AvailabilityCapped] > 0))),
    CheckRows = List.Combine({OptionalChecks, {MetadataCheck, ResourceExclusionCheck, PublicationCheck, SpareExclusionCheck}}),
    Output = Table.FromRecords(CheckRows, type table [Stage = text, Check = text, Status = text, Failures = nullable number, Details = nullable text])
in
    Table.Buffer(Output);

// Query: CapacityDiagnostics_DETAILS
// Purpose: Publish affected Resources and unavailable inputs without stopping computable capacity publication.
// Output: Resource-level spare exclusions and evaluation issues; allocation evidence is retained rather than silently replaced.
shared CapacityDiagnostics_DETAILS = let
    DetailType = type table [Stage = text, Check = text, Resource = nullable number, Role = nullable text,
        Name = nullable text, Date = nullable date, Period = nullable number, Reason = text, Action = text],
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
    Output = Table.FromRecords(List.Combine({OptionalIssueRows, SchemaIssueRows, ResourceIssueRows, PublicationIssueRows}), DetailType)
in
    Table.Buffer(Output);
