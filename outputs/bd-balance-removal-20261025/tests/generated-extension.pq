section BDBalanceRemovalNative;
shared BDBalanceRemovalTests = () as table => let
Evaluate = (Inputs as record) as record => let
ReconciledWorkers_CHECK = let
    Workers = ReconciledWorkers_Eligible,
    CurrentUnitPrepared = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit),
    CurrentUnitRoleKeys = List.Buffer(List.Distinct(List.RemoveNulls(CurrentUnitPrepared[SourceRoleKey]))),
    // Unused LINK Roles definitions do not participate in this unit and cannot block its publication.
    InvalidRoleDefinitions = Table.SelectRows(WorkerRoleMappings_Prepare, each
        [RoleMappingIssue] <> null and List.Contains(CurrentUnitRoleKeys, [RosterRoleKey])),
    RelevantRoleMappingIssues = Table.SelectRows(CurrentUnitPrepared, each
        [SourceRoleKey] <> null and [RoleMappingIssue] <> null),
    RelevantContractIssues = Table.SelectRows(ReconciledWorkers_Prepare, each
        [#"Facility-Abbrev"] = Unit and [Role] <> null and [ContractIssue] <> null),
    InvalidMembershipRows = Table.SelectRows(Workers, each
        (not [In Allocation] and not [In Records]) or [In Leave Balances]
            or not List.Contains({"Allocation only", "Records only", "Both"}, [Worker Membership])),
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
    CheckWorkerMembership = () as number =>
        Table.RowCount(InvalidMembershipRows),
    CheckPublishedNames = () as number =>
        DuplicateCount(Workers, {"Name"}),
    Checks = {
        AvailabilityCheckResult("Unique worker keys", CheckUniqueWorkerKeys),
        AvailabilityCheckResult("One preferred role per employee and facility", CheckPreferredRoleKeys),
        AvailabilityCheckResult("Valid LINK Roles definitions", CheckRoleDefinitions),
        AvailabilityCheckResult("Current worker role mappings are unambiguous", CheckWorkerRoleMappings),
        AvailabilityCheckResult("Valid relevant worker contracts", CheckRelevantContracts),
        AvailabilityCheckResult("Valid reconciled worker membership", CheckWorkerMembership),
        // Names alone must remain unique because the existing master list uses that key.
        AvailabilityCheckResult("Unambiguous published names", CheckPublishedNames)
    },
    CheckTable = Table.FromRecords(Checks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    // Small scalar check rows are reused by staff and shift publication within an evaluation.
    BufferedChecks = Table.Buffer(CheckTable)
in
    BufferedChecks,
AvailabilityCheckResult = (name as text, evaluate as function) as record =>
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
        Result,
Unit = Inputs[Unit],
ReconciledWorkers_Prepare = let
    DistinctNonNullValues = (values as list) as list =>
        let
            NonNullValues = List.RemoveNulls(values),
            DistinctValues = List.Distinct(NonNullValues)
        in
            DistinctValues,
    SourceNames = Table.Group(AvailabilityLeaveSource_Prepare, {"Payroll Code", "Facility-Abbrev"},
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
    SourceRoleKeys = Table.AddColumn(ReconciledWorkers_Facility_Prepare, "SourceRoleKey",
        each WorkerRoleKey([Role]), type nullable text),
    // Group on the normalized roster-role key before mapping so case or outer-space variants do not create duplicate workers.
    WorkerGroups = Table.Group(SourceRoleKeys,
        {"EmployeeID", "Facility-Abbrev", "SourceRoleKey"},
        {{"SourceRoles", each DistinctNonNullValues([Role]), type list},
         {"RosterNames", each DistinctNonNullValues([Employee Roster Name]), type list},
         {"EmploymentType", each CombineEmploymentTypes([EmploymentType]), type text},
         {"In Allocation", each List.Contains([In Allocation], true), type logical},
         {"In Records", each List.Contains([In Records], true), type logical},
         {"In Leave Balances", each List.Contains([In Leave Balances], true), type logical}}),
    SourceRoleLabels = Table.AddColumn(WorkerGroups, "SourceRole",
        each CombineSourceRoles([SourceRoles]), type nullable text),
    WorkerMembership = Table.AddColumn(SourceRoleLabels, "Worker Membership", each
        if [In Allocation] and [In Records] then "Both"
        else if [In Allocation] then "Allocation only"
        else if [In Records] then "Records only"
        else error "A worker must appear in Allocation or availability records.", type text),
    RemovedSourceRoleLists = Table.RemoveColumns(WorkerMembership, {"SourceRoles"}),
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
        else [ConfiguredRoleMappingIssue], type nullable text),
    RoleMappingStatus = Table.AddColumn(ResolvedRoleMappingIssues, "RoleMappingStatus", each
        if [SourceRoleKey] = null then "Missing source role"
        else if [RoleMappingMatches] = 0 then "LINK Roles role excluded: no matching Roster Roles"
        else if [RoleMappingIssue] <> null then "Invalid LINK Roles definition"
        else if [MappedRoleGroup] = null then "LINK Roles role excluded: no Role Group"
        else if not List.Contains(AvailabilityRoleGroups, [MappedRoleGroup]) then
            "Mapped Role Group outside availability scope"
        else "Mapped to availability Role Group", type text),
    // LINK Roles remains authoritative for translation; the availability model publishes only its configured Role Groups.
    MappedRole = Table.AddColumn(RoleMappingStatus, "Role", each
        if [RoleMappingStatus] = "Mapped to availability Role Group" then [MappedRoleGroup]
        else null, type nullable text),
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
                else if row[RoleMappingStatus] = "LINK Roles role excluded: no matching Roster Roles" then
                    "LINK Roles role excluded: no matching Roster Roles"
                else if row[RoleMappingStatus] = "LINK Roles role excluded: no Role Group" then
                    "LINK Roles role excluded: no Role Group"
                else if row[RoleMappingStatus] = "Mapped Role Group outside availability scope" then
                    "Mapped Role Group outside availability scope"
                else if row[Role] = null then "LINK Roles did not assign an availability Role Group"
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
    BufferedResult,
FacilityAnalysisTABLE = #table({"Facility-Abbrev"},{{Inputs[Unit]}}),
WorkerContracts_Prepare = let
    ContractColumns = Table.SelectColumns(WorkerEmployment_Prepare, {"Employee Code", "Contracted FN Hours"}),
    NormalizedEmployee = Table.TransformColumns(ContractColumns,
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
    BufferedOutput,
AvailabilityText = (value as any) as nullable text =>
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
        BlankAsNull,
WorkerEmployment_Prepare = let
    Source = #"IMPORT Worker's Report",
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    SelectedColumns = Table.SelectColumns(PromotedHeaders,
        {"Employee Code", "Location Description", "Position Description", "Employee Status Description",
         "Start Date", "Termination Date", "Contracted FN Hours"}),
    // Type conversion errors remain errors for the existing contract checks; no invalid hours become a fallback value.
    TypedColumns = Table.TransformColumnTypes(SelectedColumns,
        {{"Employee Code", type text}, {"Location Description", type text}, {"Position Description", type text},
         {"Employee Status Description", type text}, {"Start Date", type date},
         {"Termination Date", type nullable date}, {"Contracted FN Hours", type nullable number}}),
    NormalizeText = (value as nullable text) as nullable text =>
        let
            Trimmed = if value = null then null else Text.Trim(value)
        in
            if Trimmed = "" then null else Trimmed,
    NormalizedFields = Table.TransformColumns(TypedColumns,
        List.Transform({"Employee Code", "Location Description", "Position Description", "Employee Status Description"},
            each {_, NormalizeText, type nullable text})),
    WithFacility = Table.AddColumn(NormalizedFields, "Facility-Abbrev", each
        if [Location Description] = null then null else Text.Upper(Text.Start([Location Description], 2)), type nullable text),
    EmploymentColumns = Table.SelectColumns(WithFacility,
        {"Employee Code", "Facility-Abbrev", "Position Description", "Employee Status Description",
         "Start Date", "Termination Date", "Contracted FN Hours"})
in
    Table.Buffer(EmploymentColumns),
#"IMPORT Worker's Report" = Inputs[HR],
AvailabilityRoleGroups = List.Buffer({"RN", "AIN", "EN"}),
WorkerRoleMappings_Prepare = let
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
            // A roster role with no Role Group is an intentional non-availability definition, not a broken mapping.
            RoleMappingIssue = if RosterRoleKey = null then
                    "LINK Roles has a blank Roster Roles value"
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
    BufferedMappings,
WorkerRoleKey = (roleName as any) as nullable text =>
    let
        CleanRole = AvailabilityText(roleName),
        UpperRole = if CleanRole = null then null else Text.Upper(CleanRole),
        AlphaNumericCharacters = List.Transform({48..57}, each Character.FromNumber(_))
            & List.Transform({65..90}, each Character.FromNumber(_)),
        RemoveSuffix = (value as nullable text, suffix as text) as nullable text =>
            if value = null then null
            else if Text.EndsWith(value, suffix) and Text.Length(value) > Text.Length(suffix) then
                Text.Start(value, Text.Length(value) - Text.Length(suffix))
            else value,
        // A leading space makes these suffixes whole shift tokens rather than endings in an ordinary word.
        WithoutShiftQualifier = if UpperRole = null then null
            else List.Accumulate({" AM", " PM", " NS"}, UpperRole, RemoveSuffix),
        WithoutInChargeQualifier = if WithoutShiftQualifier = null then null
            else Text.Replace(WithoutShiftQualifier, "IN CHARGE", ""),
        CompactRole = if WithoutInChargeQualifier = null then null
            else Text.Select(WithoutInChargeQualifier, AlphaNumericCharacters),
        // The authoritative roster list abbreviates Assistant as Asst; reconciliation spells it in full.
        AsstPrefix = "ASSTINNURSING",
        AssistantPrefix = "ASSISTANTINNURSING",
        ExpandedAssistant = if CompactRole <> null
                and Text.StartsWith(CompactRole, AsstPrefix) then
            AssistantPrefix & Text.Range(CompactRole, Text.Length(AsstPrefix))
        else CompactRole,
        RoleKey = if ExpandedAssistant = "" then null else ExpandedAssistant
    in
        RoleKey,
#"LINK Roles" = #table({"Roster Roles","Role Group"},{{"Assistant in Nursing AM","AIN"},{"Enrolled Nurse AM","EN"},{"Registered Nurse AM","RN"},{"REGN AM","RN"}}),
ReconciledWorkers_Facility_Prepare = let
    Source = ReconciledWorkerMembership_Prepare,
    PreservedSourceFacility = Table.RenameColumns(Source, {{"Facility-Abbrev", "SourceFacility"}}),
    NormalizedFacilityKey = Table.AddColumn(PreservedSourceFacility, "FacilityKey", each
        if [SourceFacility] = null then null else Text.Upper(Text.Trim([SourceFacility])), type nullable text),
    MergedFacilityMapping = Table.NestedJoin(NormalizedFacilityKey, {"FacilityKey"},
        FacilityMappings_Prepare, {"FacilityKey"}, "FacilityMapping", JoinKind.LeftOuter),
    ExpandedCanonicalFacility = Table.ExpandTableColumn(MergedFacilityMapping, "FacilityMapping",
        {"Facility-Abbrev"}, {"MappedFacility"}),
    IdentifiedUnmappedFacilities = Table.AddColumn(ExpandedCanonicalFacility, "FacilityMappingStatus", each
        if [FacilityKey] = null or [FacilityKey] = "" then "Missing facility"
        else if [MappedFacility] = null then "Unmapped facility: inspect LINK Facilities"
        else "Mapped from LINK Facilities", type text),
    ResolvedCanonicalFacility = Table.AddColumn(IdentifiedUnmappedFacilities, "Facility-Abbrev", each
        if [MappedFacility] <> null then [MappedFacility] else [SourceFacility], type nullable text),
    RemovedJoinHelpers = Table.RemoveColumns(ResolvedCanonicalFacility, {"FacilityKey", "MappedFacility"}),
    ReorderedColumns = Table.ReorderColumns(RemovedJoinHelpers,
        Table.ColumnNames(Source) & {"SourceFacility", "FacilityMappingStatus"})
in
    ReorderedColumns,
FacilityMappings_Prepare = let
    Source = #"LINK Facilities",
    CanonicalColumn = if Table.HasColumns(Source, "Facility-Abbrev") then "Facility-Abbrev" else "field_1",
    RequiredColumns = {"Title", CanonicalColumn, "FacilityShort"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Source)),
    SelectedFacilityColumns = if List.IsEmpty(MissingColumns) then Table.SelectColumns(Source, RequiredColumns)
        else error Error.Record("Facility mapping columns",
            "LINK Facilities must expose Title, FacilityShort and Facility-Abbrev (legacy field_1). Inspect its source column names.",
            [MissingColumns = MissingColumns]),
    NamedFacilityColumns = if CanonicalColumn = "Facility-Abbrev" then SelectedFacilityColumns
        else Table.RenameColumns(SelectedFacilityColumns, {{CanonicalColumn, "Facility-Abbrev"}}),
    NormalizedFacilityColumns = Table.TransformColumns(NamedFacilityColumns, {
        {"Title", each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
        {"FacilityShort", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type nullable text},
        {"Facility-Abbrev", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type nullable text}}),
    MissingCanonicalFacilities = Table.SelectRows(NormalizedFacilityColumns, each
        ([Title] <> null and [Title] <> "" or [FacilityShort] <> null and [FacilityShort] <> "")
        and ([#"Facility-Abbrev"] = null or [#"Facility-Abbrev"] = "")),
    ValidatedCanonicalFacilities = if Table.IsEmpty(MissingCanonicalFacilities) then NormalizedFacilityColumns
        else error Error.Record("Facility mapping values",
            "A Facilities title or short code has no canonical Facility-Abbrev.", MissingCanonicalFacilities),
    PreservedCanonicalFacility = Table.DuplicateColumn(ValidatedCanonicalFacilities,
        "Facility-Abbrev", "CanonicalFacility"),
    UnpivotedFacilityAliases = Table.Unpivot(PreservedCanonicalFacility,
        {"Title", "FacilityShort", "Facility-Abbrev"}, "AliasSource", "FacilityAlias"),
    NonblankFacilityAliases = Table.SelectRows(UnpivotedFacilityAliases, each
        [FacilityAlias] <> null and [FacilityAlias] <> ""),
    NormalizedFacilityKeys = Table.AddColumn(NonblankFacilityAliases, "FacilityKey", each
        Text.Upper([FacilityAlias]), type text),
    SelectedMappingPairs = Table.SelectColumns(NormalizedFacilityKeys, {"FacilityKey", "CanonicalFacility"}),
    DistinctMappingPairs = Table.Distinct(SelectedMappingPairs),
    GroupedFacilityMappings = Table.Group(DistinctMappingPairs, {"FacilityKey"},
        {{"CanonicalFacilities", each [CanonicalFacility], type list}}),
    CountedCanonicalMappings = Table.AddColumn(GroupedFacilityMappings, "CanonicalCount", each
        List.Count([CanonicalFacilities]), Int64.Type),
    ConflictingFacilityMappings = Table.SelectRows(CountedCanonicalMappings, each [CanonicalCount] <> 1),
    ValidatedFacilityMappings = if Table.IsEmpty(ConflictingFacilityMappings) then CountedCanonicalMappings
        else error Error.Record("Conflicting facility mappings",
            "A Facilities title, short code or canonical code maps to more than one Facility-Abbrev.", ConflictingFacilityMappings),
    ExpandedCanonicalMappings = Table.ExpandListColumn(ValidatedFacilityMappings, "CanonicalFacilities"),
    NamedCanonicalMappings = Table.RenameColumns(ExpandedCanonicalMappings, {{"CanonicalFacilities", "Facility-Abbrev"}}),
    SelectedLookupColumns = Table.SelectColumns(NamedCanonicalMappings, {"FacilityKey", "Facility-Abbrev"}),
    TypedLookup = Table.TransformColumnTypes(SelectedLookupColumns,
        {{"FacilityKey", type text}, {"Facility-Abbrev", type text}}),
    // Only the small unique scalar lookup is buffered, after alias expansion and conflict validation.
    BufferedLookup = Table.Buffer(TypedLookup)
in
    BufferedLookup,
#"LINK Facilities" = Inputs[Facilities],
ReconciledWorkerMembership_Prepare = let
    RosterDates = List.RemoveNulls(PublishedRoster_Prepare[Date]),
    RosterStart = if List.IsEmpty(RosterDates) then error "Published Roster contains no roster dates." else List.Min(RosterDates),
    JoinedMembership = Table.NestedJoin(AvailabilityMembership_Prepare, {"EmployeeID", "Facility-Abbrev"},
        RosterWorkers_Prepare, {"Employee Code", "Facility-Abbrev"}, "Roster Worker", JoinKind.FullOuter),
    ExpandedRoster = Table.ExpandTableColumn(JoinedMembership, "Roster Worker",
        {"Employee Code", "Facility-Abbrev", "Employee Roster Name", "Employment Type", "Role", "In Allocation Source"},
        {"Roster EmployeeID", "Roster Facility", "Roster Name", "Roster Employment Type", "Roster Role", "In Allocation Source"}),
    WithEmployeeID = Table.AddColumn(ExpandedRoster, "Worker EmployeeID", each [EmployeeID] ?? [Roster EmployeeID], type nullable text),
    WithFacility = Table.AddColumn(WithEmployeeID, "Worker Facility", each [Roster Facility] ?? [#"Facility-Abbrev"], type nullable text),
    WithRosterFlag = Table.AddColumn(WithFacility, "In Allocation", each [In Allocation Source] = true, type logical),
    WithAvailabilityFlag = Table.AddColumn(WithRosterFlag, "In Records", each [In Records Source] = true, type logical),
    MembershipColumns = Table.SelectColumns(WithAvailabilityFlag,
        {"Worker EmployeeID", "Worker Facility", "Roster Name", "Records Employee Name", "Roster Employment Type", "Roster Role", "In Allocation", "In Records"}),
    NamedWorkerKeys = Table.RenameColumns(MembershipColumns,
        {{"Worker EmployeeID", "EmployeeID"}, {"Worker Facility", "Facility-Abbrev"}}),
    JoinedEmployment = Table.NestedJoin(NamedWorkerKeys, {"EmployeeID"}, WorkerEmployment_Selected,
        {"Employee Code"}, "Worker Employment", JoinKind.LeftOuter),
    ExpandedEmployment = Table.ExpandTableColumn(JoinedEmployment, "Worker Employment",
        {"Employee Code", "Position Description", "Employee Status Description", "Termination Date", "Current Worker Eligible Row Count"},
        {"HR EmployeeID", "Position Description", "Employee Status Description", "Termination Date", "Current Worker Eligible Row Count"}),
    WithHRFlag = Table.AddColumn(ExpandedEmployment, "In HR Records", each [HR EmployeeID] <> null, type logical),
    // Preserve the original exclusion strictly before roster start, including when a terminated person remains on the roster.
    ActiveWorkers = Table.SelectRows(WithHRFlag, each not [In HR Records]
        or [Termination Date] = null or [Termination Date] >= RosterStart),
    WithName = Table.AddColumn(ActiveWorkers, "Employee Roster Name", each [Roster Name] ?? [Records Employee Name], type nullable text),
    WithRole = Table.AddColumn(WithName, "Role", each [Roster Role] ?? [Position Description], type nullable text),
    WithEmploymentType = Table.AddColumn(WithRole, "EmploymentType", each
        let
            EmploymentType = [Roster Employment Type] ?? [Employee Status Description]
        in
            if EmploymentType = null then null else Text.Replace(EmploymentType, " ", "-"), type nullable text),
    // Keep the original preferred-role ranking so multiple roster roles produce one stable employee/facility worker.
    RoleAlignmentRank = (role as nullable text, position as nullable text) as number =>
        let
            RoleKey = if role = null then null else Text.Upper(Text.Trim(role)),
            PositionKey = if position = null then null else Text.Upper(Text.Trim(position))
        in
            if RoleKey = null then 9
            else if PositionKey = null then 2
            else if RoleKey = PositionKey then 0
            else if PositionKey = "REGISTERED NURSE" and Text.StartsWith(RoleKey, "REGN") then 0
            else if PositionKey = "ASSISTANT IN NURSING" and RoleKey = "ASSISTANT IN NURSING" then 0
            else if PositionKey = "ENROLLED NURSE" and RoleKey = "ENROLLED NURSE" then 0
            else if PositionKey = "ASSISTANT IN NURSING" and Text.Contains(RoleKey, "ASSISTANT IN NURSING") then 1
            else 2,
    WithRoleRank = Table.AddColumn(WithEmploymentType, "Role Alignment Rank", each RoleAlignmentRank([Role], [Position Description]), Int64.Type),
    PreferredColumns = {"Role", "Employee Roster Name", "EmploymentType", "In Allocation", "In Records", "In HR Records", "Current Worker Eligible Row Count"},
    PreferredCandidates = Table.SelectColumns(WithRoleRank, {"EmployeeID", "Facility-Abbrev", "Role Alignment Rank"} & PreferredColumns),
    PreferredWorkers = Table.Group(PreferredCandidates, {"EmployeeID", "Facility-Abbrev"}, {
        {"Preferred Worker", each Table.First(Table.Sort(_, {
            {"Role Alignment Rank", Order.Ascending}, {"Role", Order.Ascending},
            {"Employee Roster Name", Order.Ascending}, {"EmploymentType", Order.Ascending}})), type record}
    }),
    ExpandedPreferredWorker = Table.ExpandRecordColumn(PreferredWorkers, "Preferred Worker", PreferredColumns, PreferredColumns),
    // Carry forward only the existing non-Balance publication gates, after selecting each participating worker's preferred role.
    WithMembershipIssue = Table.AddColumn(ExpandedPreferredWorker, "Membership Issue", each
        if [EmployeeID] = null or [#"Facility-Abbrev"] = null or [Role] = null or [Employee Roster Name] = null then "Incomplete worker identity"
        else if not [In Allocation] and not [In Records] then "Invalid worker membership"
        else if not [In Allocation] and not [In HR Records] then "Availability worker has no current worker match"
        else if [Current Worker Eligible Row Count] <> null and [Current Worker Eligible Row Count] > 1 then "Current worker record ambiguous"
        else null, type nullable text),
    MembershipFailures = Table.SelectRows(WithMembershipIssue, each [Membership Issue] <> null),
    ValidatedWorkers = if Table.IsEmpty(MembershipFailures) then WithMembershipIssue
        else error Error.Record("Capacity worker membership", "Required roster, availability and employment checks failed.",
            Table.SelectColumns(MembershipFailures,
                {"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Membership Issue"})),
    WithLegacyBalanceFlag = Table.AddColumn(ValidatedWorkers, "In Leave Balances", each false, type logical),
    Output = Table.SelectColumns(WithLegacyBalanceFlag,
        {"EmployeeID", "Facility-Abbrev", "Role", "Employee Roster Name", "EmploymentType",
         "In Allocation", "In Records", "In Leave Balances", "In HR Records"})
in
    Table.Buffer(Output),
WorkerEmployment_Selected = let
    RosterDates = List.RemoveNulls(PublishedRoster_Prepare[Date]),
    RosterStart = if List.IsEmpty(RosterDates) then error "Published Roster contains no roster dates." else List.Min(RosterDates),
    IdentifiedWorkers = Table.SelectRows(WorkerEmployment_Prepare, each [Employee Code] <> null),
    OpenEmploymentRank = Table.AddColumn(IdentifiedWorkers, "Open Employment Rank", each
        if [Termination Date] = null then 0 else 1, Int64.Type),
    RosterEligibilityRank = Table.AddColumn(OpenEmploymentRank, "Roster Eligibility Rank", each
        if [Termination Date] = null or [Termination Date] >= RosterStart then 0 else 1, Int64.Type),
    GroupedWorkers = Table.Group(RosterEligibilityRank, {"Employee Code"}, {
        {"Selected Worker", each Table.First(Table.Sort(_, {
            {"Roster Eligibility Rank", Order.Ascending}, {"Open Employment Rank", Order.Ascending},
            {"Termination Date", Order.Descending}, {"Start Date", Order.Descending},
            {"Facility-Abbrev", Order.Ascending}, {"Position Description", Order.Ascending}})), type record},
        {"Current Worker Eligible Row Count", each
            let
                EligibleRows = Table.SelectRows(_, each [Termination Date] = null or [Termination Date] >= RosterStart),
                MaterialFields = Table.SelectColumns(EligibleRows,
                    {"Facility-Abbrev", "Position Description", "Employee Status Description", "Termination Date", "Contracted FN Hours"})
            in
                Table.RowCount(Table.Distinct(MaterialFields)), Int64.Type}
    }),
    ExpandedWorker = Table.ExpandRecordColumn(GroupedWorkers, "Selected Worker",
        {"Position Description", "Employee Status Description", "Termination Date"},
        {"Position Description", "Employee Status Description", "Termination Date"})
in
    Table.Buffer(ExpandedWorker),
PublishedRoster_Prepare = let
    Source = #"IMPORT Published Roster",
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    SelectedColumns = Table.SelectColumns(PromotedHeaders,
        {"Location", "Department", "Employee Roster Name", "Employment Type", "Role", "Employee Code", "Date"}),
    TypedColumns = Table.TransformColumnTypes(SelectedColumns,
        {{"Location", type text}, {"Department", type text}, {"Employee Roster Name", type text},
         {"Employment Type", type text}, {"Role", type text}, {"Employee Code", type text}, {"Date", type date}}),
    SourceLocations = {"BD: Beaudesert", "JE: Jessie Hunt", "NR: Narrabri- Robert Young", "TE: Temora ", "TE: Temora Greenstone"},
    AnalysisLocations = Table.SelectRows(TypedColumns, each List.Contains(SourceLocations, [Location])),
    // Only a terminal shift token is removed; roster worker roles otherwise retain their complete source value.
    BaseRoles = Table.AddColumn(AnalysisLocations, "Worker Role", each
        if [Department] = "Care" or [Department] = "Catering" then
            let
                RosterRole = if [Role] = null then null else Text.Trim([Role]),
                HasShiftSuffix = RosterRole <> null and List.AnyTrue(List.Transform({" AM", " PM", " NS"},
                    each Text.EndsWith(Text.Upper(RosterRole), _))),
                BaseRole = if RosterRole = null then null
                    else if HasShiftSuffix then Text.TrimEnd(Text.Start(RosterRole, Text.Length(RosterRole) - 2))
                    else RosterRole,
                ExpandedAssistant = if BaseRole = null then null else Text.Replace(BaseRole, "Asst", "Assistant")
            in
                ExpandedAssistant
        else null, type nullable text)
in
    BaseRoles,
#"IMPORT Published Roster" = Inputs[Roster],
RosterWorkers_Prepare = let
    Source = PublishedRoster_Prepare,
    SelectedColumns = Table.SelectColumns(Source,
        {"Location", "Department", "Employee Roster Name", "Employment Type", "Employee Code", "Worker Role"}),
    NamedRole = Table.RenameColumns(SelectedColumns, {{"Worker Role", "Role"}}),
    WithFacility = Table.AddColumn(NamedRole, "Facility-Abbrev", each
        let
            Location = if [Location] = null then null else Text.Trim([Location])
        in
            if Location = null or Location = "" then null else Text.Upper(Text.Start(Location, 2)), type nullable text),
    NormalizeText = (value as nullable text) as nullable text =>
        let
            Trimmed = if value = null then null else Text.Trim(value)
        in
            if Trimmed = "" then null else Trimmed,
    NormalizedFields = Table.TransformColumns(WithFacility,
        List.Transform({"Employee Code", "Employee Roster Name", "Employment Type", "Role"},
            each {_, NormalizeText, type nullable text})),
    NormalizedInChargeRole = Table.ReplaceValue(NormalizedFields, "REGN -  In Charge", "REGN", Replacer.ReplaceText, {"Role"}),
    NormalizedOtherInChargeRole = Table.ReplaceValue(NormalizedInChargeRole, "REGN - In Charge", "REGN", Replacer.ReplaceText, {"Role"}),
    NormalizedAssistantRole = Table.ReplaceValue(NormalizedOtherInChargeRole, "Assistant in Nursing ", "Assistant in Nursing", Replacer.ReplaceText, {"Role"}),
    DistinctWorkerColumns = List.RemoveItems(Table.ColumnNames(NormalizedAssistantRole), {"Location"}),
    DistinctWorkers = Table.Distinct(NormalizedAssistantRole, DistinctWorkerColumns),
    IdentifiedWorkers = Table.SelectRows(DistinctWorkers, each [Employee Code] <> null),
    NonAgencyWorkers = Table.SelectRows(IdentifiedWorkers, each [Employment Type] <> "Agency"),
    MembershipColumns = Table.SelectColumns(NonAgencyWorkers,
        {"Employee Code", "Facility-Abbrev", "Employee Roster Name", "Employment Type", "Role"}),
    WithRosterMembership = Table.AddColumn(MembershipColumns, "In Allocation Source", each true, type logical)
in
    Table.Buffer(WithRosterMembership),
AvailabilityMembership_Prepare = let
    Source = #"IMPORT Availability Leave Source",
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    SelectedColumns = Table.SelectColumns(PromotedHeaders, {"Site Name", "Employee Name", "PayrollCode"}),
    NamedEmployee = Table.RenameColumns(SelectedColumns, {{"PayrollCode", "EmployeeID"}}),
    TypedColumns = Table.TransformColumnTypes(NamedEmployee,
        {{"Site Name", type text}, {"Employee Name", type text}, {"EmployeeID", type text}}),
    ClearedInvalidMatchKeys = Table.ReplaceErrorValues(TypedColumns, {{"EmployeeID", null}, {"Site Name", null}}),
    NormalizeText = (value as nullable text) as nullable text =>
        let
            Trimmed = if value = null then null else Text.Trim(value)
        in
            if Trimmed = "" then null else Trimmed,
    NormalizedFields = Table.TransformColumns(ClearedInvalidMatchKeys,
        {{"EmployeeID", NormalizeText, type nullable text}, {"Employee Name", NormalizeText, type nullable text}}),
    WithFacility = Table.AddColumn(NormalizedFields, "Facility-Abbrev", each
        let
            SiteName = if [Site Name] = null then null else Text.Trim([Site Name])
        in
            if SiteName = null or Text.Length(SiteName) < 2 then null else Text.Upper(Text.Start(SiteName, 2)), type nullable text),
    WorkerLocationColumns = Table.SelectColumns(WorkerEmployment_Prepare, {"Employee Code", "Facility-Abbrev"}),
    CompleteWorkerLocations = Table.SelectRows(WorkerLocationColumns, each
        [Employee Code] <> null and [#"Facility-Abbrev"] <> null and Text.Length([#"Facility-Abbrev"]) = 2),
    WorkerLocations = Table.Buffer(Table.Distinct(CompleteWorkerLocations)),
    // The distinct lookup preserves every matching source row without multiplying worker candidates.
    MatchedLocations = Table.NestedJoin(WithFacility, {"EmployeeID", "Facility-Abbrev"},
        WorkerLocations, {"Employee Code", "Facility-Abbrev"}, "Worker Location", JoinKind.Inner),
    RemovedLocationMatch = Table.RemoveColumns(MatchedLocations, {"Worker Location"}),
    SourceLocations = {"BD: Beaudesert", "JE: Jessie Hunt", "NR: Narrabri- Robert Young", "TE: Temora"},
    AnalysisLocations = Table.SelectRows(RemovedLocationMatch, each List.Contains(SourceLocations, [Site Name])),
    WorkerAccounts = Table.SelectRows(AnalysisLocations, each [EmployeeID] <> "999999"
        and (if [EmployeeID] = null then true else Comparer.OrdinalIgnoreCase([EmployeeID], "HF") <> 0)),
    GroupedWorkers = Table.Group(WorkerAccounts, {"EmployeeID", "Facility-Abbrev"}, {
        {"Records Employee Name", each
            let
                Names = List.Sort(List.Distinct(List.RemoveNulls([Employee Name]), Comparer.OrdinalIgnoreCase), Comparer.OrdinalIgnoreCase)
            in
                if List.IsEmpty(Names) then null else Names{0}, type nullable text}
    }),
    WithAvailabilityMembership = Table.AddColumn(GroupedWorkers, "In Records Source", each true, type logical)
in
    Table.Buffer(WithAvailabilityMembership),
#"IMPORT Availability Leave Source" = Inputs[Availability],
AvailabilityLeaveSource_Prepare = let
    Source = #"IMPORT Availability Leave Source",
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    NamedSourceFields = Table.RenameColumns(PromotedHeaders,
        {{"PayrollCode", "Payroll Code"}, {"DateTimeStart", "Date From"},
         {"DateTimeEnd", "Date To"}, {"Comments", "Availability or Leave Reason"}}),
    NormalizedFacilityKey = Table.AddColumn(NamedSourceFields, "FacilityKey", each
        if [Site Name] = null then null else Text.Upper(Text.Trim(Text.From([Site Name]))), type nullable text),
    // Worker short codes and source titles resolve through the same unique canonical facility lookup.
    JoinedFacilities = Table.NestedJoin(NormalizedFacilityKey, {"FacilityKey"}, FacilityMappings_Prepare,
        {"FacilityKey"}, "FacilityLookup", JoinKind.LeftOuter),
    ExpandedFacilities = Table.ExpandTableColumn(JoinedFacilities, "FacilityLookup", {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    FilteredUnit = Table.SelectRows(ExpandedFacilities, each [#"Facility-Abbrev"] = Unit and AnalysisEnabled),
    SelectedColumns = Table.SelectColumns(FilteredUnit,
        {"Site Name", "Employee Name", "Payroll Code", "Type", "Date From", "Date To",
         "Availability or Leave Reason", "Facility-Abbrev"}),
    NormalizedIdentity = Table.TransformColumns(SelectedColumns,
        {{"Payroll Code", AvailabilityText, type nullable text}, {"Employee Name", AvailabilityText, type nullable text}}),
    IndexedRecords = Table.AddIndexColumn(NormalizedIdentity, "SourceRecord", 1, 1, Int64.Type),
    // Buffer the reduced scalar source shared by name resolution, parsing and diagnostics.
    BufferedRecords = Table.Buffer(IndexedRecords)
in
    BufferedRecords,
ReconciledWorkers_Eligible = let
    Included = Table.SelectRows(ReconciledWorkers_Prepare, each [Issue] = null),
    SelectedWorkers = Table.SelectColumns(Included,
        {"EmployeeID", "Facility-Abbrev", "Role", "BaseName", "EmploymentType", "Contracted FN Hours",
         "In Allocation", "In Records", "In Leave Balances", "Worker Membership"}),
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
    // Membership preparation selects the preferred roster role; retain its LINK Roles Role Group for downstream reporting.
    PreferredRole = Table.AddColumn(DisplayNames, "PreferredRole", each [Role], type text),
    RemovedNestedTables = Table.RemoveColumns(PreferredRole, {"Roles", "AllocationRoles"}),
    BufferedWorkers = Table.Buffer(RemovedNestedTables)
in
    BufferedWorkers,
AllocationRoleShares = #table(type table [EmployeeID=text, #"Facility-Abbrev"=text, Role=text, RoleCount=number, RoleShare=number],{})
in [Membership=ReconciledWorkerMembership_Prepare, Prepared=ReconciledWorkers_Prepare,
    Eligible=ReconciledWorkers_Eligible, Checks=ReconciledWorkers_CHECK],
EvaluateStaff = (Inputs as record) as record => let
ResourceContract_CHECK = let
    Contracts = ResourceContract,
    AvailabilityInput = #"EXTRACT Availability Staff List",
    AvailabilityResources = Table.SelectRows(Contracts, each [Source] = "Availability"),
    FallbackRows = Table.SelectRows(Contracts, each [Contract Source] = "Casual roster cap" and [CasualCapIssue] <> null),
    IsPositiveFiniteNumber = (value as any) as logical =>
        if not Value.Is(value, type number) then false
        else not Number.IsNaN(value) and Number.Abs(value) <> #infinity and value > 0,
    IsPositiveWholeNumber = (value as any) as logical =>
        if not IsPositiveFiniteNumber(value) then false else value = Number.RoundDown(value),
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
            or [PreferredRole] = null or [Role] <> [PreferredRole]
            or not [In Allocation] and not [In Records]
            or not List.Contains({"Allocation only", "Records only", "Both"}, [Worker Membership])))),
        StaffListCheckResult("Reconciliation is the only worker source", () =>
            Table.RowCount(Contracts) - Table.RowCount(AvailabilityResources)),
        StaffListCheckResult("Valid reported contract hours or roster cap basis", () => Table.RowCount(Table.SelectRows(Contracts, each
            if [Contracted FN Hours] = null then [Contract Source] <> "Casual roster cap"
            else not IsPositiveFiniteNumber([Contracted FN Hours]) or [Contract Source] <> "Worker record"))),
        StaffListCheckResult("Calculated resource contract caps", () => Table.RowCount(Table.SelectRows(AvailabilityResources, each
            if [Contract Source] = "Casual roster cap" then
                let
                    RawCap = [RosterShiftCap],
                    AppliedCap = [AppliedCasualRosterShiftCap],
                    ExpectedCap = if [CasualCapIssue] <> null then [Settings Max Availability]
                        else if IsPositiveWholeNumber(RawCap) then List.Min({RawCap, [Settings Max Availability]})
                        else null,
                    ExpectedBasis = if [CasualCapIssue] = null then "Casual fallback - RosterShiftCap"
                        else "Casual fallback - Settings maximum",
                    HasUsableAppliedCap = if not IsPositiveWholeNumber(AppliedCap) then false
                        else AppliedCap <= [Settings Max Availability]
                in [Contracted FN Hours] <> null or [Contracted Roster Hours] <> null
                    or [Contracted Shift Equivalent] <> null or [Contracted Shifts] <> null
                    or not HasUsableAppliedCap or ExpectedCap = null or AppliedCap <> ExpectedCap
                    or [Effective Shift Cap] <> ExpectedCap or [Limit Basis] <> ExpectedBasis
            else if not IsPositiveFiniteNumber([Contracted FN Hours]) then true
            else
                let
                    ExpectedRosterHours = [Contracted FN Hours] * [Roster Fortnights],
                    ExpectedEquivalent = ExpectedRosterHours / #"EXTRACT StaffList ShiftDuration",
                    ExpectedShifts = Number.RoundDown(ExpectedEquivalent),
                    ExpectedCap = if ExpectedShifts <= 0 then null
                        else List.Min({[Settings Max Availability], ExpectedShifts})
                in Number.Abs([Contracted Roster Hours] - ExpectedRosterHours) > 0.00000001
                    or Number.Abs([Contracted Shift Equivalent] - ExpectedEquivalent) > 0.00000001
                    or [Contracted Shifts] <> ExpectedShifts
                    or [AppliedCasualRosterShiftCap] <> null
                    or [Effective Shift Cap] <> ExpectedCap
                    or [Effective Shift Cap] = null)))
    },
    DiagnosticChecks = {
        [Check = "Workers using Settings roster-cap fallback", Status = if Table.IsEmpty(FallbackRows) then "Pass" else "Warning",
            Failures = Table.RowCount(FallbackRows), Details = "Inspect CasualRosterShiftCapExceptions in Capacity-ShiftAvailability for the worker-level reasons."]
    },
    Checks = BlockingChecks & DiagnosticChecks
in Table.Buffer(Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text])),
#"EXTRACT StaffList ShiftDuration" = 7.6,
StaffListCheckResult = (name as text, evaluate as function) as record =>
    let Result = try evaluate()
    in [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
        Failures = if Result[HasError] then null else Result[Value],
        Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null],
#"EXTRACT StaffList MaxAvailability" = 40,
#"EXTRACT StaffList RosterDates" = Table.FromColumns({List.Dates(#date(2026,10,1),29,#duration(1,0,0,0))},{"Date"}),
#"EXTRACT Availability Staff List" = Inputs[Staff],
ResourceContract = let
    Source = Masterlist_Base,
    RosterDates = #"EXTRACT StaffList RosterDates",
    RosterDateCount = Table.RowCount(RosterDates),
    RosterStart = List.Min(RosterDates[Date]),
    RosterEnd = List.Max(RosterDates[Date]),
    RosterFortnights = Number.From(RosterDateCount) / 14,
    ShiftDuration = #"EXTRACT StaffList ShiftDuration",
    SettingsMaximum = #"EXTRACT StaffList MaxAvailability",
    HasPositiveContractHours = (hours as nullable number) as logical =>
        if hours = null then false
        else not Number.IsNaN(hours) and Number.Abs(hours) <> #infinity and hours > 0,
    WithContractSource = Table.AddColumn(Source, "Contract Source", each
        if [Contracted FN Hours] = null then "Casual roster cap"
        else if [Contracted FN Hours] = 0 then "Casual roster cap"
        else "Worker record", type text),
    // Missing and zero hours mean no contract; invalid reported hours remain available to the blocking check.
    ResolvedContractHours = Table.TransformColumns(WithContractSource, {{"Contracted FN Hours", each
        if _ = null then null
        else if _ = 0 then null
        else _, type nullable number}}),
    WithRosterStart = Table.AddColumn(ResolvedContractHours, "Roster Start", each RosterStart, type date),
    WithRosterEnd = Table.AddColumn(WithRosterStart, "Roster End", each RosterEnd, type date),
    WithRosterFortnights = Table.AddColumn(WithRosterEnd, "Roster Fortnights", each RosterFortnights, type number),
    WithContractRosterHours = Table.AddColumn(WithRosterFortnights, "Contracted Roster Hours", each
        if HasPositiveContractHours([Contracted FN Hours]) then [Contracted FN Hours] * RosterFortnights
        else null, type nullable number),
    WithShiftEquivalent = Table.AddColumn(WithContractRosterHours, "Contracted Shift Equivalent", each
        if [Contracted Roster Hours] = null then null else [Contracted Roster Hours] / ShiftDuration, type nullable number),
    WithContractedShifts = Table.AddColumn(WithShiftEquivalent, "Contracted Shifts", each
        if [Contracted Shift Equivalent] = null then null else Int64.From(Number.RoundDown([Contracted Shift Equivalent])), Int64.Type),
    WithSettingsMaximum = Table.AddColumn(WithContractedShifts, "Settings Max Availability", each SettingsMaximum, Int64.Type),
    WithEffectiveCap = Table.AddColumn(WithSettingsMaximum, "Effective Shift Cap", each
        if [Contract Source] = "Casual roster cap" then [AppliedCasualRosterShiftCap]
        else if [Contracted Shifts] = null or [Contracted Shifts] <= 0 then null
        else List.Min({[Settings Max Availability], [Contracted Shifts]}), type nullable number),
    WithLimitBasis = Table.AddColumn(WithEffectiveCap, "Limit Basis", each
        // Effort recognises the Casual fallback prefix as having no exact contract-hours overage test.
        if [Contract Source] = "Casual roster cap" then
            if [CasualCapIssue] = null then "Casual fallback - RosterShiftCap"
            else "Casual fallback - Settings maximum"
        else if [Contracted Shifts] = null or [Contracted Shifts] <= 0 then "Invalid or missing contract"
        else if [Contracted Shifts] < [Settings Max Availability] then "Contracted shifts"
        else "Settings maximum", type text)
in Table.Buffer(WithLimitBasis),
Masterlist_Base = Inputs[Staff],
Masterlist = let
    Failures = Table.SelectRows(ResourceContract_CHECK, each [Status] = "Fail"),
    Checked = if Table.IsEmpty(Failures) then ResourceContract
        else error Error.Record("Resource contract validation", "Required resource contract checks failed.", Failures),
    Output = Table.SelectColumns(Checked,
        {"Name", "Role", "Resource", "Misalignment", "Source", "EmployeeID", "Facility-Abbrev", "EmploymentType", "PreferredRole",
         "In Allocation", "In Records", "Worker Membership", "Worker Record Status", "Worker Contract Issue", "Contract Source",
         "Contracted FN Hours", "Roster Start", "Roster End", "Roster Fortnights", "Contracted Roster Hours",
         "Contracted Shift Equivalent", "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"})
in Table.Sort(Output, {{"Resource", Order.Ascending}})
in [Published=Masterlist, Checks=ResourceContract_CHECK],
Fixtures = (Evaluate as function, EvaluateStaff as function) as table => let
    Start = #date(2026,10,1),
    Facilities = #table({"Title","FacilityShort","Facility-Abbrev"},{{"BD: Beaudesert","BD","BD"},{"TE: Temora","TE","TE"}}),
    RosterColumns = {"Location","Department","Employee Roster Name","Employment Type","Role","Employee Code","Date"},
    AvailabilityColumns = {"Site Name","Employee Name","PayrollCode","Department","DateTimeStart","DateTimeEnd","Comments","Type"},
    HRColumns = {"Employee Code","Location Description","Position Description","Employee Status Description","Start Date","Termination Date","Contracted FN Hours"},
    RawSheet = (records as list, columns as list) as table => Table.DemoteHeaders(Table.FromRecords(records,columns,MissingField.UseNull)),
    Roster = (id as text, optional role as nullable text, optional site as nullable text, optional name as nullable text, optional employment as nullable text) as record =>
        [Location=site ?? "BD: Beaudesert", Department="Care", #"Employee Roster Name"=name ?? ("Worker " & id),
            #"Employment Type"=employment ?? "Permanent Part Time", Role=role ?? "Assistant in Nursing AM", #"Employee Code"=id, Date=Start],
    RecordRow = (id as text, optional reason as nullable text, optional site as nullable text, optional name as nullable text) as record =>
        [#"Site Name"=site ?? "BD: Beaudesert", #"Employee Name"=name ?? ("Worker " & id), PayrollCode=id, Department="Care",
            DateTimeStart=#datetime(2026,10,1,0,0,0), DateTimeEnd=#datetime(2026,10,2,0,0,0), Comments=reason ?? "AVAIL", Type=reason ?? "AVAIL"],
    HR = (id as text, optional role as nullable text, optional site as nullable text, optional termination as nullable date, optional hours as nullable number) as record =>
        [#"Employee Code"=id, #"Location Description"=site ?? "BD: Beaudesert", #"Position Description"=role ?? "Assistant in Nursing",
            #"Employee Status Description"="Permanent Part Time", #"Start Date"=#date(2020,1,1), #"Termination Date"=termination, #"Contracted FN Hours"=hours ?? 76],
    Inputs = (roster as list, availability as list, hr as list, optional unit as nullable text) as record =>
        [Roster=RawSheet({[Location="BD: Beaudesert",Department="Care",#"Employee Roster Name"=null,#"Employment Type"="Agency",Role=null,#"Employee Code"=null,Date=Start]} & roster,RosterColumns),
            Availability=RawSheet(availability,AvailabilityColumns),HR=RawSheet(hr,HRColumns),
            Facilities=if unit="BD-FULL" then #table({"Title","FacilityShort","Facility-Abbrev"},{{"BD: Beaudesert","BD","BD-FULL"}}) else Facilities,Unit=unit ?? "BD"],
    Check = (name as text, run as function, predicate as function) as record => let
        Attempt=try run(), Verdict=if Attempt[HasError] then [HasError=false,Value=false] else try predicate(Attempt[Value]),
        Passed=not Attempt[HasError] and not Verdict[HasError] and Verdict[Value]=true,
        Details=if Attempt[HasError] then (try Attempt[Error][Message] otherwise "Evaluation failed")
            else if Verdict[HasError] then (try Verdict[Error][Message] otherwise "Assertion failed") else if Passed then null else "Unexpected result"
    in [Case=name,Passed=Passed,Details=Details],
    Membership = (inputs as record) as table => Table.Buffer(Evaluate(inputs)[Membership]),
    ErrorBlocked = (inputs as record) as logical => (try Table.RowCount(Membership(inputs)))[HasError],
    Cases = {
        Check("roster-only worker",()=>Membership(Inputs({Roster("1001")},{},{HR("1001")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=true and _{0}[In Records]=false and _{0}[EmployeeID]="1001"),
        Check("availability-only worker",()=>Membership(Inputs({},{RecordRow("1002")},{HR("1002")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=false and _{0}[In Records]=true),
        Check("both sources one worker",()=>Membership(Inputs({Roster("1003")},{RecordRow("1003")},{HR("1003")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=true and _{0}[In Records]=true),
        Check("HR-only and hypothetical Balance-only IDs excluded",()=>Membership(Inputs({Roster("1001")},{},{HR("1001"),HR("HR-ONLY"),HR("BALANCE-ONLY")})),each Table.RowCount(_)=1 and _{0}[EmployeeID]="1001"),
        Check("wholly UNAVAIL worker retained",()=>Evaluate(Inputs({},{RecordRow("1004","UNAVAIL")},{HR("1004")}))[Eligible],each Table.RowCount(_)=1 and _{0}[EmployeeID]="1004"),
        Check("roster facility retained against HR home",()=>Membership(Inputs({Roster("1005")},{},{HR("1005",null,"TE: Temora")})),each Table.RowCount(_)=1 and _{0}[#"Facility-Abbrev"]="BD"),
        Check("availability wrong HR home excluded",()=>Membership(Inputs({},{RecordRow("wrong-home")},{HR("wrong-home",null,"TE: Temora")})),each Table.IsEmpty(_)),
        Check("source alias maps complete canonical facility key",()=>Evaluate(Inputs({Roster("1006")},{RecordRow("1006")},{HR("1006")},"BD-FULL"))[Prepared],each Table.RowCount(_)=1 and _{0}[#"Facility-Abbrev"]="BD-FULL"),
        Check("complete text employee ID preserved",()=>Membership(Inputs({Roster("001234-X")},{RecordRow("001234-X")},{HR("001234-X")})),each Table.RowCount(_)=1 and _{0}[EmployeeID]="001234-X"),
        Check("agency roster excluded",()=>Membership(Inputs({Roster("AGENCY",null,null,null,"Agency ")},{},{})),each Table.IsEmpty(_)),
        Check("HF and extraction-error records excluded",()=>Membership(Inputs({},{RecordRow("HF"),RecordRow("999999")},{HR("HF"),HR("999999")})),each Table.IsEmpty(_)),
        Check("termination before roster start excluded",()=>Membership(Inputs({Roster("1007")},{},{HR("1007",null,null,#date(2026,9,30))})),each Table.IsEmpty(_)),
        Check("termination on roster start retained",()=>Membership(Inputs({Roster("1008")},{},{HR("1008",null,null,Start)})),each Table.RowCount(_)=1),
        Check("conflicting eligible HR records block",()=>ErrorBlocked(Inputs({Roster("1009")},{},{HR("1009"),HR("1009","Enrolled Nurse")})),each _=true),
        Check("identical duplicate HR rows allowed",()=>Membership(Inputs({Roster("1010")},{},{HR("1010"),HR("1010")})),each Table.RowCount(_)=1),
        Check("preferred roster role aligned with HR",()=>Membership(Inputs({Roster("1011","Assistant in Nursing AM"),Roster("1011","Enrolled Nurse AM")},{},{HR("1011","Enrolled Nurse")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("roster role before HR fallback",()=>Membership(Inputs({Roster("1012","Enrolled Nurse AM")},{},{HR("1012","Assistant in Nursing")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("availability role from HR",()=>Membership(Inputs({},{RecordRow("1013")},{HR("1013","Enrolled Nurse")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("availability with missing HR excluded",()=>Membership(Inputs({},{RecordRow("missing-HR")},{})),each Table.IsEmpty(_)),
        Check("no Balance role fallback",()=>ErrorBlocked(Inputs({Record.TransformFields(Roster("1014"),{{"Role",each null}})},{},{})),each _=true),
        Check("leave provenance compatibility false",()=>Membership(Inputs({Roster("1015")},{RecordRow("1015")},{HR("1015")})),each Table.RowCount(_)=1 and List.AllTrue(List.Transform(_[In Leave Balances],each _=false))),
        Check("production eligibility and checks",()=>Evaluate(Inputs({Roster("1016")},{RecordRow("1016")},{HR("1016")})),each Table.RowCount([Eligible])=1 and List.AllTrue(List.Transform([Checks][Status],each _="Pass"))),
        Check("reported contract retained",()=>Evaluate(Inputs({Roster("1017")},{},{HR("1017",null,null,null,57)}))[Eligible],each Table.RowCount(_)=1 and _{0}[Contracted FN Hours]=57),
        Check("roster without HR retains null contract",()=>Evaluate(Inputs({Roster("no-HR")},{},{}))[Eligible],each Table.RowCount(_)=1 and _{0}[Contracted FN Hours]=null)
    },
    StaffRow = (id as text, hours as nullable number, cap as nullable number, applied as nullable number, issue as nullable text) as record =>
        [Name="Staff "&id,Role="AIN",Resource=1,Misalignment=0,Source="Availability",EmployeeID=id,#"Facility-Abbrev"="BD",EmploymentType=if hours=null then "Casual" else "Permanent Part Time",PreferredRole="AIN",
            #"In Allocation"=true,#"In Records"=true,#"Worker Membership"="Both",#"Worker Record Status"="Found",#"Worker Contract Issue"=null,#"Contracted FN Hours"=hours,
            RosterShiftCap=cap,AppliedCasualRosterShiftCap=applied,CasualCapIssue=issue],
    StaffCases={
        Check("staff reported contract proration and cap unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("contract",76,null,null,null)})])[Published]{0},each Number.Abs([Contracted Roster Hours]-76*29/14)<0.00000001 and [Contracted Shifts]=20 and [Effective Shift Cap]=20 and [Contract Source]="Worker record"),
        Check("staff casual roster cap unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("casual",null,13,13,null)})])[Published]{0},each [Effective Shift Cap]=13 and [Contracted FN Hours]=null and [Contracted Roster Hours]=null),
        Check("staff casual Settings fallback unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("fallback",null,null,40,"Missing roster cap")})])[Published]{0},each [Effective Shift Cap]=40 and [Limit Basis]="Casual fallback - Settings maximum")
    }
in Table.FromRecords(Cases & StaffCases)
in Fixtures(Evaluate,EvaluateStaff);
