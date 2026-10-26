// Query: IMPORT Published Roster
// Purpose: Land the existing unit's published roster unchanged for worker membership.
// Inputs: Roster Data in 1. Input/Allocation/Published Roster.xlsx.
shared #"IMPORT Published Roster" = let
    SourcePath = #"FilePath - 1Input" & "\Allocation\Published Roster.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    RosterSheet = Navigation{[Item = "Roster Data", Kind = "Sheet"]}[Data]
in
    RosterSheet;

// Query: PublishedRoster_Prepare
// Purpose: Preserve the existing roster location, department and base-role rules before worker selection.
// Notes: Dates retain the original all-analysis-location roster-start basis, before agency or worker filtering.
shared PublishedRoster_Prepare = let
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
    BaseRoles;

// Query: RosterWorkers_Prepare
// Purpose: Retain identified non-agency roster workers with the existing source role and employment attributes.
// Output: Distinct roster candidates before preferred-role selection, keyed by complete employee ID and source facility.
shared RosterWorkers_Prepare = let
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
    Table.Buffer(WithRosterMembership);

// Query: WorkerEmployment_Prepare
// Purpose: Prepare worker-report employment facts while preserving complete employee IDs and source contract errors.
// Inputs: Unchanged Sheet1 from IMPORT Worker's Report.
// Notes: All HR home locations remain available for the established availability-location match.
shared WorkerEmployment_Prepare = let
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
    Table.Buffer(EmploymentColumns);

// Query: WorkerEmployment_Selected
// Purpose: Preserve deterministic employee-level HR selection and the existing material employment-ambiguity count.
// Notes: A termination on roster start remains eligible; HR presence alone never creates worker membership.
shared WorkerEmployment_Selected = let
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
    Table.Buffer(ExpandedWorker);

// Query: AvailabilityMembership_Prepare
// Purpose: Preserve availability worker membership using full employee IDs and the existing HR-home-location match.
// Notes: Presence does not depend on event type, dates or positive available hours; system accounts retain their existing exclusion.
shared AvailabilityMembership_Prepare = let
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
    Table.Buffer(WithAvailabilityMembership);

// Query: ReconciledWorkerMembership_Prepare
// Purpose: Supply the existing worker interface from roster, availability and HR without reading Balance or Worker Reconciliation.
// Output: One preferred source-role worker per employee/facility; legacy In Leave Balances is always false.
// Notes: In Records means availability membership; In HR Records identifies employment enrichment only.
shared ReconciledWorkerMembership_Prepare = let
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
    Table.Buffer(Output);
