section FacilityMappingTests;
shared FacilityMappingTests = () as table => let
    Facilities = #table({"Title","Facility-Abbrev","FacilityShort"},{
        {"BD: Beaudesert","BD","BD"}, {"JE: Jessie Hunt","JH-RY","JE"},
        {"NR: Narrabri- Robert Young","JH-RY","NR"}, {"TE: Temora","TE","TE"},
        {"TE: Temora Greenstone","TE","TE"}, {"TE: Temora","TE","TE"}}),
    Workers = #table({"EmployeeID","Facility-Abbrev","Role"}, {
        {"JE1","JE","AIN"},{"NR1","NR","AIN"},{"CAN1","JH-RY","RN"},
        {"BD1","BD","EN"},{"TE1","TE","RN"},{"XX1","XX","AIN"},{"NULL1",null,"AIN"}}),
    EventSource = #table({"Column1","Column2","Column3","Column4","Column5","Column6","Column7"},{
        {"Site Name","Employee Name","PayrollCode","Type","DateTimeStart","DateTimeEnd","Comments"},
        {"NR: Narrabri- Robert Young","Worker NR","NR1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"JE: Jessie Hunt","Worker JE","JE1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"TE: Temora","Worker TE","TE1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"BD: Beaudesert","Worker BD","BD1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null}}),
    AllocationSource = #table({"Location","Code","Role","Hours","Name"},{
        {"NR: Narrabri- Robert Young","CROSS1","AIN",2,"Worker Cross"},
        {"JE: Jessie Hunt","CROSS1","AIN",3,"Worker Cross"},
        {"TE: Temora","TE1","RN",4,"Worker TE"},{"BD: Beaudesert","BD1","EN",5,"Worker BD"}}),
    Evaluate = (reference as table, selectedUnit as text, sourceWorkers as table, enabled as logical) as record => let
        Unit = selectedUnit,
        #"IMPORT Facilities List" = #table({"Id","Items"},{{"1987eadb-ad2e-491a-a927-e5585667d4c5",reference}}),
        ReconciledWorkerMembership_Prepare = sourceWorkers,
        #"IMPORT Availability Leave Source" = EventSource,
        #"IMPORT AllocationExtracted" = AllocationSource,
        FacilityAnalysisTABLE = if enabled then #table({"Unit"},{{selectedUnit}}) else #table({"Unit"},{}),
        AvailabilityRoleGroups = {"AIN","EN","RN"},#"AvailabilityText" = (value as any) as nullable text =>
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
#"LINK Facilities" = let
    FacilityItems = #"IMPORT Facilities List"{[Id="1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items]
in
    FacilityItems,
#"FacilityMappings_Prepare" = let
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
#"ReconciledWorkers_Facility_Prepare" = let
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
#"AvailabilityLeaveSource_Prepare" = let
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
#"AllocationRoleHours_Prepare" = let
    ApprovedRoleNames = AvailabilityRoleGroups,
    ApprovedRoles = Table.SelectRows(#"IMPORT AllocationExtracted", each List.Contains(ApprovedRoleNames, [Role])),
    NormalizedFacilityKey = Table.AddColumn(ApprovedRoles, "FacilityKey", each
        if [Location] = null then null else Text.Upper(Text.Trim(Text.From([Location]))), type nullable text),
    JoinedFacilities = Table.NestedJoin(NormalizedFacilityKey, {"FacilityKey"}, FacilityMappings_Prepare,
        {"FacilityKey"}, "FacilityLookup", JoinKind.LeftOuter),
    CountedMappings = Table.AddColumn(JoinedFacilities, "FacilityMatches", each
        Table.RowCount([FacilityLookup]), Int64.Type),
    InvalidMappings = Table.SelectRows(CountedMappings, each [FacilityMatches] <> 1),
    ValidatedMappings = if Table.IsEmpty(InvalidMappings) then CountedMappings
        else error Error.Record("Allocation role facility mapping",
            "Each allocation Location must match exactly one canonical Facilities mapping.",
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
    BufferedResult
    in [Raw=#"LINK Facilities",Mapping=FacilityMappings_Prepare,Workers=ReconciledWorkers_Facility_Prepare,
        Events=AvailabilityLeaveSource_Prepare,Allocations=AllocationRoleHours_Prepare],
    Baseline = Evaluate(Facilities,"JH-RY",Workers,true),
    Worker = (id as text) as record => Table.SelectRows(Baseline[Workers],each [EmployeeID]=id){0},
    Check = (name as text, run as function) as record => let Attempt=try run(),
        Passed=if Attempt[HasError] then false else Attempt[Value]=true
    in [Case=name,Passed=Passed,Details=if Attempt[HasError] then Attempt[Error][Message] else if Passed then null else "Unexpected result"],
    LowerCaseWorkers = #table({"EmployeeID","Facility-Abbrev","Role"},{{"lower"," nr ","AIN"}}),
    ChangedReference = Table.ReplaceValue(Facilities,"NR","ZZ",Replacer.ReplaceValue,{"FacilityShort"}),
    ChangedWorkers = #table({"EmployeeID","Facility-Abbrev","Role"},{{"dynamic","ZZ","AIN"}}),
    ConflictingReference = Table.Combine({Facilities,#table({"Title","Facility-Abbrev","FacilityShort"},{{"Conflicting site","BD","NR"}})}),
    MissingCanonical = Table.ReplaceValue(Facilities,"JH-RY",null,Replacer.ReplaceValue,{"Facility-Abbrev"}),
    LegacyReference = Table.RenameColumns(Facilities,{{"Facility-Abbrev","field_1"}}),
    MatchedEvents = Table.NestedJoin(Baseline[Events],{"Payroll Code","Facility-Abbrev"},Baseline[Workers],{"EmployeeID","Facility-Abbrev"},"Worker",JoinKind.LeftOuter),
    Cases = {
        Check("NR maps to JH-RY",()=> Worker("NR1")[#"Facility-Abbrev"]="JH-RY"),
        Check("JE maps to JH-RY",()=> Worker("JE1")[#"Facility-Abbrev"]="JH-RY"),
        Check("Already canonical code retained",()=> Worker("CAN1")[#"Facility-Abbrev"]="JH-RY"),
        Check("BD preserved",()=> Worker("BD1")[#"Facility-Abbrev"]="BD"),
        Check("TE preserved",()=> Worker("TE1")[#"Facility-Abbrev"]="TE"),
        Check("Original NR code remains visible",()=> Worker("NR1")[SourceFacility]="NR"),
        Check("Unmapped value remains diagnostic",()=> Worker("XX1")[#"Facility-Abbrev"]="XX" and Worker("XX1")[FacilityMappingStatus]="Unmapped facility: inspect LINK Facilities"),
        Check("Null facility remains diagnostic",()=> Worker("NULL1")[#"Facility-Abbrev"]=null and Worker("NULL1")[FacilityMappingStatus]="Missing facility"),
        Check("Worker row grain preserved",()=> Table.RowCount(Baseline[Workers])=Table.RowCount(Workers)),
        Check("Duplicate TE mappings do not multiply lookup",()=> Table.RowCount(Baseline[Mapping])=10),
        Check("Lookup keys unique",()=> Table.RowCount(Baseline[Mapping])=Table.RowCount(Table.Distinct(Baseline[Mapping],{"FacilityKey"}))),
        Check("Case and whitespace normalized",()=> Evaluate(Facilities,"JH-RY",LowerCaseWorkers,true)[Workers]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("Reference-driven alias without code edits",()=> Evaluate(ChangedReference,"JH-RY",ChangedWorkers,true)[Workers]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("NR and JE availability included",()=> Table.RowCount(Baseline[Events])=2 and List.AllTrue(List.Transform(Baseline[Events][#"Facility-Abbrev"],each _="JH-RY"))),
        Check("NR and JE source identities join to workers",()=> List.AllTrue(List.Transform(MatchedEvents[Worker],each Table.RowCount(_)=1))),
        Check("Both sites allocation grouped once",()=> Table.RowCount(Baseline[Allocations])=1 and Baseline[Allocations]{0}[AllocatedHours]=5 and Baseline[Allocations]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("BD unit scope preserved",()=> Table.RowCount(Evaluate(Facilities,"BD",Workers,true)[Events])=1 and Evaluate(Facilities,"BD",Workers,true)[Allocations]{0}[AllocatedHours]=5),
        Check("TE duplicates do not multiply events or hours",()=> Table.RowCount(Evaluate(Facilities,"TE",Workers,true)[Events])=1 and Evaluate(Facilities,"TE",Workers,true)[Allocations]{0}[AllocatedHours]=4),
        Check("Disabled analysis remains excluded",()=> Table.IsEmpty(Evaluate(Facilities,"JH-RY",Workers,false)[Events]) and Table.IsEmpty(Evaluate(Facilities,"JH-RY",Workers,false)[Allocations])),
        Check("Legacy field_1 projection supported",()=> Evaluate(LegacyReference,"JH-RY",Workers,true)[Workers]{1}[#"Facility-Abbrev"]="JH-RY"),
        Check("LINK lands full list unchanged",()=> Table.ColumnNames(Evaluate(Table.AddColumn(Facilities,"Extra",each 1),"JH-RY",Workers,true)[Raw])=Table.ColumnNames(Table.AddColumn(Facilities,"Extra",each 1)) and Table.RowCount(Baseline[Raw])=6),
        Check("Conflicting alias blocked",()=> (try Table.RowCount(Evaluate(ConflictingReference,"JH-RY",Workers,true)[Mapping]))[HasError]),
        Check("Missing FacilityShort is explicit error",()=> (try Table.RowCount(Evaluate(Table.RemoveColumns(Facilities,{"FacilityShort"}),"JH-RY",Workers,true)[Mapping]))[HasError]),
        Check("Missing canonical facility blocked",()=> (try Table.RowCount(Evaluate(MissingCanonical,"JH-RY",Workers,true)[Mapping]))[HasError])
    }
in Table.FromRecords(Cases);