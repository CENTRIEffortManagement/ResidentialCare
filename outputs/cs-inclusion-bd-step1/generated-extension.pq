section CSRoleDemandNative;
shared CSRoleDemandNativeTests = () as table => let
Evaluate = (Inputs as record) as record => let
RoleDemand_CHECK = let
    CheckStage = (Name as text, Work as function) as record =>
        let Attempt = try Work()
        in [Check = Name, Severity = if Attempt[HasError] then "Error" else "Pass",
            Actual = if Attempt[HasError] then null else Attempt[Value],
            Message = if Attempt[HasError] then Attempt[Error][Message] else "Validated"],
    Stages = #"MW Check Table"({
        CheckStage("Unique role-analysis controls", () => Table.RowCount(#"RoleDemand Role Analysis Prepare")),
        CheckStage("Unambiguous care role-group mappings", () => Table.RowCount(#"RoleDemand Care Role Mapping Prepare")),
        CheckStage("Established source fortnight", () => Table.RowCount(#"RoleDemand Source Weeks Prepare")),
        CheckStage("Existing care publication and source-cell agreement", () => Table.RowCount(#"RoleDemand Care Cells Prepare")),
        CheckStage("CS roster-role and source rows", () => Table.RowCount(
            Table.SelectRows(#"RoleDemand CS Rows Prepare", each [LocationMatchCount] = 1))),
        CheckStage("CS source-hour conservation and proven zero cells", () => Table.RowCount(#"RoleDemand CS Pattern Prepare")),
        CheckStage("Complete combined role-demand handover", () => Table.RowCount(#"RoleDemand Pattern Prepare"))
    }),
    CSControlAttempt = try Table.SelectRows(#"RoleDemand Role Analysis Prepare", each [Roles] = "CS"),
    // Invalid or missing settings remain stage errors; a valid disabled setting must not look like successful CS inclusion.
    CSInclusion = if CSControlAttempt[HasError] then #"MW Check Table"({})
        else if Table.RowCount(CSControlAttempt[Value]) <> 1 then #"MW Check Table"({})
        else
            let Enabled = CSControlAttempt[Value]{0}[Effort Management Analysis] = true
            in #"MW Check Table"({
                [Check = "CS effort-analysis inclusion", Severity = if Enabled then "Pass" else "Warning",
                    Role = "CS", Actual = if Enabled then 1 else 0, Expected = 1,
                    Message = if Enabled then "CS is enabled for effort management."
                        else "CS preparation was skipped and no CS demand was published because Role Analysis does not enable Effort Management Analysis for CS."]
            }),
    CareAttempt = try MinuteWorkersFTE_CHECK,
    // These notices describe expected rows outside the configured source scope.
    // They remain in the original care diagnostics; all other warnings are preserved here.
    CareWarnings = if CareAttempt[HasError] then #"MW Check Table"({}) else
        Table.SelectRows(CareAttempt[Value], each [Severity] = "Warning"
            and [Check] <> "Master Roster Location is outside configured Units")
in
    Table.Combine({Stages, CSInclusion, CareWarnings}),
#"MW Check Table" = (Rows as list) as table =>
let
    EmptyCheckTable = #table(
        type table [
            Check = nullable text,
            Severity = nullable text,
            Facility = nullable text,
            MinuteCategory = nullable text,
            Role = nullable text,
            Actual = nullable number,
            Expected = nullable number,
            Message = nullable text
        ],
        {}
    ),
    Result =
        if List.Count(Rows) = 0 then
            EmptyCheckTable
        else
            // Check records share a fixed interface. Missing optional context
            // fields are intentionally represented as null in the check output.
            Table.FromRecords(
                Rows,
                type table [
                    Check = nullable text,
                    Severity = nullable text,
                    Facility = nullable text,
                    MinuteCategory = nullable text,
                    Role = nullable text,
                    Actual = nullable number,
                    Expected = nullable number,
                    Message = nullable text
                ],
                MissingField.UseNull
            )
in
    Result,
MinuteWorkersFTE_CHECK = Inputs[CareChecks],
#"RoleDemand Role Analysis Prepare" = let
    Selected = Table.SelectColumns(#"EXTRACT RoleAnalysis", {"Roles", "Effort Management Analysis"}),
    Trimmed = Table.TransformColumns(Selected, {{"Roles", each if _ = null then null else Text.Trim(_), type nullable text}}),
    Invalid = Table.SelectRows(Trimmed, each [Roles] = null or [Roles] = ""
        or ([Effort Management Analysis] <> null and not Value.Is([Effort Management Analysis], type logical))),
    Counts = Table.Group(Trimmed, {"Roles"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    Duplicates = Table.SelectRows(Counts, each [Count] <> 1),
    Validated = if not Table.IsEmpty(Invalid) then error "Role demand requires nonblank role groups and logical effort flags."
        else if not Table.IsEmpty(Duplicates) then error "Role demand requires one analysis row per role group."
        else Trimmed
in
    Table.Buffer(Validated),
#"EXTRACT RoleAnalysis" = Inputs[Analysis],
#"RoleDemand Pattern Prepare" = let
    Columns = {"Facility", "Role", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift", "DemandHRS", "ProductiveHRS", "DemandBasis"},
    Care = #"RoleDemand Care Cells Prepare",
    CS = #"RoleDemand CS Pattern Prepare",
    Combined = Table.SelectColumns(Table.Combine({Care, CS}), Columns),
    CareGroups = Table.SelectColumns(Table.SelectRows(#"RoleDemand Care Role Mapping Prepare", each [Effort Management Analysis] = true), {"Role"}),
    Groups = Table.Distinct(Table.Combine({CareGroups,
        if Table.IsEmpty(#"RoleDemand CS Role Mapping Prepare") then #table(type table [Role = text], {}) else #table(type table [Role = text], {{"CS"}})})),
    Facilities = Table.Distinct(Table.SelectColumns(#"RoleDemand Source Weeks Prepare", {"Facility"})),
    Expected = Table.ExpandTableColumn(Table.AddColumn(Facilities, "Groups", each Groups), "Groups", {"Role"}),
    Counts = Table.Group(Combined, {"Facility", "Role"}, {{"CellCount", each Table.RowCount(_), Int64.Type}}),
    Coverage = Table.NestedJoin(Expected, {"Facility", "Role"}, Counts, {"Facility", "Role"}, "Actual", JoinKind.LeftOuter),
    Missing = Table.SelectRows(Coverage, each Table.RowCount([Actual]) <> 1 or [Actual]{0}[CellCount] <> 42),
    DuplicateKeys = Table.RowCount(Combined) <> Table.RowCount(Table.Distinct(Table.SelectColumns(Combined, {"Facility", "Role", "FortnightDayIndex", "Shift"}))),
    Invalid = Table.SelectRows(Combined, each not #"RoleDemand Finite Number"([DemandHRS]) or [DemandHRS] < 0
        or (if [DemandBasis] = "MasterRoster" then [Role] <> "CS" or [ProductiveHRS] <> null
            else [DemandBasis] <> "CareMinutes" or [Role] = "CS"
                or not #"RoleDemand Finite Number"([ProductiveHRS]) or [ProductiveHRS] < 0)),
    Validated = if Table.IsEmpty(Expected) or not Table.IsEmpty(Missing) or DuplicateKeys
            or Table.RowCount(Combined) <> Table.RowCount(Expected) * 42 or not Table.IsEmpty(Invalid) then
            error "Combined role demand requires 42 unique valid cells for every enabled facility/role."
        else Combined,
    Typed = Table.TransformColumnTypes(Validated, {{"Facility", type text}, {"Role", type text}, {"Week No", Int64.Type},
        {"FortnightWeek", Int64.Type}, {"DayOfWeek", Int64.Type}, {"FortnightDayIndex", Int64.Type}, {"Shift", type text},
        {"DemandHRS", type number}, {"ProductiveHRS", type nullable number}, {"DemandBasis", type text}})
in
    Table.Buffer(Table.Sort(Typed, {{"Facility", Order.Ascending}, {"Role", Order.Ascending},
        {"FortnightDayIndex", Order.Ascending}, {"Shift", Order.Ascending}})),
#"RoleDemand Finite Number" = (Value as any) as logical =>
    Value.Is(Value, type number) and not Number.IsNaN(Value) and Number.Abs(Value) <> #infinity,
#"RoleDemand Source Weeks Prepare" = let
    Selected = Table.Distinct(Table.SelectColumns(#"MW Historical Weeks",
        {"Facility", "Week No", "FortnightWeek", "HistoricalWeeksInPeriod"})),
    Invalid = Table.SelectRows(Selected, each [Facility] = null or Text.Trim([Facility]) = ""
        or not #"RoleDemand Finite Number"([Week No]) or Number.RoundDown([Week No]) <> [Week No]
        or not List.Contains({1, 2}, [FortnightWeek]) or [HistoricalWeeksInPeriod] <> 2),
    Coverage = Table.Group(Selected, {"Facility"}, {{"Valid", each
        let Weeks = Table.Sort(_, {{"FortnightWeek", Order.Ascending}})
        in Table.RowCount(Weeks) = 2 and List.Sort(Weeks[FortnightWeek]) = {1, 2}
            and Weeks{0}[Week No] < Weeks{1}[Week No], type logical}}),
    Validated = if Table.IsEmpty(Selected) or not Table.IsEmpty(Invalid)
            or not Table.IsEmpty(Table.SelectRows(Coverage, each not [Valid])) then
            error "Role demand requires exactly two ordered source weeks per modelled facility."
        else Selected
in
    Table.Buffer(Table.RemoveColumns(Validated, {"HistoricalWeeksInPeriod"})),
#"MW Historical Weeks" = let
    // Source locations have already mapped to canonical Facility. Distinct weeks
    // are therefore evaluated at the same three-facility grain as targets.
    Source = Table.SelectColumns(#"Master Prepare", {"Facility", "Week No"}),
    ValidWeeks = Table.Distinct(Table.SelectRows(Source, each [Week No] <> null)),
    FacilityWeeks = Table.Group(ValidWeeks, {"Facility"}, {
        {"Weeks", each Table.AddIndexColumn(Table.Sort(Table.SelectColumns(_, {"Week No"}), {{"Week No", Order.Ascending}}),
            "FortnightWeek", 1, 1, Int64.Type), type table},
        {"HistoricalWeeksInPeriod", each Table.RowCount(_), Int64.Type}
    }),
    Result = Table.ExpandTableColumn(FacilityWeeks, "Weeks", {"Week No", "FortnightWeek"}, {"Week No", "FortnightWeek"})
in
    Result,
#"Master Prepare" = let
    Source = #"Master Source Prepare",
    #"Added Source Location Key" = Table.AddColumn(
        Source,
        "SourceLocationKey",
        each if [Location] = null then null else Text.Upper(Text.Trim(Text.From([Location]))),
        type nullable text
    ),
    #"Joined Facility Mapping" = Table.NestedJoin(
        #"Added Source Location Key",
        {"SourceLocationKey"},
        #"MW Facility Mapping Prepare",
        {"SourceLocationKey"},
        "FacilityMapping",
        JoinKind.LeftOuter
    ),
    #"Added Facility Mapping Count" = Table.AddColumn(
        #"Joined Facility Mapping",
        "FacilityMappingCount",
        each Table.RowCount([FacilityMapping]),
        Int64.Type
    ),
    #"Expanded Canonical Facility" = Table.ExpandTableColumn(
        #"Added Facility Mapping Count",
        "FacilityMapping",
        {"Facility"},
        {"Facility"}
    ),
    // EXTRACT Units is the location inclusion list: unlisted locations are
    // excluded with a warning. Ambiguous configured mappings remain errors in
    // MW Input Check; no source row is assigned by guesswork.
    #"Retained Unique Facility Mapping" = Table.SelectRows(
        #"Expanded Canonical Facility",
        each [FacilityMappingCount] = 1
    ),
    #"Selected Historical Columns" = Table.SelectColumns(
        #"Retained Unique Facility Mapping",
        {"Location", "Facility", "Role", "Employee Code", "Week No", "Week Day", "Start Time", "End Time", "Shift Net Length"}
    ),
    // Keep the established downstream column name while taking hours from the net shift field.
    #"Named Net Roster Hours" = Table.RenameColumns(
        #"Selected Historical Columns",
        {{"Shift Net Length", "Roster Hours"}}
    ),
    #"Renamed Original Role" = Table.RenameColumns(
        #"Named Net Roster Hours",
        {{"Role", "OriginalRole"}}
    ),
    #"Added Roster Role Key" = Table.AddColumn(
        #"Renamed Original Role",
        "RosterRoleKey",
        each if [OriginalRole] = null then null else Text.Upper(Text.Trim([OriginalRole])),
        type nullable text
    ),
    #"Merged Explicit Role Mapping" = Table.NestedJoin(
        #"Added Roster Role Key",
        {"RosterRoleKey"},
        #"MW Roster Role Mapping Prepare",
        {"RosterRoleKey"},
        "RoleMapping",
        JoinKind.LeftOuter
    ),
    #"Expanded Explicit Role Mapping" = Table.ExpandTableColumn(
        #"Merged Explicit Role Mapping",
        "RoleMapping",
        {"DC Role", "RoleKey", "DC Category", "MinuteCategory", "Direct Care %"},
        {"DC Role", "RoleKey", "DC Category", "MinuteCategory", "Direct Care %"}
    ),
    // This table is an inclusion list. NA and unmatched roster roles remain
    // outside MinuteWorker allocation without being treated as input errors.
    #"Filtered To Direct Care Roles" = Table.SelectRows(
        #"Expanded Explicit Role Mapping",
        each List.Contains({"RN", "OTHER"}, [DC Category])
    ),
    #"Added Analytical Role" = Table.AddColumn(
        #"Filtered To Direct Care Roles",
        "Role",
        each [RoleKey],
        type nullable text
    ),
    #"Assigned Shift" = Table.AddColumn(
        #"Added Analytical Role",
        "Shift",
        each
            if [Start Time] <= StartAM then "NS"
            else if [Start Time] <= StartPM then "AM"
            else if [Start Time] <= StartNS then "PM"
            else if [Start Time] <= #time(23, 59, 59) then "NS"
            else "ERROR",
        type text
    ),
    #"Merged Shift Index" = Table.NestedJoin(
        #"Assigned Shift",
        {"Shift"},
        #"ShftEnd Prepare",
        {"Shift"},
        "ShftEnd Prepare",
        JoinKind.LeftOuter
    ),
    #"Expanded ShftEnd Prepare" = Table.ExpandTableColumn(
        #"Merged Shift Index",
        "ShftEnd Prepare",
        {"Index"},
        {"ShiftIndex"}
    )
in
    #"Expanded ShftEnd Prepare",
#"ShftEnd Prepare" = let
    Source = #"INPUT SHiftEnd",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type)
in
    #"Added Index",
#"INPUT SHiftEnd" = Inputs[ShiftEnds],
StartNS = let
    Source = #"ShftEnd Prepare",
    End = Source{1}[End]
in
    End,
StartPM = let
    Source = #"ShftEnd Prepare",
    End = Source{0}[End]
in
    End,
StartAM = let
    Source = #"ShftEnd Prepare",
    End = Source{2}[End]
in
    End,
#"MW Roster Role Mapping Prepare" = let
    Source = #"INPUT MinuteWorkers",
    #"Trimmed Mapping Fields" = Table.TransformColumns(
        Source,
        {
            {"Roster Roles", each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
            {"DC Role", each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
            {"DC Category", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type nullable text}
        }
    ),
    // Preserve whether the workbook supplied a true numeric value, then make
    // invalid values null so downstream multiplication cannot fail before the
    // input check reports the configuration error.
    #"Flagged Numeric Direct Care" = Table.AddColumn(
        #"Trimmed Mapping Fields",
        "DirectCareInputIsNumeric",
        each Value.Is([#"Direct Care %"], type number),
        type logical
    ),
    #"Normalised Direct Care Percentage" = Table.TransformColumns(
        #"Flagged Numeric Direct Care",
        {{"Direct Care %", each if Value.Is(_, type number) then Number.From(_) else null, type nullable number}}
    ),
    #"Added Roster Role Key" = Table.AddColumn(
        #"Normalised Direct Care Percentage",
        "RosterRoleKey",
        each if [Roster Roles] = null then null else Text.Upper([Roster Roles]),
        type nullable text
    ),
    #"Added DC Role Key" = Table.AddColumn(
        #"Added Roster Role Key",
        "RoleKey",
        each if [DC Role] = null then null else Text.Upper([DC Role]),
        type nullable text
    ),
    #"Added Minute Category" = Table.AddColumn(
        #"Added DC Role Key",
        "MinuteCategory",
        each if [DC Category] = "RN" then "RN" else if [DC Category] = "OTHER" then "OTHERS" else null,
        type nullable text
    )
in
    #"Added Minute Category",
#"INPUT MinuteWorkers" = Inputs[Roles],
#"MW Facility Mapping Prepare" = let
    Source = Table.SelectColumns(#"EXTRACT Units", {"Location", "Unit"}),
    #"Normalised Mapping Values" = Table.TransformColumns(
        Source,
        {
            {"Location", each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
            {
                "Unit",
                each
                    if _ = null then
                        null
                    else
                        let CanonicalFacility = Text.Upper(Text.Trim(Text.From(_)))
                        in if CanonicalFacility = "" then null else CanonicalFacility,
                type nullable text
            }
        }
    ),
    #"Renamed Canonical Facility" = Table.RenameColumns(
        #"Normalised Mapping Values",
        {{"Unit", "Facility"}}
    ),
    #"Added Source Location Key" = Table.AddColumn(
        #"Renamed Canonical Facility",
        "SourceLocationKey",
        each if [Location] = null then null else Text.Upper([Location]),
        type nullable text
    ),
    // Source codes are the explicit label before the Location delimiter. This
    // avoids assuming every source key is exactly two characters long.
    #"Added Source Facility Code" = Table.AddColumn(
        #"Added Source Location Key",
        "SourceFacilityCode",
        each
            if [Location] = null or [Location] = "" then
                null
            else if Text.Contains([Location], ":") then
                Text.Upper(Text.Trim(Text.BeforeDelimiter([Location], ":")))
            else
                Text.Upper(Text.Trim([Location])),
        type nullable text
    ),
    #"Selected Mapping Contract" = Table.SelectColumns(
        #"Added Source Facility Code",
        {"SourceLocationKey", "SourceFacilityCode", "Facility"}
    ),
    Result = Table.Buffer(Table.Distinct(#"Selected Mapping Contract"))
in
    Result,
#"EXTRACT Units" = Inputs[Units],
#"Master Source Prepare" = Inputs[Master],
#"RoleDemand CS Role Mapping Prepare" = let
    Selected = Table.SelectColumns(#"EXTRACT Roles", {"Roster Roles", "Role Group", "DC Category"}),
    Normalised = Table.TransformColumns(Selected, {
        {"Roster Roles", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text},
        {"Role Group", each if _ = null then null else Text.Trim(_), type nullable text},
        {"DC Category", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text}
    }),
    CSDefinitions = Table.SelectRows(Normalised, each [Role Group] = "CS"),
    Invalid = Table.SelectRows(CSDefinitions, each [Roster Roles] = null or [Roster Roles] = "" or [DC Category] <> "NA"),
    Keys = List.Buffer(List.Distinct(CSDefinitions[Roster Roles])),
    RelevantDefinitions = Table.Distinct(Table.SelectRows(Normalised, each List.Contains(Keys, [Roster Roles]))),
    Counts = Table.Group(RelevantDefinitions, {"Roster Roles"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    CSAnalysis = Table.SelectRows(#"RoleDemand Role Analysis Prepare", each [Roles] = "CS"),
    Enabled = if Table.RowCount(CSAnalysis) = 1 then CSAnalysis{0}[Effort Management Analysis] = true else false,
    Validated = if not Table.IsEmpty(Invalid) or not Table.IsEmpty(Table.SelectRows(Counts, each [Count] <> 1)) then
            error "CS roster labels require one unambiguous non-direct-care role assignment."
        else if not Table.IsEmpty(CSDefinitions) and Table.RowCount(CSAnalysis) <> 1 then error "CS requires a role-analysis setting."
        else if Enabled and Table.IsEmpty(CSDefinitions) then error "CS is enabled without a roster-role mapping."
        else if Enabled then Table.Distinct(CSDefinitions) else Table.FirstN(CSDefinitions, 0)
in
    Table.Buffer(Table.RenameColumns(Table.SelectColumns(Validated, {"Roster Roles"}), {{"Roster Roles", "RosterRoleKey"}})),
#"EXTRACT Roles" = Inputs[Roles],
#"RoleDemand Care Role Mapping Prepare" = let
    Selected = Table.SelectColumns(#"EXTRACT Roles", {"DC Role", "DC Category", "Direct Care %", "Role Group"}),
    Normalised = Table.TransformColumns(Selected, {
        {"DC Role", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text},
        {"DC Category", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text},
        {"Role Group", each if _ = null then null else Text.Trim(_), type nullable text}
    }),
    CareDefinitions = Table.SelectRows(Normalised, each List.Contains({"RN", "OTHER"}, [DC Category])),
    Invalid = Table.SelectRows(CareDefinitions, each [DC Role] = null or [DC Role] = ""
        or [Role Group] = null or [Role Group] = ""
        or not #"RoleDemand Finite Number"([#"Direct Care %"]) or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1),
    ValidDefinitions = if not Table.IsEmpty(Invalid) then error "Invalid direct-care role demand definition." else CareDefinitions,
    Joined = Table.NestedJoin(ValidDefinitions, {"Role Group"}, #"RoleDemand Role Analysis Prepare", {"Roles"}, "Analysis", JoinKind.LeftOuter),
    Missing = Table.SelectRows(Joined, each Table.RowCount([Analysis]) <> 1),
    ValidAnalysis = if not Table.IsEmpty(Missing) then error "A direct-care role group has no unique analysis setting." else Joined,
    Expanded = Table.ExpandTableColumn(ValidAnalysis, "Analysis", {"Effort Management Analysis"}),
    DistinctDefinitions = Table.Distinct(Expanded),
    Counts = Table.Group(DistinctDefinitions, {"DC Role"}, {
        {"Count", each Table.RowCount(_), Int64.Type},
        {"Enabled", each List.Contains([Effort Management Analysis], true), type logical}
    }),
    Conflicts = Table.SelectRows(Counts, each [Enabled] and [Count] <> 1),
    Validated = if not Table.IsEmpty(Conflicts) then
            error Error.Record("Ambiguous care demand mapping", "An enabled DC role cannot share conflicting groups, percentages or inclusion settings.", Conflicts)
        else DistinctDefinitions,
    // Excluded definitions require only a unique lookup key; no arbitrary group is assigned.
    Enabled = Table.SelectRows(Validated, each [Effort Management Analysis] = true),
    DisabledKeys = Table.Distinct(Table.SelectColumns(Table.SelectRows(Validated, each [Effort Management Analysis] <> true), {"DC Role"})),
    Disabled = Table.AddColumn(DisabledKeys, "Effort Management Analysis", each false, type logical),
    Resolved = Table.Combine({Enabled, Disabled}),
    Named = Table.RenameColumns(Resolved, {{"DC Role", "SourceRoleKey"}, {"Role Group", "Role"},
        {"DC Category", "MappedCategory"}, {"Direct Care %", "MappedDirectCare"}})
in
    Table.Buffer(Named),
#"RoleDemand CS Pattern Prepare" = let
    Empty = #table(type table [Facility = text, Role = text, #"Week No" = number, FortnightWeek = number,
        DayOfWeek = number, FortnightDayIndex = number, Shift = text, DemandHRS = number,
        ProductiveHRS = nullable number, DemandBasis = text], {}),
    Result = if Table.IsEmpty(#"RoleDemand CS Role Mapping Prepare") then Empty else
        let
            Weeks = #"RoleDemand Source Weeks Prepare",
            Rows = Table.SelectRows(#"RoleDemand CS Rows Prepare", each [LocationMatchCount] = 1),
            ExpectedFacilities = List.Buffer(List.Distinct(Weeks[Facility])),
            MissingFacilities = List.Difference(ExpectedFacilities, List.Distinct(Rows[Facility])),
            ValidRows = if not List.IsEmpty(MissingFacilities) then
                    error Error.Record("Missing CS demand source", "Enabled CS cannot publish an entirely absent facility as zero demand.", MissingFacilities)
                else Rows,
            // Coverage uses all raw roster roles, since CS can legitimately have no shift on a given day.
            RawCoverage = Table.SelectColumns(#"Master Source Prepare", {"Location", "Week No", "Week Day"}),
            WithLocationKey = Table.AddColumn(RawCoverage, "SourceLocationKey",
                each if [Location] = null then null else Text.Upper(Text.Trim([Location])), type nullable text),
            Locations = Table.Distinct(Table.SelectColumns(#"MW Facility Mapping Prepare", {"SourceLocationKey", "Facility"})),
            MappedCoverage = Table.NestedJoin(WithLocationKey, {"SourceLocationKey"}, Locations, {"SourceLocationKey"}, "LocationMapping", JoinKind.Inner),
            ExpandedCoverage = Table.ExpandTableColumn(MappedCoverage, "LocationMapping", {"Facility"}),
            WithDay = Table.AddColumn(ExpandedCoverage, "DayOfWeek",
                each List.PositionOf({"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}, [#"Week Day"]) + 1, Int64.Type),
            ValidDays = Table.SelectRows(WithDay, each List.Contains({1..7}, [DayOfWeek])),
            Coverage = Table.Group(ValidDays, {"Facility", "Week No"}, {{"DaysPresent", each List.Count(List.Distinct([DayOfWeek])), Int64.Type}}),
            JoinedCoverage = Table.NestedJoin(Weeks, {"Facility", "Week No"}, Coverage, {"Facility", "Week No"}, "Coverage", JoinKind.LeftOuter),
            Incomplete = Table.SelectRows(JoinedCoverage, each Table.RowCount([Coverage]) <> 1 or [Coverage]{0}[DaysPresent] <> 7),
            ValidWeeks = if not Table.IsEmpty(Incomplete) then error "Complete raw facility/week coverage is required to prove zero CS demand cells." else Weeks,
            FortnightDays = Table.SelectColumns(#"MW Fortnight Days", {"FortnightWeek", "DayOfWeek", "FortnightDayIndex"}),
            JoinedDays = Table.NestedJoin(ValidWeeks, {"FortnightWeek"}, FortnightDays, {"FortnightWeek"}, "Days", JoinKind.Inner),
            ExpandedDays = Table.ExpandTableColumn(JoinedDays, "Days", {"DayOfWeek", "FortnightDayIndex"}),
            WithShifts = Table.AddColumn(ExpandedDays, "Shift", each {"AM", "PM", "NIGHT"}, type list),
            Grid = Table.ExpandListColumn(WithShifts, "Shift"),
            Keys = {"Facility", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift"},
            Observed = Table.Group(ValidRows, Keys, {{"DemandHRS", each List.Sum([DemandHRS]), type number}}),
            JoinedHours = Table.NestedJoin(Grid, Keys, Observed, Keys, "Observed", JoinKind.LeftOuter),
            ExpandedHours = Table.ExpandTableColumn(JoinedHours, "Observed", {"DemandHRS"}),
            ProvenZeros = Table.ReplaceValue(ExpandedHours, null, 0, Replacer.ReplaceValue, {"DemandHRS"}),
            WithRole = Table.AddColumn(ProvenZeros, "Role", each "CS", type text),
            WithProductiveHours = Table.AddColumn(WithRole, "ProductiveHRS", each null, type nullable number),
            WithBasis = Table.AddColumn(WithProductiveHours, "DemandBasis", each "MasterRoster", type text),
            SourceTotals = Table.Group(ValidRows, {"Facility", "Week No"}, {{"SourceHours", each List.Sum([DemandHRS]), type number}}),
            PatternTotals = Table.Group(WithBasis, {"Facility", "Week No"}, {{"PatternHours", each List.Sum([DemandHRS]), type number}}),
            JoinedTotals = Table.NestedJoin(SourceTotals, {"Facility", "Week No"}, PatternTotals, {"Facility", "Week No"}, "Pattern", JoinKind.LeftOuter),
            BadTotals = Table.SelectRows(JoinedTotals, each Table.RowCount([Pattern]) <> 1
                or Number.Abs([SourceHours] - [Pattern]{0}[PatternHours]) > 0.0000001),
            Validated = if not Table.IsEmpty(BadTotals) then error "CS standard-shift pattern does not conserve source net hours." else WithBasis
        in Validated
in
    Table.Buffer(Result),
#"MW Fortnight Days" = let
    Weeks = #table(type table [FortnightWeek = Int64.Type], {{1}, {2}}),
    Days = #table(type table [DayOfWeek = Int64.Type, #"Week Day" = text], {
        {1, "Monday"}, {2, "Tuesday"}, {3, "Wednesday"}, {4, "Thursday"},
        {5, "Friday"}, {6, "Saturday"}, {7, "Sunday"}
    }),
    ExpandedDays = Table.ExpandTableColumn(Table.AddColumn(Weeks, "Days", each Days),
        "Days", {"DayOfWeek", "Week Day"}, {"DayOfWeek", "Week Day"}),
    AddedDayIndex = Table.AddColumn(ExpandedDays, "FortnightDayIndex",
        each ([FortnightWeek] - 1) * 7 + [DayOfWeek], Int64.Type),
    Result = Table.AddColumn(AddedDayIndex, "FortnightDay",
        each Text.From([FortnightWeek], "en-AU") & "-" & [#"Week Day"], type text)
in
    Result,
#"RoleDemand CS Rows Prepare" = let
    Empty = #table(type table [Facility = nullable text, #"Week No" = nullable number,
        FortnightWeek = nullable number, DayOfWeek = nullable number, FortnightDayIndex = nullable number,
        Shift = nullable text, DemandHRS = nullable number, LocationMatchCount = Int64.Type], {}),
    Mapping = #"RoleDemand CS Role Mapping Prepare",
    Result = if Table.IsEmpty(Mapping) then Empty else
        let
            Selected = Table.SelectColumns(#"Master Source Prepare", {"Location", "Role", "Week No", "Week Day", "Start Time", "Shift Net Length"}),
            WithRoleKey = Table.AddColumn(Selected, "RosterRoleKey",
                each if [Role] = null then null else Text.Upper(Text.Trim([Role])), type nullable text),
            SelectedCS = Table.NestedJoin(WithRoleKey, {"RosterRoleKey"}, Mapping, {"RosterRoleKey"}, "CSMapping", JoinKind.Inner),
            Narrow = Table.RemoveColumns(SelectedCS, {"Role", "RosterRoleKey", "CSMapping"}),
            WithLocationKey = Table.AddColumn(Narrow, "SourceLocationKey",
                each if [Location] = null then null else Text.Upper(Text.Trim([Location])), type nullable text),
            Locations = Table.Distinct(Table.SelectColumns(#"MW Facility Mapping Prepare", {"SourceLocationKey", "Facility"})),
            JoinedLocations = Table.NestedJoin(WithLocationKey, {"SourceLocationKey"}, Locations, {"SourceLocationKey"}, "LocationMapping", JoinKind.LeftOuter),
            WithLocationCount = Table.AddColumn(JoinedLocations, "LocationMatchCount", each Table.RowCount([LocationMapping]), Int64.Type),
            ValidLocations = if not Table.IsEmpty(Table.SelectRows(WithLocationCount, each [LocationMatchCount] > 1)) then
                    error "A CS source location maps to multiple canonical facilities." else WithLocationCount,
            ExpandedLocation = Table.ExpandTableColumn(ValidLocations, "LocationMapping", {"Facility"}),
            WithDay = Table.AddColumn(ExpandedLocation, "DayOfWeek",
                each List.PositionOf({"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}, [#"Week Day"]) + 1, Int64.Type),
            JoinedWeeks = Table.NestedJoin(WithDay, {"Facility", "Week No"}, #"RoleDemand Source Weeks Prepare",
                {"Facility", "Week No"}, "SourceWeek", JoinKind.LeftOuter),
            WithWeekCount = Table.AddColumn(JoinedWeeks, "WeekMatchCount", each Table.RowCount([SourceWeek]), Int64.Type),
            ExpandedWeeks = Table.ExpandTableColumn(WithWeekCount, "SourceWeek", {"FortnightWeek"}),
            Invalid = Table.SelectRows(ExpandedWeeks, each [LocationMatchCount] = 1 and (
                [WeekMatchCount] <> 1 or not List.Contains({1..7}, [DayOfWeek])
                or not Value.Is([Start Time], type time)
                or not #"RoleDemand Finite Number"([Shift Net Length]) or [Shift Net Length] < 0)),
            ValidRows = if not Table.IsEmpty(Invalid) then error "CS requires valid source weeks, weekdays, start times and nonnegative net hours." else ExpandedWeeks,
            WithDayIndex = Table.AddColumn(ValidRows, "FortnightDayIndex",
                each if [FortnightWeek] = null then null else ([FortnightWeek] - 1) * 7 + [DayOfWeek], type nullable number),
            // Preserve the existing start-time classification; downstream Settings supplies the full standard shift span.
            WithShift = Table.AddColumn(WithDayIndex, "Shift", each
                if not Value.Is([Start Time], type time) then null
                else if [Start Time] <= StartAM then "NIGHT"
                else if [Start Time] <= StartPM then "AM"
                else if [Start Time] <= StartNS then "PM" else "NIGHT", type nullable text),
            NamedHours = Table.RenameColumns(WithShift, {{"Shift Net Length", "DemandHRS"}}),
            Prepared = Table.SelectColumns(NamedHours,
                {"Facility", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift", "DemandHRS", "LocationMatchCount"})
        in Prepared
in
    Table.Buffer(Result),
#"RoleDemand Care Cells Prepare" = let
    UpstreamChecks = MinuteWorkersFTE_CHECK,
    BadChecks = Table.SelectRows(UpstreamChecks, each not List.Contains({"Pass", "Warning"}, [Severity])),
    ValidChecks = if Table.IsEmpty(UpstreamChecks) or not Table.IsEmpty(BadChecks) then
            error "Existing care-demand publication checks must pass before the combined handover." else true,
    Demand = if ValidChecks then #"RoleDemand Care Source Prepare"(MinuteWorkersFTE_TABLE, false) else error "Care checks failed.",
    Profile = #"RoleDemand Care Source Prepare"(MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE, true),
    IncludedDemand = Table.SelectRows(Demand, each [Effort Management Analysis] = true),
    IncludedProfile = Table.SelectRows(Profile, each [Effort Management Analysis] = true),
    Keys = {"Facility", "SourceRoleKey", "Week No", "FortnightDayIndex", "Shift"},
    Orphans = Table.NestedJoin(IncludedDemand, Keys, IncludedProfile, Keys, "Profile", JoinKind.LeftAnti),
    Joined = Table.NestedJoin(IncludedProfile, Keys, IncludedDemand, Keys, "DemandRows", JoinKind.LeftOuter),
    WithCount = Table.AddColumn(Joined, "DemandRowCount", each Table.RowCount([DemandRows]), Int64.Type),
    Invalid = Table.SelectRows(WithCount, each [HistoricalCoverageStatus] <> "PASS"
        or not List.Contains({"PASS", "PASS ZERO"}, [ProfileAlignmentStatus])
        or [RedistributionMatchCount] <> [DemandRowCount]
        or (if [DemandRowCount] = 1 then
            Number.Abs([RedistributedDemandHours] - [DemandRows]{0}[RedistributedDemandHours]) > 0.0000001
            or [#"Direct Care %"] <> [DemandRows]{0}[#"Direct Care %"] or [Role] <> [DemandRows]{0}[Role]
            else [DemandRowCount] <> 0 or [ProfileAlignmentStatus] <> "PASS ZERO"
                or [HistoricalRosterHours] <> 0 or [RedistributedDemandHours] <> 0)
        or ([ProfileAlignmentStatus] = "PASS ZERO" and
            ([RedistributionMatchCount] <> 0 or [HistoricalRosterHours] <> 0 or [RedistributedDemandHours] <> 0))
        or ([ProfileAlignmentStatus] = "PASS" and [HistoricalRosterHours] <= 0)),
    Coverage = Table.Group(IncludedProfile, {"Facility", "SourceRoleKey"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    Validated = if not Table.IsEmpty(Orphans) or not Table.IsEmpty(Invalid)
            or not Table.IsEmpty(Table.SelectRows(Coverage, each [Count] <> 42)) then
            error "Care demand/profile cells do not reconcile or lack a complete 42-cell source pattern."
        else Table.RemoveColumns(WithCount, {"DemandRows", "DemandRowCount"}),
    WithHours = Table.AddColumn(Validated, "DemandHRS", each [RedistributedDemandHours], type number),
    // Apply each original care percentage before summing into a role group.
    WithProductiveHours = Table.AddColumn(WithHours, "ProductiveHRS", each [DemandHRS] * [#"Direct Care %"], type number),
    WithStandardShift = Table.TransformColumns(WithProductiveHours, {{"Shift", each if _ = "NS" then "NIGHT" else _, type text}}),
    Grouped = Table.Group(WithStandardShift, {"Facility", "Role", "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift"}, {
        {"DemandHRS", each List.Sum([DemandHRS]), type number},
        {"ProductiveHRS", each List.Sum([ProductiveHRS]), type nullable number}
    }),
    WithBasis = Table.AddColumn(Grouped, "DemandBasis", each "CareMinutes", type text)
in
    Table.Buffer(WithBasis),
MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE = Inputs[CareProfile],
#"RoleDemand Care Source Prepare" = (Raw as table, IsProfile as logical) as table =>
let
    Common = {"Facility", "Role", "DC Role", "DC Category", "MinuteCategory", "Direct Care %",
        "Week No", "FortnightWeek", "DayOfWeek", "FortnightDayIndex", "Shift", "RedistributedDemandHours"},
    ProfileColumns = {"HistoricalRosterHours", "RedistributionMatchCount", "HistoricalCoverageStatus", "ProfileAlignmentStatus"},
    Selected = Table.SelectColumns(Raw, Common & (if IsProfile then ProfileColumns else {})),
    Named = Table.RenameColumns(Selected, {{"Role", "SourceRole"}}),
    WithKey = Table.AddColumn(Named, "SourceRoleKey",
        each if [DC Role] = null then null else Text.Upper(Text.Trim([DC Role])), type nullable text),
    InvalidKeys = Table.SelectRows(WithKey, each [Facility] = null or Text.Trim([Facility]) = ""
        or [SourceRoleKey] = null or [SourceRoleKey] = "" or [SourceRole] = null
        or Text.Upper(Text.Trim([SourceRole])) <> [SourceRoleKey]
        or not #"RoleDemand Finite Number"([Week No])
        or not List.Contains({1, 2}, [FortnightWeek]) or not List.Contains({1..7}, [DayOfWeek])
        or [FortnightDayIndex] <> ([FortnightWeek] - 1) * 7 + [DayOfWeek]
        or not List.Contains({"AM", "PM", "NS"}, [Shift])
        or not #"RoleDemand Finite Number"([RedistributedDemandHours]) or [RedistributedDemandHours] < 0
        or not #"RoleDemand Finite Number"([#"Direct Care %"]) or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1
        or (if IsProfile then not #"RoleDemand Finite Number"([HistoricalRosterHours]) or [HistoricalRosterHours] < 0 else false)),
    Keys = {"Facility", "SourceRoleKey", "Week No", "FortnightDayIndex", "Shift"},
    DuplicateKeys = Table.RowCount(WithKey) <> Table.RowCount(Table.Distinct(Table.SelectColumns(WithKey, Keys))),
    ValidKeys = if not Table.IsEmpty(InvalidKeys) or DuplicateKeys then error "Invalid or duplicate care demand source cells." else WithKey,
    JoinedWeeks = Table.NestedJoin(ValidKeys, {"Facility", "Week No", "FortnightWeek"}, #"RoleDemand Source Weeks Prepare",
        {"Facility", "Week No", "FortnightWeek"}, "SourceWeek", JoinKind.LeftOuter),
    ValidWeeks = if not Table.IsEmpty(Table.SelectRows(JoinedWeeks, each Table.RowCount([SourceWeek]) <> 1)) then
            error "Care demand does not match the established source fortnight." else Table.RemoveColumns(JoinedWeeks, {"SourceWeek"}),
    JoinedMapping = Table.NestedJoin(ValidWeeks, {"SourceRoleKey"}, #"RoleDemand Care Role Mapping Prepare",
        {"SourceRoleKey"}, "Mapping", JoinKind.LeftOuter),
    ValidMapping = if not Table.IsEmpty(Table.SelectRows(JoinedMapping, each Table.RowCount([Mapping]) <> 1)) then
            error "A published care role has no unique role-demand mapping." else JoinedMapping,
    Expanded = Table.ExpandTableColumn(ValidMapping, "Mapping",
        {"Role", "MappedCategory", "MappedDirectCare", "Effort Management Analysis"}),
    InvalidAttributes = Table.SelectRows(Expanded, each [Effort Management Analysis] = true and (
        Text.Upper(Text.Trim([DC Category])) <> [MappedCategory] or [#"Direct Care %"] <> [MappedDirectCare]
        or [MinuteCategory] <> (if [MappedCategory] = "RN" then "RN" else "OTHERS"))),
    Validated = if not Table.IsEmpty(InvalidAttributes) then error "Published care attributes differ from the role-demand mapping." else Expanded
in
    Table.Buffer(Validated),
MinuteWorkersFTE_TABLE = Inputs[CareDemand],
RoleDemand_TABLE = let
    Checks = Table.Buffer(RoleDemand_CHECK),
    Failures = Table.SelectRows(Checks, each [Severity] = "Error"),
    Validated = if not Table.IsEmpty(Failures) then
        error Error.Record("Role demand publication validation failed",
            Text.Combine(List.Transform(Table.ToRecords(Failures), each [Check] & ": " & [Message]), "#(lf)"), Failures)
        else #"RoleDemand Pattern Prepare"
in
    Validated
in [Demand=RoleDemand_TABLE, Checks=RoleDemand_CHECK],
Fixtures = (Evaluate as function) as table => let
    Near = (A as number, B as number) as logical => Number.Abs(A-B) <= 0.0000001,
    Change = (Base as record, Fields as record) as record => Record.Combine({Base,Fields}),
    Set = (Row as record, Field as text, Value as any) as record => Record.TransformFields(Row,{{Field,each Value}}),
    UpdateFirst = (T as table, Field as text, Value as any) as table =>
        Table.FromRecords({Set(T{0},Field,Value)} & List.Skip(Table.ToRecords(T))),
    Check = (Name as text, Work as function) as record => let
        Attempt = try Work(),
        Passed = if Attempt[HasError] then false else Attempt[Value] = true
    in [Case=Name,Passed=Passed,Details=if Attempt[HasError] then Attempt[Error][Message] else if Passed then null else "Unexpected result"],
    ErrorChecks = (T as table) as table => Table.SelectRows(T,each
        Record.FieldOrDefault(_,"Severity",null) = "Error"
        or List.Contains({"Fail","Error","FAIL","ERROR"},Record.FieldOrDefault(_,"Status",null))),
    Reject = (I as record) as logical => let
        R = Evaluate(I), Checks = try Table.Buffer(R[Checks]),
        Publication = try Table.RowCount(Table.Buffer(R[Demand]))
    in not Checks[HasError] and Table.RowCount(ErrorChecks(Checks[Value])) > 0 and Publication[HasError],
    Roles = #table({"Roster Roles","DC Role","Role Group","DC Category","Direct Care %"},{
        {"Registered Nurse","RN","RN","RN",1},
        {"Assistant in Nursing","AIN","AIN","OTHER",1},
        {"Enrolled Nurse","EN","EN","OTHER",0.5},
        {"Care Services Employee",null,"CS","NA",null},
        {"Care Services Employee alternate",null,"CS","NA",null},
        {"Care Services Hospitality",null,"CSH","NA",null}
    }),
    Analysis = #table({"Roles","Effort Management Analysis","Leave Balance Analysis"},{
        {"RN",true,true},{"AIN",true,true},{"EN",true,true},{"CS",true,true},{"CSH",true,true}
    }),
    Units = #table({"Location","Unit"},{{"BD: Beaudesert","BD"},{"TE: Temora","TE"}}),
    UnitAnalysis = #table({"Title","Effort Management Analysis"},{{"BD",true},{"TE",false}}),
    ShiftEnds = #table(type table [Shift=text,End=time],{{"AM",#time(14,0,0)},{"PM",#time(21,0,0)},{"NS",#time(6,0,0)}}),
    Days = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},
    RawColumns = {"Location","Role","Employee Code","Week No","Week Day","Start Time","End Time","Shift Net Length","Roster Hours"},
    CSAmount = (W as number, D as number, S as number) as number => (if W=1 then 5 else 8) + D/10 + S/100,
    MakeRaw = (Label as text, Hours as function) as table => Table.FromRecords(List.Combine(List.Transform({1..2},(W) =>
        List.Combine(List.Transform({1..7},(D) => List.Transform({1..3},(S) =>
            [Location="BD: Beaudesert",Role=Label,#"Employee Code"=100+S,#"Week No"=29+W,
                #"Week Day"=Days{D-1},#"Start Time"={#time(7,0,0),#time(15,0,0),#time(22,0,0)}{S-1},
                #"End Time"={#time(14,0,0),#time(21,0,0),#time(6,0,0)}{S-1},
                #"Shift Net Length"=Hours(W,D,S),#"Roster Hours"=9999]
        ))))),RawColumns),
    CSRaw = MakeRaw("Care Services Employee",CSAmount),
    CareRaw = MakeRaw("Registered Nurse",(W,D,S) => 1),
    HospitalityRaw = MakeRaw("Care Services Hospitality",(W,D,S) => 1000),
    Master = Table.Combine({CSRaw,CareRaw,HospitalityRaw}),
    CareDefinitions = {[Role="RN",Category="RN",Percentage=1,Base=8],
        [Role="AIN",Category="OTHER",Percentage=1,Base=10],[Role="EN",Category="OTHER",Percentage=0.5,Base=4]},
    CareRows = Table.FromRecords(List.Combine(List.Transform(CareDefinitions,(R) =>
        List.Combine(List.Transform({1..14},(D) => List.Transform({1..3},(S) => let
            W = 1 + Number.IntegerDivide(D-1,7), Day = Number.Mod(D-1,7)+1,
            H = R[Base] + W + Day/10 + S/100
        in [Facility="BD",Role=R[Role],#"DC Role"=R[Role],#"DC Category"=R[Category],
            MinuteCategory=if R[Category]="RN" then "RN" else "OTHERS",#"Direct Care %"=R[Percentage],
            #"Week No"=29+W,FortnightWeek=W,DayOfWeek=Day,#"Week Day"=Days{Day-1},
            FortnightDayIndex=D,FortnightDay=Text.From(W)&"-"&Days{Day-1},Shift={"AM","PM","NS"}{S-1},ShiftIndex=S,
            RedistributedDemandHours=H,FTE=H/7.6,WeekdayShiftTargetMinutes=H*R[Percentage]*60,
            HistoricalRosterHours=H/1.25,HistoricalProductiveHours=H/1.25*R[Percentage],
            HistoricalCoverageStatus="PASS",ProfileAlignmentStatus="PASS",RedistributionMatchCount=1]
        )))))),
    CareChecks = #table(type table [Check=nullable text,Severity=nullable text,Facility=nullable text,
        MinuteCategory=nullable text,Role=nullable text,Actual=nullable number,Expected=nullable number,Message=nullable text],{}),
    Inputs = [Master=Master,Roles=Roles,Analysis=Analysis,Units=Units,UnitAnalysis=UnitAnalysis,ShiftEnds=ShiftEnds,
        CareDemand=CareRows,CareProfile=CareRows,CareChecks=CareChecks],
    Result = Evaluate(Inputs),
    Demand = Table.Buffer(Result[Demand]),
    Checks = Table.Buffer(Result[Checks]),
    CS = Table.SelectRows(Demand,each [Role]="CS"),
    Care = Table.SelectRows(Demand,each [Role]<>"CS"),
    ExpectedColumns = {"Facility","Role","Week No","FortnightWeek","DayOfWeek","FortnightDayIndex","Shift","DemandHRS","ProductiveHRS","DemandBasis"},
    Cell = (T as table,R as text,D as number,S as text) as record => Table.SelectRows(T,each [Role]=R and [FortnightDayIndex]=D and [Shift]=S){0},
    Sorted = (T as table) as table => Table.Sort(Table.SelectColumns(T,ExpectedColumns),{{"Facility",Order.Ascending},{"Role",Order.Ascending},{"FortnightDayIndex",Order.Ascending},{"Shift",Order.Ascending}}),
    SameCare = (I as record) as logical => Sorted(Table.SelectRows(Evaluate(I)[Demand],each [Role]<>"CS")) = Sorted(Care),
    WithoutCSCell = Table.SelectRows(Master,each not ([Role]="Care Services Employee" and [#"Week No"]=30 and [#"Week Day"]="Monday" and [#"Start Time"]=#time(7,0,0))),
    MissingAllDay = Table.SelectRows(Master,each not ([#"Week No"]=30 and [#"Week Day"]="Monday")),
    SparseWeeks = Table.SelectRows(Master,each not ([Role]="Care Services Employee" and [#"Week No"]=30)),
    ForwardingWarnings = #table(Table.ColumnNames(CareChecks),{
        {"Master Roster Location is outside configured Units","Warning",null,null,null,1,0,"Known excluded-location warning"},
        {"Synthetic global care warning","Warning",null,null,null,1,0,"Keep this global warning"},
        {"Synthetic BD care warning","Warning","BD",null,"RN",1,0,"Keep this enabled-facility warning"},
        {"Synthetic TE care warning","Warning","TE",null,"RN",1,0,"Keep this warning despite the unrelated disabled facility flag"}
    }),
    WarningResult = Evaluate(Change(Inputs,[CareChecks=ForwardingWarnings])),
    ForwardedChecks = Table.Buffer(WarningResult[Checks]),
    Rows = {
        Check("Published common schema is exact",()=>Table.ColumnNames(Demand)=ExpectedColumns),
        Check("Every published value and check is evaluated without cell errors",()=>Table.RowCount(Table.SelectRowsWithErrors(Demand))=0 and Table.RowCount(Table.SelectRowsWithErrors(Checks))=0 and Table.RowCount(ErrorChecks(Checks))=0),
        Check("All three care groups and CS publish unique complete fortnight patterns",()=>Table.RowCount(Demand)=4*42 and Table.RowCount(Table.Distinct(Demand,{"Facility","Role","Week No","DayOfWeek","Shift"}))=4*42 and List.Sort(List.Distinct(Demand[Role]))={"AIN","CS","EN","RN"}),
        Check("CS uses raw net hours and preserves every source cell",()=>Table.RowCount(CS)=42 and List.AllTrue(List.Transform(Table.ToRecords(CS),each Near([DemandHRS],CSAmount([FortnightWeek],[DayOfWeek],List.PositionOf({"AM","PM","NIGHT"},[Shift])+1))))),
        Check("CS fortnight total is the independent raw net-hour sum",()=>Near(List.Sum(CS[DemandHRS]),List.Sum(CSRaw[Shift Net Length]))),
        Check("CS basis is MasterRoster and productive hours remain null",()=>List.Distinct(CS[DemandBasis])={"MasterRoster"} and List.NonNullCount(CS[ProductiveHRS])=0),
        Check("Enabled CS has an explicit passing inclusion diagnostic",()=>let R=Table.SelectRows(Checks,each [Check]="CS effort-analysis inclusion" and [Role]="CS") in Table.RowCount(R)=1 and R{0}[Severity]="Pass" and R{0}[Actual]=1 and R{0}[Expected]=1),
        Check("Care roster demand remains unchanged",()=>List.AllTrue(List.Transform(Table.ToRecords(CareRows),(R)=>Near(Cell(Care,R[Role],R[FortnightDayIndex],if R[Shift]="NS" then "NIGHT" else R[Shift])[DemandHRS],R[RedistributedDemandHours])))),
        Check("Care productive hours apply each source role percentage once",()=>List.AllTrue(List.Transform(Table.ToRecords(CareRows),(R)=>Near(Cell(Care,R[Role],R[FortnightDayIndex],if R[Shift]="NS" then "NIGHT" else R[Shift])[ProductiveHRS],R[RedistributedDemandHours]*R[#"Direct Care %"]))) and List.Distinct(Care[DemandBasis])={"CareMinutes"}),
        Check("The two source weeks remain distinct without averaging",()=>Near(Cell(CS,"CS",1,"AM")[DemandHRS],5.11) and Near(Cell(CS,"CS",8,"AM")[DemandHRS],8.11) and Cell(CS,"CS",1,"AM")[Week No]=30 and Cell(CS,"CS",8,"AM")[Week No]=31),
        Check("Night shift is published as NIGHT",()=>List.Sort(List.Distinct(Demand[Shift]))={"AM","NIGHT","PM"}),
        Check("CSH is excluded even though its own effort flag is true",()=>not List.Contains(Demand[Role],"CSH") and Near(List.Sum(CS[DemandHRS]),List.Sum(CSRaw[Shift Net Length]))),
        Check("False or null CS inclusion warns explicitly and preserves care publication",()=>List.AllTrue(List.Transform({false,null},(Flag)=>let
            A=Table.FromRecords(List.Transform(Table.ToRecords(Analysis),each if [Roles]="CS" then Set(_,"Effort Management Analysis",Flag) else _)),
            I=Change(Inputs,[Analysis=A]), R=Evaluate(I), C=Table.Buffer(R[Checks]),
            Inclusion=Table.SelectRows(C,each [Check]="CS effort-analysis inclusion" and [Role]="CS")
        in Table.RowCount(Table.Buffer(R[Demand]))=126 and SameCare(I)
            and Table.RowCount(Inclusion)=1 and Inclusion{0}[Severity]="Warning"
            and Inclusion{0}[Actual]=0 and Inclusion{0}[Expected]=1))),
        Check("Known out-of-scope care-location warning is not forwarded",()=>Table.RowCount(Table.SelectRows(ForwardedChecks,each [Check]="Master Roster Location is outside configured Units"))=0),
        Check("Other global care warnings remain unchanged",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic global care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{1}),
        Check("In-scope facility care warnings remain unchanged",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic BD care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{2}),
        Check("Facility Analysis does not hide other facility care warnings",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic TE care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{3}),
        Check("Forwarding diagnostics does not change demand or introduce error checks",()=>Sorted(Table.Buffer(WarningResult[Demand]))=Sorted(Demand) and Table.RowCount(Table.SelectRowsWithErrors(ForwardedChecks))=0 and Table.RowCount(ErrorChecks(ForwardedChecks))=0),
        Check("Identical mapping duplicates do not multiply CS demand",()=>let I=Change(Inputs,[Roles=Table.Combine({Roles,Table.SelectRows(Roles,each [Role Group]="CS")})]) in Sorted(Evaluate(I)[Demand])=Sorted(Demand)),
        Check("A second configured CS roster label adds its exact net hours",()=>let Added=Set(Set(CSRaw{0},"Role","Care Services Employee alternate"),"Shift Net Length",1.25), I=Change(Inputs,[Master=Table.Combine({Master,Table.FromRecords({Added},RawColumns)})]) in Near(Cell(Evaluate(I)[Demand],"CS",1,"AM")[DemandHRS],6.36) and SameCare(I)),
        Check("Missing CS source cell becomes a proven zero when facility-day coverage exists",()=>let R=Evaluate(Change(Inputs,[Master=WithoutCSCell]))[Demand] in Table.RowCount(Table.SelectRows(R,each [Role]="CS"))=42 and Cell(R,"CS",1,"AM")[DemandHRS]=0 and Near(Cell(R,"CS",8,"AM")[DemandHRS],8.11)),
        Check("CS sparse first week retains the actual second-week key",()=>let R=Evaluate(Change(Inputs,[Master=SparseWeeks]))[Demand] in List.Sum(Table.SelectRows(R,each [Role]="CS" and [FortnightWeek]=1)[DemandHRS])=0 and Near(Cell(R,"CS",8,"AM")[DemandHRS],8.11)),
        Check("Missing all source coverage for a facility day blocks publication",()=>Reject(Change(Inputs,[Master=MissingAllDay]))),
        Check("Enabled CS without any CS source rows blocks publication",()=>Reject(Change(Inputs,[Master=Table.SelectRows(Master,each [Role]<>"Care Services Employee")]))),
        Check("Negative CS net hours block publication",()=>Reject(Change(Inputs,[Master=Table.Combine({UpdateFirst(CSRaw,"Shift Net Length",-1),CareRaw,HospitalityRaw})]))),
        Check("Null CS net hours block publication",()=>Reject(Change(Inputs,[Master=Table.Combine({UpdateFirst(CSRaw,"Shift Net Length",null),CareRaw,HospitalityRaw})]))),
        Check("Missing care facility field blocks publication",()=>Reject(Change(Inputs,[CareProfile=Table.RemoveColumns(CareRows,{"Facility"})]))),
        Check("Blank care facility key blocks publication",()=>Reject(Change(Inputs,[CareProfile=UpdateFirst(CareRows,"Facility",null)]))),
        Check("Conflicting CS mapping blocks publication",()=>Reject(Change(Inputs,[Roles=Table.Combine({Roles,#table(Table.ColumnNames(Roles),{{"Care Services Employee",null,"CSH","NA",null}})})]))),
        Check("Duplicate analysis keys block publication without an inclusion-control row",()=>let I=Change(Inputs,[Analysis=Table.Combine({Analysis,Table.SelectRows(Analysis,each [Roles]="CS")})]) in Reject(I) and Table.RowCount(Table.SelectRows(Evaluate(I)[Checks],each [Check]="CS effort-analysis inclusion"))=0),
        Check("Missing CS analysis row blocks publication without an inclusion-control row",()=>let I=Change(Inputs,[Analysis=Table.SelectRows(Analysis,each [Roles]<>"CS")]) in Reject(I) and Table.RowCount(Table.SelectRows(Evaluate(I)[Checks],each [Check]="CS effort-analysis inclusion"))=0),
        Check("Ambiguous configured facility mapping blocks publication",()=>Reject(Change(Inputs,[Units=Table.Combine({Units,#table({"Location","Unit"},{{"BD: Beaudesert","TE"}})})]))),
        Check("Unlisted CS locations stay excluded from demand and mapped-row counts without a summary warning",()=>let
            Added=Set(CSRaw{0},"Location","XX: Unlisted"),
            R=Evaluate(Change(Inputs,[Master=Table.Combine({Master,Table.FromRecords({Added},RawColumns)})])),
            C=Table.Buffer(R[Checks]), Stage=Table.SelectRows(C,each [Check]="CS roster-role and source rows")
        in Sorted(R[Demand])=Sorted(Demand) and Table.RowCount(Stage)=1 and Stage{0}[Actual]=42
            and Table.RowCount(Table.SelectRows(C,each [Check]="CS unlisted source locations"))=0),
        Check("Legacy care publication errors block the common publication",()=>Reject(Change(Inputs,[CareChecks=#table(Table.ColumnNames(CareChecks),{{"Synthetic care error","Error","BD",null,"RN",1,0,"Deliberate input error"}})])))
    }
in Table.FromRecords(Rows)
in Fixtures(Evaluate);
