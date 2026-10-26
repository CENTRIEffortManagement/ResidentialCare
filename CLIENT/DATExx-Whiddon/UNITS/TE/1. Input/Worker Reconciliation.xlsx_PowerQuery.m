section Section1;

// Query: IMPORT CentriSyncPaths
// Purpose: Read the fixed public-machine workbook used by the ResidentialCare path resolver.
shared #"IMPORT CentriSyncPaths" = let
    SourcePath = "C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true)
in
    Navigation;

// Query: PathTABLE
// Purpose: Resolve the saved workbook path through the shared ResidentialCare sync mapping.
// Inputs: FilePathUrl and the fixed public-machine CentriSyncPaths table.
// Output: Workbook, unit and portfolio roots used by every external import in this source.
shared PathTABLE = let
    Source = Excel.CurrentWorkbook(),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TrimmedValue = if value = null then null else Text.Trim(value),
            SlashNormalized = if value = null then null else Text.Replace(TrimmedValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    WorkbookPathInputs = Table.SelectRows(Source, each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
    WorkbookPathTable = if Table.RowCount(WorkbookPathInputs) = 1 then WorkbookPathInputs{0}[Content]
        else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
    WorkbookPathColumn = if Table.HasColumns(WorkbookPathTable, {"FilePath"}) then Table.SelectColumns(WorkbookPathTable, {"FilePath"})
        else if Table.ColumnNames(WorkbookPathTable) = {"Column1"} then Table.RenameColumns(WorkbookPathTable, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or be a single named cell.",
    WorkbookPathRow = if Table.RowCount(WorkbookPathColumn) = 1 then WorkbookPathColumn{0}
        else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = WorkbookPathRow[FilePath],
    NormalizedWorkbookPath = if not Value.Is(RawFilePath, type text) then
            error "FilePathUrl[FilePath] must contain the current workbook path as text."
        else if Text.Trim(RawFilePath) = "" then
            error "FilePathUrl[FilePath] is blank. Save and recalculate Worker Reconciliation.xlsx."
        else NormalizePath(RawFilePath),
    InputWorkbookFolder = Text.BeforeDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputWorkbookSuffix = Text.AfterDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputFileName = if Text.StartsWith(InputWorkbookSuffix, "[") then Text.BetweenDelimiters(InputWorkbookSuffix, "[", "]") else InputWorkbookSuffix,
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "Worker Reconciliation.xlsx") = 0 then InputWorkbookFolder & "\" & InputFileName
        else error "FilePathUrl must identify Worker Reconciliation.xlsx.",
    CentriSyncPaths_Source = #"IMPORT CentriSyncPaths",
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_Selected = Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}),
    CentriSyncPaths_Typed = Table.TransformColumnTypes(
        CentriSyncPaths_Selected,
        {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    CentriSyncPaths_Normalized = Table.TransformColumns(CentriSyncPaths_Typed, {
        {"SharepointRootUrl", each NormalizePath(_), type nullable text},
        {"SyncedFolderRootPath", each NormalizePath(_), type nullable text}
    }),
    BufferedMappings = Table.Buffer(CentriSyncPaths_Normalized),
    SharePointCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    SharePointDocumentsCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each
        if [SharepointRootUrl] = null or Text.Trim([SharepointRootUrl]) = "" then null
        else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    CandidatesWithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each
        if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(CandidatesWithLength, each
        [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    LongestMatch = if Table.IsEmpty(SortedMatches) then error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath else SortedMatches{0},
    LongestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    MatchedDestinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch = if List.Count(MatchedDestinations) = 1 then LongestMatch
        else error "CentriSyncPaths contains conflicting mappings for: " & FilePath,
    UntrimmedRelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath = Text.TrimStart(UntrimmedRelativePath, "\"),
    LocalFullPath = if RelativePath = "" then BestMatch[SyncedFolderRootPath]
        else BestMatch[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    FolderSegments = Text.Split(WorkbookFolder, "\"),
    Segments = List.Select(FolderSegments, each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase),
    InputIndex = List.PositionOf(Segments, "1. Input", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidUnitFolder = UnitsIndex >= 0 and InputIndex = UnitsIndex + 2 and InputIndex = List.Count(Segments) - 1,
    UnitRoot = if ValidUnitFolder then Text.BeforeDelimiter(WorkbookFolder, "\", {0, RelativePosition.FromEnd})
        else error "Worker Reconciliation.xlsx must be under UNITS/<unit>/1. Input.",
    PortfolioSegments = List.FirstN(Segments, ResidentialCareIndex),
    PortfolioRoot = if ResidentialCareIndex > 0 then Text.Combine(PortfolioSegments, "\")
        else error "Resolved workbook path does not contain the ResidentialCare boundary.",
    UnitName = Segments{UnitsIndex + 1},
    Output = #table(
        {"Variable Name", "Value"},
        {
            {"Root Path", WorkbookFolder},
            {"Unit Root", UnitRoot},
            {"Portfolio Root", PortfolioRoot},
            {"Unit", UnitName},
            {"FileName", InputFileName}
        }),
    #"Buffered Path Output" = Table.Buffer(Output)
in
    #"Buffered Path Output";

// Query: Unit1Path
// Purpose: Expose the resolved unit folder for local ResidentialCare imports.
shared Unit1Path = let
    Source = PathTABLE,
    SelectedPathRow = Source{[Variable Name = "Unit Root"]},
    UnitPath = SelectedPathRow[Value]
in
    UnitPath;

// Query: PortfolioPath
// Purpose: Preserve the resolved sync portfolio folder as an existing workbook interface.
shared PortfolioPath = let
    Source = PathTABLE,
    SelectedPathRow = Source{[Variable Name = "Portfolio Root"]},
    PortfolioRootPath = SelectedPathRow[Value]
in
    PortfolioRootPath;

// Query: AnalysisFacilityCodes
// Purpose: Define the approved analysis-facility codes shared by leave membership, summaries and publication checks.
// Output: BD, JE, NR and TE using the existing two-character facility-key convention.
// Notes: Employee home locations do not restrict employee-ID role, contract or termination enrichment.
shared AnalysisFacilityCodes = let
    ApprovedFacilityCodes = {"BD", "JE", "NR", "TE"}
in
    ApprovedFacilityCodes;

// Query: IMPORT Combined Roster
// Purpose: Read the local published-roster Roster Data sheet using its existing prepared import interface.
// Inputs: Roster Data sheet in 1. Input/Allocation/Published Roster.xlsx.
// Output: The existing header-promoted, typed roster table used by Roster Prepare.
// Notes: Header promotion and source typing are retained here to preserve this established import interface.
shared #"IMPORT Combined Roster" = let
    SourcePath = Unit1Path & "\1. Input\Allocation\Published Roster.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    #"Selected Roster Data Sheet" = Navigation{[Item="Roster Data",Kind="Sheet"]}[Data],
    #"Promoted Roster Headers" = Table.PromoteHeaders(#"Selected Roster Data Sheet", [PromoteAllScalars=true]),
    #"Typed Roster Source Columns" = Table.TransformColumnTypes(#"Promoted Roster Headers",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", type text}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", type text}})
in
    #"Typed Roster Source Columns";

// Query: IMPORT Combined Availabilities
// Purpose: Read the local availability/leave Combined Output sheet unchanged for records-worker preparation.
shared #"IMPORT Combined Availabilities" = let
    SourcePath = Unit1Path & "\1. Input\whiddon_availability_leave_extraction.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "Combined Output" and [Kind] = "Sheet"),
    CombinedOutput = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Worker reconciliation records import", "Expected exactly one Combined Output sheet.", [Matches = Table.RowCount(Matches)])
in
    CombinedOutput;

// Query: IMPORT Leave Balances
// Purpose: Read the approved LeaveBalance report navigation unchanged for worker-source reconciliation.
// Inputs: LeaveBalance/Workflows/LeaveBalance/Analysis/1.Input/Leave Balances Report.xlsx beneath PortfolioPath.
shared #"IMPORT Leave Balances" = let
    SourcePath = PortfolioPath & "\LeaveBalance\Workflows\LeaveBalance\Analysis\1.Input\Leave Balances Report.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    Sheet1_Sheet = Navigation{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Employee Code", Int64.Type}, {"Facility", type text}, {"Position Description", type text}, {"Employee Status Description", type text}, {"Annual Leave Table Code", type text}, {"Department Description", type text}, {"AL Units Available", type number}, {"AL Total Liability", type number}, {"LS Units Available", type number}, {"LS Total Liability", type number}})
in
    #"Changed Type";

// Query: LeaveBalances_Prepare
// Purpose: Prepare leave worker identity and restrict leave membership to the approved analysis facilities.
// Inputs: Table2 from IMPORT Leave Balances and AnalysisFacilityCodes.
// Output: Employee Code, Facility, Role and the existing two-character Facility-Abbrev key for analysis sites.
// Notes: Both hard-coded CS exclusions remain; downstream leave-identity checks apply to this scoped population.
shared LeaveBalances_Prepare = let
    Source = #"IMPORT Leave Balances",

    Selected = Table.SelectColumns(
        Source,
        {"Employee Code", "Facility", "Position Description"}
    ),

    Typed = Table.TransformColumnTypes(Selected, {
        {"Employee Code", type text},
        {"Facility", type text},
        {"Position Description", type text}
    }),

    CleanText = (value as nullable text) as nullable text =>
        let
            Trimmed = if value = null then null else Text.Trim(value),
            Output = if Trimmed = "" then null else Trimmed
        in
            Output,

    Normalized = Table.TransformColumns(Typed, {
        {"Employee Code", CleanText, type nullable text},
        {"Facility", CleanText, type nullable text},
        {"Position Description", CleanText, type nullable text}
    }),

    RenamedRole = Table.RenameColumns(
        Normalized,
        {{"Position Description", "Role"}}
    ),

    WithFacilityAbbrev = Table.AddColumn(
        RenamedRole,
        "Facility-Abbrev",
        each if [Facility] = null then null
            else Text.Upper(Text.Start([Facility], 2)),
        type nullable text
    ),

    #"Filter LOCATIONS" = Table.SelectRows(WithFacilityAbbrev, each
        List.Contains(AnalysisFacilityCodes, [#"Facility-Abbrev"], Comparer.Ordinal)
            and [Facility] <> "CS: Community Serv Beaudesert"
            and [Facility] <> "CS: Community Services Narrabr"
    ),

    BufferedLeaveWorkers = Table.Buffer(#"Filter LOCATIONS")
in
    BufferedLeaveWorkers;

// Query: LeaveBalances_Workers
// Purpose: Reduce leave-report membership to one employee-ID row without multiplying reconciled workers.
// Inputs: LeaveBalances_Prepare.
// Output: One leave-source marker, role, facility key and original facility label per Employee Code.
// Notes: Conflicting roles or six-character locations are exposed by LeaveBalances_DIAGNOSTICS and block publication.
shared LeaveBalances_Workers = let
    Source = LeaveBalances_Prepare,
    GroupedWorkers = Table.Group(Source, {"Employee Code"}, {
        {"Leave Balance Role", each
            let
                Roles = List.RemoveNulls([Role]),
                DistinctRoles = List.Distinct(Roles, Comparer.OrdinalIgnoreCase),
                SortedRoles = List.Sort(DistinctRoles, Comparer.OrdinalIgnoreCase)
            in
                if List.IsEmpty(SortedRoles) then null else SortedRoles{0}, type nullable text},
        {"Leave Balance Location", each
            let
                Locations = List.RemoveNulls([Facility]),
                SortedLocations = List.Sort(Locations, Comparer.Ordinal)
            in
                if List.IsEmpty(SortedLocations) then null else SortedLocations{0}, type nullable text},
        // Repeated rows with the same role and comparison location are one membership, even if balance values differ.
        {"Leave Balance Identity Count", each
            let
                IdentityFields = Table.SelectColumns(_, {"Facility", "Role"}),
                ComparisonIdentity = Table.TransformColumns(IdentityFields, {
                    {"Facility", each if _ = null then null else Text.Start(_, 6), type nullable text},
                    {"Role", each if _ = null then null else Text.Upper(_), type nullable text}}),
                DistinctIdentities = Table.Distinct(ComparisonIdentity)
            in
                Table.RowCount(DistinctIdentities), Int64.Type}
    }),
    WithFacilityAbbrev = Table.AddColumn(GroupedWorkers, "Leave Balance Facility-Abbrev", each
        if [Leave Balance Location] = null then null else Text.Upper(Text.Start([Leave Balance Location], 2)), type nullable text),
    WithSourceMarker = Table.AddColumn(WithFacilityAbbrev, "In Leave Balances Source", each true, type logical),
    #"Buffered Leave Worker Output" = Table.Buffer(WithSourceMarker)
in
    #"Buffered Leave Worker Output";

// Query: Roster Prepare
// Purpose: Normalize the published roster and derive the base role used for reconciliation.
shared #"Roster Prepare" = let
    Source = #"IMPORT Combined Roster",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", type text}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", type text}}),
    // Only a terminal AM, PM or NS token is a shift suffix. Non-shift titles must not lose their final three characters.
    #"Added Shift Suffix Flag" = Table.AddColumn(#"Changed Type", "Shift", each
        let
            RoleValue = if [Role] = null then null else Text.Upper(Text.TrimEnd(Text.From([Role])))
        in
            RoleValue <> null and List.AnyTrue(List.Transform({" AM", " PM", " NS"}, each Text.EndsWith(RoleValue, _))),
        type logical),
    #"Renamed Source Role" = Table.RenameColumns(#"Added Shift Suffix Flag",{{"Role", "RoleX"}}),
    #"Added Base Role" = Table.AddColumn(
    #"Renamed Source Role",
    "Role",
    each
        if [Department] = "Care" or [Department] = "Catering" then
            let
                RoleValue = if [RoleX] = null then null else Text.Trim(Text.From([RoleX])),
                BaseRole = if RoleValue = null then null
                    else if [Shift] then Text.TrimEnd(Text.Start(RoleValue, Text.Length(RoleValue) - 2))
                    else RoleValue
            in
                BaseRole
        else
            null,
    type text
),
    #"Expanded Assistant Label" = Table.ReplaceValue(#"Added Base Role","Asst","Assistant",Replacer.ReplaceText,{"Role"}),
    #"Filter LOCATIONS" = Table.SelectRows(#"Expanded Assistant Label", each ([Location] = "BD: Beaudesert" or [Location] = "JE: Jessie Hunt" or [Location] = "NR: Narrabri- Robert Young" or [Location] = "TE: Temora " or [Location] = "TE: Temora Greenstone"))
in
    #"Filter LOCATIONS";

// Query: Roster Blank AGENCY
// Purpose: Review the existing blank or agency roster identity projection.
// Inputs: IMPORT Combined Roster.
// Notes: Preserve the existing source-value comparisons, label shortening and location-filter order.
shared #"Roster Blank AGENCY" = let
    Source = #"IMPORT Combined Roster",
    #"Kept Blank Or Agency Employment" = Table.SelectRows(Source, each [Employment Type] = " " or [Employment Type] = "Agency "),
    #"Sorted Employment Types" = Table.Sort(#"Kept Blank Or Agency Employment", {{"Employment Type", Order.Descending}}),
    #"Distinct Employee Locations" = Table.Distinct(#"Sorted Employment Types", {"Employee Code", "Location"}),
    #"Selected Roster Audit Fields" = Table.SelectColumns(#"Distinct Employee Locations", {"Location", "Employment Type", "Role", "Employee Code"}),
    #"Shortened Role Labels" = Table.TransformColumns(#"Selected Roster Audit Fields", {{"Role", each if _ = null then null else Text.Start(_, 15), type nullable text}}),
    #"Shortened Location Labels" = Table.TransformColumns(#"Shortened Role Labels", {{"Location", each if _ = null then null else Text.Start(_, 2), type nullable text}}),
    #"Distinct Audit Rows" = Table.Distinct(#"Shortened Location Labels"),
    #"Sorted Employee Codes" = Table.Sort(#"Distinct Audit Rows", {{"Employee Code", Order.Descending}}),
    #"Filter LOCATIONS" = Table.SelectRows(#"Sorted Employee Codes", each
        [Location] = "BD: Beaudesert" or [Location] = "JE: Jessie Hunt" or [Location] = "NR: Narrabri- Robert Young"
            or [Location] = "TE: Temora " or [Location] = "TE: Temora Greenstone")
in
    #"Filter LOCATIONS";

// Query: RosteredEmployeeShiftsCOUNT
// Purpose: Count published roster rows by location and employment type.
// Inputs: Roster Prepare after its analysis-location filter.
// Output: One location row with counts pivoted across the observed employment types.
shared RosteredEmployeeShiftsCOUNT = let
    Source = #"Roster Prepare",
    #"Counted Location Employment Types" = Table.Group(Source, {"Location", "Employment Type"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    EmploymentTypeValues = #"Counted Location Employment Types"[Employment Type],
    DistinctEmploymentTypes = List.Distinct(EmploymentTypeValues),
    #"Pivoted Roster Counts" = Table.Pivot(#"Counted Location Employment Types", DistinctEmploymentTypes, "Employment Type", "Count", List.Sum)
in
    #"Pivoted Roster Counts";

// Query: RosteredEmployeeOnlyShifts
// Purpose: Retain the existing non-agency roster rows with an assigned employee code.
// Inputs: Roster Prepare after its analysis-location filter.
// Notes: Preserve the existing exact source values used for Agency and blank-code exclusions.
shared RosteredEmployeeOnlyShifts = let
    Source = #"Roster Prepare",
    #"Kept Non Agency Shifts" = Table.SelectRows(Source, each [Employment Type] <> "Agency "),
    #"Kept Assigned Employee Codes" = Table.SelectRows(#"Kept Non Agency Shifts", each [Employee Code] <> " ")
in
    #"Kept Assigned Employee Codes";

// Query: RosteredEmployeesTABLE
// Purpose: Reduce allocation shifts to distinct worker/facility/role candidates with normalized join keys.
shared RosteredEmployeesTABLE = let
    Source = #"Roster Prepare",
    #"Selected Worker Fields" = Table.SelectColumns(Source,{"Location", "Department", "Employee Roster Name", "Employment Type", "Employee Code", "Role"}),
    #"Added Facility" = Table.AddColumn(#"Selected Worker Fields", "Facility-Abbrev", each
        let Location = if [Location] = null then null else Text.Trim([Location])
        in if Location = null or Location = "" then null else Text.Upper(Text.Start(Location, 2)), type nullable text),
    #"Normalized Worker Fields" = Table.TransformColumns(#"Added Facility", {
        {"Employee Code", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Employee Roster Name", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Employment Type", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Role", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text}
    }),
    #"Replaced REGEN-" = Table.ReplaceValue(#"Normalized Worker Fields","REGN -  In Charge","REGN",Replacer.ReplaceText,{"Role"}),
    #"Replaced REGEN" = Table.ReplaceValue(#"Replaced REGEN-","REGN - In Charge","REGN",Replacer.ReplaceText,{"Role"}),
    #"Replaced ASSISTANT IN NURSING" = Table.ReplaceValue(#"Replaced REGEN","Assistant in Nursing ","Assistant in Nursing",Replacer.ReplaceText,{"Role"}),
    // Retain the original location label while preserving the existing worker deduplication keys.
    #"Worker Deduplication Columns" = List.RemoveItems(Table.ColumnNames(#"Replaced ASSISTANT IN NURSING"), {"Location"}),
    #"Removed Exact Duplicates" = Table.Distinct(#"Replaced ASSISTANT IN NURSING", #"Worker Deduplication Columns"),
    // Blank roster lines are vacancies or unassigned shifts, not workers.
    #"Filtered Identified Workers" = Table.SelectRows(#"Removed Exact Duplicates", each [Employee Code] <> null),
    // Preserve source presence independently of whether the worker identifier is complete.
    #"Added Allocation Source Marker" = Table.AddColumn(#"Filtered Identified Workers", "In Allocation Source", each true, type logical),
    #"Buffered Roster Workers" = Table.Buffer(#"Added Allocation Source Marker"),
    #"Kept Non Agency Workers" = Table.SelectRows(#"Buffered Roster Workers", each [Employment Type] <> "Agency")
in
    #"Kept Non Agency Workers";

// Query: RosterStartDate
// Purpose: Expose the first published-roster date used by employment eligibility and termination checks.
shared RosterStartDate = let
    RosterDates = List.RemoveNulls(#"Roster Prepare"[Date]),
    Output = if List.IsEmpty(RosterDates) then error "Published Roster contains no roster dates." else List.Min(RosterDates)
in
    Output;

// Query: IMPORT Current_Workers
// Purpose: Read the unfiltered worker report using its existing prepared import interface.
// Inputs: Sheet1 in 1. Input/Worker's Report.xlsx.
// Output: The existing header-promoted, typed worker report used for employment, role and contract enrichment.
// Notes: Header promotion and source typing are retained here to preserve this established import interface.
shared #"IMPORT Current_Workers" = let
    SourcePath = Unit1Path & "\1. Input\Worker's Report.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    #"Selected Worker Report Sheet" = Navigation{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Worker Report Headers" = Table.PromoteHeaders(#"Selected Worker Report Sheet", [PromoteAllScalars=true]),
    #"Typed Worker Report Columns" = Table.TransformColumnTypes(#"Promoted Worker Report Headers",{{"Employee Code", Int64.Type}, {"Period Description", type text}, {"Position Description", type text}, {"Employee Status Description", type text}, {"Start Date", type date}, {"Termination Date", type date}, {"Termination Reason Description", type text}, {"Age", type number}, {"Contracted FN Hours", type number}, {"Location Description", type text}, {"Site Description", type text}, {"Department Description", type text}})
in
    #"Typed Worker Report Columns";

// Query: CurrentWorkers_Prepare
// Purpose: Validate and normalize current-worker employment, termination and contract fields used by reconciliation.
shared CurrentWorkers_Prepare = let
    Source = #"IMPORT Current_Workers",

    RequiredColumns = {
        "Employee Code",
        "Location Description",
        "Position Description",
        "Employee Status Description",
        "Start Date",
        "Termination Date",
        "Termination Reason Description",
        "Contracted FN Hours"
    },

    AvailableColumns = Table.ColumnNames(Source),

    MissingColumns = List.Difference(
        RequiredColumns,
        AvailableColumns
    ),

    SelectedColumns =
        if List.IsEmpty(MissingColumns) then
            Table.SelectColumns(Source, RequiredColumns)
        else
            error Error.Record(
                "Current workers preparation",
                "Worker's Report.xlsx/Sheet1 is missing required columns.",
                [MissingColumns = MissingColumns]
            ),

    TypedColumns = Table.TransformColumnTypes(
        SelectedColumns,
        {
            {"Employee Code", type text},
            {"Location Description", type text},
            {"Position Description", type text},
            {"Employee Status Description", type text},
            {"Start Date", type date},
            {"Termination Date", type nullable date},
            {"Termination Reason Description", type text},
            {"Contracted FN Hours", type nullable number}
        }
    ),

    // The existing facility rule uses the first two characters
    // of the trimmed location description.
    AddedFacility = Table.AddColumn(
        TypedColumns,
        "Facility-Abbrev",
        each
            let
                RawLocation = [Location Description],
                TrimmedLocation =
                    if RawLocation = null then
                        null
                    else
                        Text.Trim(Text.From(RawLocation)),
                FacilityCode =
                    if TrimmedLocation = null or TrimmedLocation = "" then
                        null
                    else
                        Text.Start(TrimmedLocation, 2),
                UppercaseFacilityCode =
                    if FacilityCode = null then
                        null
                    else
                        Text.Upper(FacilityCode)
            in
                UppercaseFacilityCode,
        type nullable text
    ),

    TrimmedEmployeeCodes = Table.TransformColumns(
        AddedFacility,
        {
            {
                "Employee Code",
                each if _ = null then null else Text.Trim(Text.From(_)),
                type nullable text
            }
        }
    ),

    NormalisedEmployeeCodes = Table.ReplaceValue(
        TrimmedEmployeeCodes,
        "",
        null,
        Replacer.ReplaceValue,
        {"Employee Code"}
    ),

    TrimmedFacilityCodes = Table.TransformColumns(
        NormalisedEmployeeCodes,
        {
            {
                "Facility-Abbrev",
                each if _ = null then null else Text.Trim(Text.From(_)),
                type nullable text
            }
        }
    ),

    UppercaseFacilityCodes = Table.TransformColumns(
        TrimmedFacilityCodes,
        {
            {
                "Facility-Abbrev",
                each if _ = null then null else Text.Upper(_),
                type nullable text
            }
        }
    ),

    NormalisedFacilityCodes = Table.ReplaceValue(
        UppercaseFacilityCodes,
        "",
        null,
        Replacer.ReplaceValue,
        {"Facility-Abbrev"}
    ),

    TrimmedPositionDescriptions = Table.TransformColumns(
        NormalisedFacilityCodes,
        {
            {
                "Position Description",
                each if _ = null then null else Text.Trim(Text.From(_)),
                type nullable text
            }
        }
    ),

    NormalisedPositionDescriptions = Table.ReplaceValue(
        TrimmedPositionDescriptions,
        "",
        null,
        Replacer.ReplaceValue,
        {"Position Description"}
    ),

    TrimmedEmploymentStatuses = Table.TransformColumns(
        NormalisedPositionDescriptions,
        {
            {
                "Employee Status Description",
                each if _ = null then null else Text.Trim(Text.From(_)),
                type nullable text
            }
        }
    ),

    NormalisedEmploymentStatuses = Table.ReplaceValue(
        TrimmedEmploymentStatuses,
        "",
        null,
        Replacer.ReplaceValue,
        {"Employee Status Description"}
    ),

    // Keep all home locations so selected workers retain employee-ID matches for role, contract and termination details.
    BufferedWorkers = Table.Buffer(NormalisedEmploymentStatuses)
in
    BufferedWorkers;

// Query: CurrentWorkers_Reconciled
// Purpose: Select one deterministic worker-report record per employee before the membership join.
// Notes: Employment is employee-level. Availability at another site does not change the worker's home facility in the report.
shared CurrentWorkers_Reconciled = let
    Source = CurrentWorkers_Prepare,
    #"Filtered Valid Join Keys" = Table.SelectRows(Source, each [Employee Code] <> null),
    #"Added Open Employment Rank" = Table.AddColumn(#"Filtered Valid Join Keys", "Open Employment Rank", each if [Termination Date] = null then 0 else 1, Int64.Type),
    #"Added Roster Eligibility Rank" = Table.AddColumn(#"Added Open Employment Rank", "Roster Eligibility Rank", each
        if [Termination Date] = null or [Termination Date] >= RosterStartDate then 0 else 1, Int64.Type),
    #"Grouped Worker" = Table.Group(#"Added Roster Eligibility Rank", {"Employee Code"}, {
        {"Selected Current Worker", each
            let
                SortedWorkerCandidates = Table.Sort(_, {
                    {"Roster Eligibility Rank", Order.Ascending},
                    {"Open Employment Rank", Order.Ascending},
                    {"Termination Date", Order.Descending},
                    {"Start Date", Order.Descending},
                    {"Facility-Abbrev", Order.Ascending},
                    {"Position Description", Order.Ascending}
                }),
                SelectedWorker = Table.First(SortedWorkerCandidates)
            in
                SelectedWorker, type record},
        {"Current Worker Row Count", each Table.RowCount(_), Int64.Type},
        // Repeated source rows with the same active employment facts are not materially ambiguous.
        {"Current Worker Eligible Row Count", each
            let
                EligibleRows = Table.SelectRows(_, each [Termination Date] = null or [Termination Date] >= RosterStartDate),
                MaterialFields = Table.SelectColumns(EligibleRows,
                    {"Facility-Abbrev", "Position Description", "Employee Status Description", "Termination Date", "Contracted FN Hours"}),
                DistinctEligibleRows = Table.Distinct(MaterialFields)
            in
                Table.RowCount(DistinctEligibleRows), Int64.Type}
    }),
    #"Expanded Selected Current Worker" = Table.ExpandRecordColumn(#"Grouped Worker", "Selected Current Worker",
        {"Facility-Abbrev", "Location Description", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Termination Reason Description", "Contracted FN Hours"},
        {"Facility-Abbrev", "Location Description", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Termination Reason Description", "Contracted FN Hours"}),
    #"Buffered Reconciled Worker Output" = Table.Buffer(#"Expanded Selected Current Worker")
in
    #"Buffered Reconciled Worker Output";

// Query: LeaveBalances_DIAGNOSTICS
// Purpose: Expose incomplete leave-report identity and conflicting roles or locations at employee-ID grain.
// Inputs: LeaveBalances_Workers.
// Output: Scoped leave-source issues that block Employees-TABLE publication.
// Notes: LeaveBalances_Prepare has already excluded non-analysis sites and the two hard-coded CS sites.
shared LeaveBalances_DIAGNOSTICS = let
    Source = LeaveBalances_Workers,
    #"Kept Invalid Leave Identities" = Table.SelectRows(Source, each
        [Employee Code] = null
            or [#"Leave Balance Facility-Abbrev"] = null
            or [Leave Balance Identity Count] > 1),
    #"Named Facility Key" = Table.RenameColumns(#"Kept Invalid Leave Identities", {
        {"Leave Balance Facility-Abbrev", "Facility-Abbrev"}}),
    #"Added Issue" = Table.AddColumn(#"Named Facility Key", "Issue", each
        if [Employee Code] = null or [#"Facility-Abbrev"] = null then "Incomplete leave balance identity"
        else "Ambiguous leave balance identity", type text),
    #"Added Issue Details" = Table.AddColumn(#"Added Issue", "Details", each
        if [Employee Code] = null or [#"Facility-Abbrev"] = null then "Leave Table2 requires an employee code and facility."
        else Text.From([Leave Balance Identity Count]) & " distinct role/location identities share this employee ID.", type text),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Added Issue Details")
in
    #"Buffered Diagnostic Output";

// Query: WorkerRecordLocations
// Purpose: Group the Worker Record into the employee/facility pairs applicable to availability.
// Inputs: CurrentWorkers_Prepare, before analysis-facility filtering.
// Output: One row per full Employee Code and two-character Worker Facility.
// Notes: The small grouped lookup is buffered; applicability does not depend on roster-date HR reconciliation.
shared WorkerRecordLocations = let
    Source = CurrentWorkers_Prepare,
    #"Selected Worker Location Fields" = Table.SelectColumns(Source, {"Employee Code", "Facility-Abbrev"}),
    #"Named Worker Facility" = Table.RenameColumns(#"Selected Worker Location Fields", {{"Facility-Abbrev", "Worker Facility"}}),
    // Empty or incomplete keys cannot authorize availability rows.
    #"Filter LOCATIONS" = Table.SelectRows(#"Named Worker Facility", each
        [Employee Code] <> null and [Employee Code] <> "" and [Worker Facility] <> null and Text.Length([Worker Facility]) = 2),
    #"Grouped Worker Locations" = Table.Group(#"Filter LOCATIONS", {"Employee Code", "Worker Facility"}, {}),
    #"Buffered Worker Locations" = Table.Buffer(#"Grouped Worker Locations")
in
    #"Buffered Worker Locations";

// Query: Availabilities_LocationMatched
// Purpose: Prepare availability and retain rows matching a Worker Record employee/facility pair.
// Inputs: IMPORT Combined Availabilities and WorkerRecordLocations.
// Output: Matched availability rows, before analysis-facility filtering.
// Notes: Match full employee IDs and the first two location characters. Grouped lookup keys preserve every matching date/time row without multiplying it.
shared Availabilities_LocationMatched = let
    Source = #"IMPORT Combined Availabilities",
    #"Promoted Availability Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Named Availability Fields" = Table.RenameColumns(#"Promoted Availability Headers", {
        {"PayrollCode", "Payroll Code"}, {"DateTimeStart", "Date From"}, {"DateTimeEnd", "Date To"}, {"Comments", "Availability or Leave Reason"}}),
    #"Typed Availability Fields" = Table.TransformColumnTypes(#"Named Availability Fields", {
        {"Site Name", type text}, {"Employee Name", type text}, {"Payroll Code", type text},
        {"Date From", type datetime}, {"Date To", type datetime}, {"Availability or Leave Reason", type text}}),
    #"Cleared Invalid Match Keys" = Table.ReplaceErrorValues(#"Typed Availability Fields", {{"Payroll Code", null}, {"Site Name", null}}),
    CleanWorkerText = (value as nullable text) as nullable text =>
        let
            TrimmedValue = if value = null then null else Text.Trim(value),
            CleanValue = if TrimmedValue = "" then null else TrimmedValue
        in
            CleanValue,
    #"Normalized Worker Fields" = Table.TransformColumns(#"Cleared Invalid Match Keys", {
        {"Employee Name", CleanWorkerText, type nullable text},
        {"Payroll Code", CleanWorkerText, type nullable text}}),
    #"Added Facility Abbreviation" = Table.AddColumn(#"Normalized Worker Fields", "Facility-Abbrev", each
        let
            Site = if [Site Name] = null then null else Text.Trim([Site Name]),
            FacilityCode = if Site = null or Text.Length(Site) < 2 then null else Text.Upper(Text.Start(Site, 2))
        in
            FacilityCode, type nullable text),
    #"Added Availability Flag" = Table.AddColumn(#"Added Facility Abbreviation", "Available", each
        let
            Reason = if [Availability or Leave Reason] = null then "" else Text.Upper(Text.From([Availability or Leave Reason])),
            IsAvailable = if Text.Contains(Reason, "UNAVAIL") then false else Text.Contains(Reason, "AVAIL")
        in
            IsAvailable, type logical),
    // Join availability to the small grouped Worker Record lookup before applying facility filters.
    #"Joined Worker Locations" = Table.NestedJoin(#"Added Availability Flag", {"Payroll Code", "Facility-Abbrev"},
        WorkerRecordLocations, {"Employee Code", "Worker Facility"}, "Worker Record", JoinKind.LeftOuter),
    #"Expanded Worker Match" = Table.ExpandTableColumn(#"Joined Worker Locations", "Worker Record", {"Employee Code"}, {"Worker Record Match"}),
    #"Filter LOCATIONS" = Table.SelectRows(#"Expanded Worker Match", each [Worker Record Match] <> null),
    #"Removed Worker Match" = Table.RemoveColumns(#"Filter LOCATIONS", {"Worker Record Match"})
in
    #"Removed Worker Match";

// Query: Availabilities Prepare
// Purpose: Apply the existing analysis-facility filters after Worker Record employee/location matching.
// Inputs: Availabilities_LocationMatched.
// Output: Applicable availability/leave records with the existing prepared-column interface.
// Notes: Keep all matched date/time rows; buffer the reduced table for downstream worker-membership preparation.
shared #"Availabilities Prepare" = let
    Source = Availabilities_LocationMatched,
    #"Filter LOCATIONS" = Table.SelectRows(Source, each
        [Site Name] = "BD: Beaudesert" or [Site Name] = "JE: Jessie Hunt" or [Site Name] = "NR: Narrabri- Robert Young" or [Site Name] = "TE: Temora"),
    #"Buffered Applicable Availability" = Table.Buffer(#"Filter LOCATIONS")
in
    #"Buffered Applicable Availability";

// Query: Availabilities-Employees
// Purpose: Retain identified availability/leave worker records while excluding extraction error and system-account rows.
shared #"Availabilities-Employees" = let
    Source = #"Availabilities Prepare",
    #"Replaced Errors" = Table.ReplaceErrorValues(Source, {{"Payroll Code", "999999"}}),
    #"Filtered Error Code" = Table.SelectRows(#"Replaced Errors", each [Payroll Code] <> "999999"),
    // Humanforce Admin is a system account in the records extract, not an available worker.
    #"Filtered System Account" = Table.SelectRows(#"Filtered Error Code", each
        [Payroll Code] = null or Comparer.OrdinalIgnoreCase(Text.Trim([Payroll Code]), "HF") <> 0)
in
    #"Filtered System Account";

// Query: Availabilities-EmployeesLIST
// Purpose: Reduce availability membership to one worker/facility row while preserving its employee name and original site label.
// Output: One row per Payroll Code and Facility-AbbrevZ.
shared #"Availabilities-EmployeesLIST" = let
    Source = #"Availabilities-Employees",
    #"Selected Worker Fields" = Table.SelectColumns(Source,{"Payroll Code", "Facility-Abbrev", "Employee Name", "Site Name"}),
    #"Renamed Facility" = Table.RenameColumns(#"Selected Worker Fields",{{"Facility-Abbrev", "Facility-AbbrevZ"}}),
    #"Grouped Worker Facility" = Table.Group(#"Renamed Facility", {"Payroll Code", "Facility-AbbrevZ"}, {
        {"Records Employee Name", each
            let
                NonNullNames = List.RemoveNulls([Employee Name]),
                DistinctNames = List.Distinct(NonNullNames, Comparer.OrdinalIgnoreCase),
                Names = List.Sort(DistinctNames, Comparer.OrdinalIgnoreCase)
            in
                if List.IsEmpty(Names) then null else Names{0}, type nullable text},
        {"Availabilities Location", each
            let
                OriginalSites = List.RemoveNulls([Site Name]),
                SortedSites = List.Sort(OriginalSites, Comparer.Ordinal)
            in
                if List.IsEmpty(SortedSites) then null else SortedSites{0}, type nullable text}
    }),
    // Preserve source presence independently of whether the worker identifier is complete.
    #"Added Records Source Marker" = Table.AddColumn(#"Grouped Worker Facility", "In Records Source", each true, type logical),
    #"Buffered Availability Worker Output" = Table.Buffer(#"Added Records Source Marker")
in
    #"Buffered Availability Worker Output";

// Query: WorkerSourceAttributes_SUMMARY
// Purpose: List distinct full location/facility names for client checking.
// Inputs: Roster Prepare, Availabilities Prepare, CurrentWorkers_Prepare and LeaveBalances_Prepare before worker grouping.
// Output: Four independent sorted source lists; shorter columns are padded with nulls and rows do not imply source matches.
// Notes: Each entry is the full source location/facility name. Quoted values retain spaces; employee/payroll IDs are excluded.
shared WorkerSourceAttributes_SUMMARY = let
    // Match the configured wanted facilities even when one has no current roster or availability rows.
    // Use the shared membership scope here; the complete employee lookup remains available for enrichment.
    WantedFacilityCodes = AnalysisFacilityCodes,
    DisplayAttribute = (value as nullable text) as text =>
        if value = null then "(missing)" else """" & Text.Replace(value, """", """""") & """",
    SourceAttributeList = (source as table, locationColumn as text) as list =>
        let
            SelectedAttributes = Table.SelectColumns(source, {locationColumn}),
            NamedAttributes = Table.RenameColumns(SelectedAttributes, {{locationColumn, "Location / Facility"}}),
            TextAttributes = Table.TransformColumnTypes(NamedAttributes, {{"Location / Facility", type text}}),
            // Use the existing two-character facility convention for scope while preserving the full original label in the output.
            WithFacilityCode = Table.AddColumn(TextAttributes, "Facility Code", each
                let
                    FullLocation = [#"Location / Facility"],
                    TrimmedLocation = if FullLocation = null then null else Text.Trim(FullLocation),
                    FacilityCode = if TrimmedLocation = null or TrimmedLocation = "" then null
                        else Text.Upper(Text.Start(TrimmedLocation, 2))
                in
                    FacilityCode, type nullable text),
            #"Filter LOCATIONS" = Table.SelectRows(WithFacilityCode, each
                List.Contains(WantedFacilityCodes, [Facility Code], Comparer.Ordinal)),
            // Collapse repeated workers to site attributes while retaining distinct full labels within the same facility code.
            DistinctAttributes = Table.Distinct(#"Filter LOCATIONS", {"Location / Facility", "Facility Code"}),
            SortedAttributes = Table.Sort(DistinctAttributes, {
                {"Facility Code", Order.Ascending},
                {"Location / Facility", Order.Ascending}}),
            FormattedAttributes = Table.AddColumn(SortedAttributes, "Combination", each
                DisplayAttribute([#"Location / Facility"]), type text),
            OutputList = FormattedAttributes[Combination]
        in
            OutputList,
    RosterSource = #"Roster Prepare",
    RosterAttributes = SourceAttributeList(RosterSource, "Location"),
    AvailabilitiesSource = #"Availabilities Prepare",
    AvailabilitiesAttributes = SourceAttributeList(AvailabilitiesSource, "Site Name"),
    WorkerReportSource = CurrentWorkers_Prepare,
    WorkerReportAttributes = SourceAttributeList(WorkerReportSource, "Location Description"),
    // LeaveBalances_Prepare already excludes both hard-coded CS facilities; keep that exclusion before summary scoping.
    LeaveBalanceSource = LeaveBalances_Prepare,
    LeaveBalanceAttributes = SourceAttributeList(LeaveBalanceSource, "Facility"),
    Output = Table.FromColumns(
        {RosterAttributes, AvailabilitiesAttributes, WorkerReportAttributes, LeaveBalanceAttributes},
        type table [Roster = nullable text, Availabilities = nullable text, #"Worker's Report" = nullable text, #"Leave Balances" = nullable text])
in
    Output;

// Query: EmployeeSourceMembership
// Purpose: Reconcile roster and availability workers, then attach leave membership by employee ID and append leave-only IDs.
// Inputs: Availabilities-EmployeesLIST, RosteredEmployeesTABLE and LeaveBalances_Workers.
// Output: Worker/facility/role candidates with independent roster, availability and leave source markers.
// Notes: Roster/availability joins retain their facility key; leave membership is employee-level within the analysis scope.
shared EmployeeSourceMembership = let
    Source = #"Availabilities-EmployeesLIST",
    RosterSource = RosteredEmployeesTABLE,
    #"Joined Roster Membership" = Table.NestedJoin(Source, {"Payroll Code", "Facility-AbbrevZ"}, RosterSource, {"Employee Code", "Facility-Abbrev"}, "RosteredEmployeeShifts (2)", JoinKind.FullOuter),
    #"Expanded RosteredEmployeeShifts (2)" = Table.ExpandTableColumn(#"Joined Roster Membership", "RosteredEmployeeShifts (2)", {"Department", "Employee Roster Name", "Employment Type", "Employee Code", "Role", "Facility-Abbrev", "In Allocation Source", "Location"}, {"Department", "Employee Roster Name", "Employment Type", "Employee Code", "RoleX", "Facility-AbbrevX", "In Allocation Source", "Roster Location"}),
    #"Added In Records" = Table.AddColumn(#"Expanded RosteredEmployeeShifts (2)", "In Records", each [#"In Records Source"] = true, type logical),
    #"Added In Allocation" = Table.AddColumn(#"Added In Records", "In Allocation", each [#"In Allocation Source"] = true, type logical),
    #"Added Assigned Employee" = Table.AddColumn(#"Added In Allocation", "AssignedAvailID", each [Payroll Code] ?? [Employee Code], type nullable text),
    #"Added Join Facility" = Table.AddColumn(#"Added Assigned Employee", "Join Facility-Abbrev", each [#"Facility-AbbrevX"] ?? [#"Facility-AbbrevZ"], type nullable text),
    #"Selected Existing Employee IDs" = Table.SelectColumns(#"Added Join Facility", {"AssignedAvailID"}),
    #"Existing Employee IDs" = Table.Distinct(#"Selected Existing Employee IDs"),
    // A leave-report ID absent from both roster and availability contributes its own worker/facility row.
    #"Leave Only Workers" = Table.NestedJoin(LeaveBalances_Workers, {"Employee Code"}, #"Existing Employee IDs",
        {"AssignedAvailID"}, "Existing Employee Match", JoinKind.LeftAnti),
    #"Selected Leave Only Identity" = Table.SelectColumns(#"Leave Only Workers", {"Employee Code", "Leave Balance Facility-Abbrev"}),
    #"Named Leave Only Identity" = Table.RenameColumns(#"Selected Leave Only Identity", {
        {"Employee Code", "AssignedAvailID"}, {"Leave Balance Facility-Abbrev", "Join Facility-Abbrev"}}),
    #"Added Leave Only In Records" = Table.AddColumn(#"Named Leave Only Identity", "In Records", each false, type logical),
    #"Added Leave Only In Allocation" = Table.AddColumn(#"Added Leave Only In Records", "In Allocation", each false, type logical),
    #"Combined Worker Membership" = Table.Combine({#"Added Join Facility", #"Added Leave Only In Allocation"}),
    // Leave membership follows the complete Employee Code, even when its facility differs from the roster or availability site.
    #"Merged Leave Balances" = Table.NestedJoin(#"Combined Worker Membership", {"AssignedAvailID"},
        LeaveBalances_Workers, {"Employee Code"}, "Leave Balance Worker", JoinKind.LeftOuter),
    #"Expanded Leave Balances" = Table.ExpandTableColumn(#"Merged Leave Balances", "Leave Balance Worker",
        {"Leave Balance Role", "Leave Balance Location", "In Leave Balances Source"},
        {"Leave Balance Role", "Leave Balance Location", "In Leave Balances Source"}),
    #"Added In Leave Balances" = Table.AddColumn(#"Expanded Leave Balances", "In Leave Balances", each
        [In Leave Balances Source] = true, type logical),
    #"Added Worker Membership" = Table.AddColumn(#"Added In Leave Balances", "Worker Membership", each
        if [In Allocation] and [In Records] then "Both"
        else if [In Allocation] then "Allocation only"
        else if [In Records] then "Records only"
        else if [In Leave Balances] then "Leave balances only"
        else error "A reconciled worker must appear in Roster, Availabilities or Leave Balances.", type text),
    #"Buffered Source Membership Output" = Table.Buffer(#"Added Worker Membership")
in
    #"Buffered Source Membership Output";

// Query: Employees-ALL
// Purpose: Enrich roster, availability and leave membership with current-worker employment and termination details.
// Inputs: EmployeeSourceMembership, CurrentWorkers_Reconciled and RosterStartDate.
// Output: Auditable employee/facility/role candidates before termination filtering.
shared #"Employees-ALL" = let
    RosterStart = RosterStartDate,
    Source = EmployeeSourceMembership,
    // Employment status is employee-level; the published facility remains the roster/availability or leave-only facility.
    #"Merged Current Workers" = Table.NestedJoin(Source, {"AssignedAvailID"}, CurrentWorkers_Reconciled, {"Employee Code"}, "Current Worker", JoinKind.LeftOuter),
    #"Expanded Current Worker" = Table.ExpandTableColumn(#"Merged Current Workers", "Current Worker",
        {"Employee Code", "Facility-Abbrev", "Position Description", "Employee Status Description", "Contracted FN Hours", "Start Date", "Termination Date", "Termination Reason Description", "Current Worker Row Count", "Current Worker Eligible Row Count", "Location Description"},
        {"Employee Code.1", "Facility-AbbrevY", "Position Description", "Employee Status Description", "Contracted FN Hours", "Start Date", "Termination Date", "Termination Reason Description", "Current Worker Row Count", "Current Worker Eligible Row Count", "Records Location"}),
    #"Added Current Worker Match Missing" = Table.AddColumn(#"Expanded Current Worker", "Current Worker Match Missing", each [Employee Code.1] = null, type logical),
    #"Added Current Worker Record Ambiguous" = Table.AddColumn(#"Added Current Worker Match Missing", "Current Worker Record Ambiguous", each
        [Current Worker Eligible Row Count] <> null and [Current Worker Eligible Row Count] > 1, type logical),
    #"Added Terminated Before Roster" = Table.AddColumn(#"Added Current Worker Record Ambiguous", "Terminated Before Roster", each
        not [Current Worker Match Missing] and [Termination Date] <> null and [Termination Date] < RosterStart, type logical)
in
    #"Added Terminated Before Roster";

// Query: EmployeesTABLE_Prepare
// Purpose: Prepare active reconciled workers with source membership, original locations and one preferred role.
// Output: One preferred worker row per employee/facility.
shared EmployeesTABLE_Prepare = let
    Source = #"Employees-ALL",
    #"Filtered Terminated Workers" = Table.SelectRows(Source, each not [Terminated Before Roster]),
    #"Renamed Allocation Name" = Table.RenameColumns(#"Filtered Terminated Workers",{{"Employee Roster Name", "Allocation Employee Name"}}),
    #"Added Employee ID" = Table.AddColumn(#"Renamed Allocation Name", "EmployeeID", each [AssignedAvailID] ?? [Employee Code.1], type nullable text),
    #"Added Facility" = Table.AddColumn(#"Added Employee ID", "Facility-Abbrev", each [#"Join Facility-Abbrev"] ?? [#"Facility-AbbrevY"], type nullable text),
    #"Added Employee Name" = Table.AddColumn(#"Added Facility", "Employee Roster Name", each [Allocation Employee Name] ?? [Records Employee Name], type nullable text),
    #"Added Employment Type" = Table.AddColumn(#"Added Employee Name", "EmploymentType", each [Employment Type] ?? [Employee Status Description], type nullable text),
    // Retain the existing role priority; leave Position Description supplies Role only when both earlier sources are missing.
    #"Added Preferred Role Candidate" = Table.AddColumn(#"Added Employment Type", "Role", each [RoleX] ?? [Position Description] ?? [Leave Balance Role], type nullable text),
    #"Selected Worker Fields" = Table.SelectColumns(#"Added Preferred Role Candidate",{
        "Allocation Employee Name", "Records Employee Name", "Employee Roster Name", "EmployeeID", "Facility-Abbrev", "EmploymentType", "Role",
        "Records Location", "Roster Location", "Availabilities Location", "Leave Balance Location",
        "In Allocation", "In Records", "In Leave Balances", "Worker Membership", "Current Worker Match Missing", "Current Worker Record Ambiguous",
        "Current Worker Row Count", "Current Worker Eligible Row Count", "Position Description", "Contracted FN Hours", "Start Date", "Termination Date", "Terminated Before Roster"}),
    #"Sorted Employee Candidates" = Table.Sort(#"Selected Worker Fields",{{"EmployeeID", Order.Ascending}}),
    #"Normalized Employment Type Separators" = Table.ReplaceValue(#"Sorted Employee Candidates"," ","-",Replacer.ReplaceText,{"EmploymentType"}),
    // Prefer an allocation role that agrees with the worker's current position; otherwise retain a deterministic source role.
    RoleAlignmentRank = (role as nullable text, position as nullable text) as number =>
        let
            RoleKey = if role = null then null else Text.Upper(Text.Trim(role)),
            PositionKey = if position = null then null else Text.Upper(Text.Trim(position)),
            Rank = if RoleKey = null then 9
                else if PositionKey = null then 2
                else if RoleKey = PositionKey then 0
                else if PositionKey = "REGISTERED NURSE" and Text.StartsWith(RoleKey, "REGN") then 0
                else if PositionKey = "ASSISTANT IN NURSING" and RoleKey = "ASSISTANT IN NURSING" then 0
                else if PositionKey = "ENROLLED NURSE" and RoleKey = "ENROLLED NURSE" then 0
                else if PositionKey = "ASSISTANT IN NURSING" and Text.Contains(RoleKey, "ASSISTANT IN NURSING") then 1
                else 2
        in
            Rank,
    #"Added Role Alignment Rank" = Table.AddColumn(#"Normalized Employment Type Separators", "Role Alignment Rank", each
        RoleAlignmentRank([Role], [Position Description]), Int64.Type),
    CandidateWorkerColumns = Table.ColumnNames(#"Added Role Alignment Rank"),
    PreferredWorkerColumns = List.RemoveItems(CandidateWorkerColumns, {"EmployeeID", "Facility-Abbrev"}),
    #"Grouped Preferred Workers" = Table.Group(#"Added Role Alignment Rank", {"EmployeeID", "Facility-Abbrev"}, {
        {"Preferred Worker", each
            let
                SortedRoleCandidates = Table.Sort(_, {
                    {"Role Alignment Rank", Order.Ascending},
                    {"Role", Order.Ascending},
                    {"Employee Roster Name", Order.Ascending},
                    {"EmploymentType", Order.Ascending}
                }),
                PreferredWorker = Table.First(SortedRoleCandidates)
            in
                PreferredWorker, type record}
    }),
    #"Expanded Preferred Worker" = Table.ExpandRecordColumn(#"Grouped Preferred Workers", "Preferred Worker", PreferredWorkerColumns, PreferredWorkerColumns),
    #"Removed Preference Fields" = Table.RemoveColumns(#"Expanded Preferred Worker",{"Role Alignment Rank", "Position Description"}),
    #"Buffered Preferred Worker Output" = Table.Buffer(#"Removed Preference Fields")
in
    #"Buffered Preferred Worker Output";

// Query: LeaveBalances_NotInEmployees_DIAGNOSTICS
// Purpose: Show leave-report IDs absent after employee eligibility and termination checks.
// Inputs: LeaveBalances_Workers and EmployeesTABLE_Prepare.
// Output: Leave workers whose employee ID is absent from Employees-TABLE preparation.
shared LeaveBalances_NotInEmployees_DIAGNOSTICS = let
    Source = LeaveBalances_Workers,
    EmployeeSource = EmployeesTABLE_Prepare,
    #"Selected Employee Keys" = Table.SelectColumns(EmployeeSource, {"EmployeeID"}),
    // Compare by employee ID because one leave identity can already be represented at an analysis facility.
    #"Kept Leave Workers Without Employees" = Table.NestedJoin(
        Source, {"Employee Code"},
        #"Selected Employee Keys", {"EmployeeID"},
        "Employee Match", JoinKind.LeftAnti),
    #"Removed Match Column" = Table.RemoveColumns(#"Kept Leave Workers Without Employees", {"Employee Match"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Removed Match Column")
in
    #"Buffered Diagnostic Output";

// Query: TerminatedWorkers_DIAGNOSTICS
// Purpose: Show reconciled candidates excluded because their termination date precedes the roster start.
// Inputs: Employees-ALL, before terminated candidates are excluded from publication.
// Output: Distinct excluded employee/facility identities and termination details.
shared TerminatedWorkers_DIAGNOSTICS = let
    Source = #"Employees-ALL",
    #"Kept Workers Terminated Before Roster" = Table.SelectRows(Source, each [Terminated Before Roster]),
    #"Added Employee ID" = Table.AddColumn(#"Kept Workers Terminated Before Roster", "EmployeeID", each
        [AssignedAvailID] ?? [Employee Code.1], type nullable text),
    #"Added Facility Key" = Table.AddColumn(#"Added Employee ID", "Facility-Abbrev", each
        [#"Join Facility-Abbrev"] ?? [#"Facility-AbbrevY"], type nullable text),
    #"Named Allocation Employee Name" = Table.RenameColumns(#"Added Facility Key", {
        {"Employee Roster Name", "Allocation Employee Name"}}),
    #"Added Employee Name" = Table.AddColumn(#"Named Allocation Employee Name", "Employee Roster Name", each
        [Allocation Employee Name] ?? [Records Employee Name], type nullable text),
    #"Selected Termination Details" = Table.SelectColumns(#"Added Employee Name", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Termination Date",
        "Termination Reason Description", "In Allocation", "In Records", "Worker Membership"}),
    #"Distinct Terminated Workers" = Table.Distinct(#"Selected Termination Details"),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Distinct Terminated Workers")
in
    #"Buffered Diagnostic Output";

// Query: DuplicateWorkers_DIAGNOSTICS
// Purpose: Show every employee/facility key that still has more than one preferred candidate row.
// Inputs: EmployeesTABLE_Prepare.
// Output: Duplicate employee/facility keys with candidate counts and distinct identity values.
shared DuplicateWorkers_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    DistinctText = (values as list) as text =>
        let
            TextValues = List.Transform(values, each Text.From(_)),
            NonNullValues = List.RemoveNulls(TextValues),
            DistinctValues = List.Distinct(NonNullValues, Comparer.OrdinalIgnoreCase),
            SortedValues = List.Sort(DistinctValues, Comparer.OrdinalIgnoreCase),
            DisplayText = Text.Combine(SortedValues, " | ")
        in
            DisplayText,
    #"Grouped Employee Facility Candidates" = Table.Group(Source, {"EmployeeID", "Facility-Abbrev"}, {
        {"Candidate Count", each Table.RowCount(_), Int64.Type},
        {"Employee Roster Name", each DistinctText([Employee Roster Name]), type text},
        {"Role", each DistinctText([Role]), type text},
        {"Worker Membership", each DistinctText([Worker Membership]), type text}
    }),
    #"Kept Duplicate Worker Keys" = Table.SelectRows(#"Grouped Employee Facility Candidates", each [Candidate Count] > 1),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Kept Duplicate Worker Keys")
in
    #"Buffered Diagnostic Output";

// Query: IncompleteWorkerIdentity_DIAGNOSTICS
// Purpose: Show retained candidates missing required identity, allowing identifier-only names for leave-only workers.
// Inputs: EmployeesTABLE_Prepare.
// Output: Incomplete employee/facility identities with the names of their missing required fields.
// Notes: The approved leave source has no employee name; leave-only rows may retain a null Employee Roster Name.
// Internal Records only membership means availability without roster membership; it does not confirm a Worker's Report match.
shared IncompleteWorkerIdentity_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    IsBlank = (value as any) as logical =>
        let
            BlankValue = value = null or (Value.Is(value, type text) and Text.Trim(value) = "")
        in
            BlankValue,
    MissingRequiredName = (worker as record) as logical =>
        let
            MissingName = IsBlank(worker[Employee Roster Name]),
            LeaveOnlyWorker = worker[In Leave Balances] and not worker[In Allocation] and not worker[In Records],
            RequiredNameMissing = MissingName and not LeaveOnlyWorker
        in
            RequiredNameMissing,
    #"Kept Incomplete Worker Identities" = Table.SelectRows(Source, each
        IsBlank([EmployeeID])
            or IsBlank([#"Facility-Abbrev"])
            or MissingRequiredName(_)
            or IsBlank([Role])),
    #"Added Missing Field Names" = Table.AddColumn(#"Kept Incomplete Worker Identities", "Missing Fields", each
        let
            FieldLabels = {
                if IsBlank([EmployeeID]) then "EmployeeID" else null,
                if IsBlank([#"Facility-Abbrev"]) then "Facility-Abbrev" else null,
                if MissingRequiredName(_) then "Employee Roster Name" else null,
                if IsBlank([Role]) then "Role" else null
            },
            MissingLabels = List.RemoveNulls(FieldLabels),
            MissingFieldText = Text.Combine(MissingLabels, ", ")
        in
            MissingFieldText, type text),
    #"Selected Identity Diagnostic Columns" = Table.SelectColumns(#"Added Missing Field Names", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Missing Fields"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Selected Identity Diagnostic Columns")
in
    #"Buffered Diagnostic Output";

// Query: UnmatchedCurrentWorkers_DIAGNOSTICS
// Purpose: Show membership candidates that could not be matched to the worker report by employee ID.
// Inputs: EmployeesTABLE_Prepare.
// Output: Unmatched employee/facility identities and the sources that provide their membership.
// Notes: Roster or leave membership can supply Role without a worker-report match; unmatched contracts remain null.
shared UnmatchedCurrentWorkers_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    #"Kept Missing Worker Report Matches" = Table.SelectRows(Source, each [Current Worker Match Missing]),
    #"Selected Unmatched Worker Columns" = Table.SelectColumns(#"Kept Missing Worker Report Matches", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership",
        "In Allocation", "In Records", "In Leave Balances"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Selected Unmatched Worker Columns")
in
    #"Buffered Diagnostic Output";

// Query: AmbiguousCurrentWorkers_DIAGNOSTICS
// Purpose: Show workers with more than one worker-report row eligible at roster start.
// Inputs: EmployeesTABLE_Prepare.
// Output: Ambiguous employee/facility identities with total and eligible worker-report row counts.
shared AmbiguousCurrentWorkers_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    #"Kept Ambiguous Worker Report Matches" = Table.SelectRows(Source, each [Current Worker Record Ambiguous]),
    #"Selected Ambiguous Worker Columns" = Table.SelectColumns(#"Kept Ambiguous Worker Report Matches", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership",
        "Current Worker Row Count", "Current Worker Eligible Row Count"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Selected Ambiguous Worker Columns")
in
    #"Buffered Diagnostic Output";

// Query: InvalidWorkerMembership_DIAGNOSTICS
// Purpose: Show candidates missing all roster, availability and leave membership or carrying an invalid internal label.
// Inputs: EmployeesTABLE_Prepare.
// Output: Invalid employee/facility membership labels and the three source-membership flags.
// Notes: Internal Records membership means Availabilities, not a Worker's Report match.
shared InvalidWorkerMembership_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    ValidMembershipValues = {"Allocation only", "Records only", "Both", "Leave balances only"},
    #"Kept Invalid Worker Memberships" = Table.SelectRows(Source, each
        (not [In Allocation] and not [In Records] and not [In Leave Balances])
            or not List.Contains(ValidMembershipValues, [Worker Membership])),
    #"Selected Membership Diagnostic Columns" = Table.SelectColumns(#"Kept Invalid Worker Memberships", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership",
        "In Allocation", "In Records", "In Leave Balances"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Selected Membership Diagnostic Columns")
in
    #"Buffered Diagnostic Output";

// Query: LeakedTerminatedWorkers_DIAGNOSTICS
// Purpose: Independently anti-check that no excluded terminated worker/facility key remains in the publishable population.
// Inputs: EmployeesTABLE_Prepare and TerminatedWorkers_DIAGNOSTICS.
// Output: Any retained employee/facility row also identified as excluded by termination, with its excluded date.
shared LeakedTerminatedWorkers_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    TerminatedSource = TerminatedWorkers_DIAGNOSTICS,
    #"Selected Terminated Key Columns" = Table.SelectColumns(TerminatedSource, {
        "EmployeeID", "Facility-Abbrev", "Termination Date"}),
    #"Distinct Terminated Keys" = Table.Distinct(#"Selected Terminated Key Columns"),
    // The employee/facility grain must match both the excluded and publishable populations.
    #"Matched Excluded Terminated Keys" = Table.NestedJoin(
        Source, {"EmployeeID", "Facility-Abbrev"},
        #"Distinct Terminated Keys", {"EmployeeID", "Facility-Abbrev"},
        "Terminated Match", JoinKind.Inner),
    #"Expanded Excluded Termination Date" = Table.ExpandTableColumn(#"Matched Excluded Terminated Keys",
        "Terminated Match", {"Termination Date"}, {"Excluded Termination Date"}),
    #"Buffered Diagnostic Output" = Table.Buffer(#"Expanded Excluded Termination Date")
in
    #"Buffered Diagnostic Output";

// Query: OutsideAnalysisFacilities_DIAGNOSTICS
// Purpose: Expose any retained worker or membership location outside the approved analysis facilities.
// Inputs: EmployeesTABLE_Prepare and AnalysisFacilityCodes.
// Output: Blocking employee/facility rows with the offending membership fields and original location labels.
// Notes: Records home location is enrichment context and is deliberately excluded from analysis-membership scope checks.
shared OutsideAnalysisFacilities_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    IsAnalysisLocation = (location as nullable text) as logical =>
        let
            TrimmedLocation = if location = null then null else Text.Trim(location),
            FacilityCode = if TrimmedLocation = null or TrimmedLocation = "" then null
                else Text.Upper(Text.Start(TrimmedLocation, 2)),
            IsAllowed = List.Contains(AnalysisFacilityCodes, FacilityCode, Comparer.Ordinal)
        in
            IsAllowed,
    #"Checked Worker Facility" = Table.AddColumn(Source, "Analysis Facility Valid", each
        List.Contains(AnalysisFacilityCodes, [#"Facility-Abbrev"], Comparer.Ordinal), type logical),
    #"Checked Roster Location" = Table.AddColumn(#"Checked Worker Facility", "Roster Location Valid", each
        if [In Allocation] then IsAnalysisLocation([Roster Location]) else true, type logical),
    #"Checked Availability Location" = Table.AddColumn(#"Checked Roster Location", "Availability Location Valid", each
        if [In Records] then IsAnalysisLocation([Availabilities Location]) else true, type logical),
    #"Checked Leave Location" = Table.AddColumn(#"Checked Availability Location", "Leave Location Valid", each
        if [In Leave Balances] then IsAnalysisLocation([Leave Balance Location]) else true, type logical),
    #"Added Scope Details" = Table.AddColumn(#"Checked Leave Location", "Outside Analysis Scope", each
        let
            FieldLabels = {
                if [Analysis Facility Valid] then null else "Facility-Abbrev",
                if [Roster Location Valid] then null else "Roster Location",
                if [Availability Location Valid] then null else "Availabilities Location",
                if [Leave Location Valid] then null else "Leave Balance Location"
            },
            InvalidFields = List.RemoveNulls(FieldLabels),
            Details = Text.Combine(InvalidFields, ", ")
        in
            Details, type text),
    // This diagnostic keeps location-scope failures; Employees-TABLE must not publish them.
    #"Filter LOCATIONS" = Table.SelectRows(#"Added Scope Details", each [Outside Analysis Scope] <> ""),
    #"Selected Scope Diagnostic Columns" = Table.SelectColumns(#"Filter LOCATIONS", {
        "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership",
        "Roster Location", "Availabilities Location", "Leave Balance Location", "Outside Analysis Scope"}),
    Output = Table.Buffer(#"Selected Scope Diagnostic Columns")
in
    Output;

// Query: WorkerReconciliation_ISSUES
// Purpose: Consolidate every blocking reconciliation issue into one row-level review table.
// Inputs: Worker identity, membership, employment and analysis-scope diagnostics.
// Output: Consistent issue, employee/facility identity, membership and details columns for blocking cases.
shared WorkerReconciliation_ISSUES = let
    Source = DuplicateWorkers_DIAGNOSTICS,
    IssueColumns = {"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"},
    #"Named Duplicate Issues" = Table.AddColumn(Source, "Issue", each "Duplicate employee/facility", type text),
    #"Added Duplicate Details" = Table.AddColumn(#"Named Duplicate Issues", "Details", each
        Text.From([Candidate Count]) & " candidate rows remain for this key.", type text),
    #"Selected Duplicate Issues" = Table.SelectColumns(#"Added Duplicate Details", IssueColumns),

    IdentitySource = IncompleteWorkerIdentity_DIAGNOSTICS,
    #"Named Identity Issues" = Table.AddColumn(IdentitySource, "Issue", each "Incomplete worker identity", type text),
    #"Added Identity Details" = Table.AddColumn(#"Named Identity Issues", "Details", each "Missing: " & [Missing Fields], type text),
    #"Selected Identity Issues" = Table.SelectColumns(#"Added Identity Details", IssueColumns),

    MembershipSource = InvalidWorkerMembership_DIAGNOSTICS,
    #"Named Membership Issues" = Table.AddColumn(MembershipSource, "Issue", each "Invalid worker membership", type text),
    #"Added Membership Details" = Table.AddColumn(#"Named Membership Issues", "Details", each
        "Roster, availability or leave source membership is missing or invalid.", type text),
    #"Selected Membership Issues" = Table.SelectColumns(#"Added Membership Details", IssueColumns),

    UnmatchedSource = UnmatchedCurrentWorkers_DIAGNOSTICS,
    // Roster and leave sources can supply Role independently; availability-only workers still require worker-report enrichment.
    #"Kept Blocking Unmatched Workers" = Table.SelectRows(UnmatchedSource, each not [In Allocation] and not [In Leave Balances]),
    #"Named Unmatched Issues" = Table.AddColumn(#"Kept Blocking Unmatched Workers", "Issue", each "Current worker match missing", type text),
    #"Added Unmatched Details" = Table.AddColumn(#"Named Unmatched Issues", "Details", each
        "No Worker's Report row matched this availability-only employee ID.", type text),
    #"Selected Unmatched Issues" = Table.SelectColumns(#"Added Unmatched Details", IssueColumns),

    AmbiguousSource = AmbiguousCurrentWorkers_DIAGNOSTICS,
    #"Named Ambiguous Issues" = Table.AddColumn(AmbiguousSource, "Issue", each "Current worker record ambiguous", type text),
    #"Added Ambiguous Details" = Table.AddColumn(#"Named Ambiguous Issues", "Details", each
        Text.From([Current Worker Eligible Row Count]) & " Worker's Report rows are eligible at roster start for this employee.", type text),
    #"Selected Ambiguous Issues" = Table.SelectColumns(#"Added Ambiguous Details", IssueColumns),

    TerminatedSource = LeakedTerminatedWorkers_DIAGNOSTICS,
    #"Named Termination Issues" = Table.AddColumn(TerminatedSource, "Issue", each "Terminated worker leaked", type text),
    #"Added Termination Details" = Table.AddColumn(#"Named Termination Issues", "Details", each
        "Excluded termination date: " & Date.ToText([Excluded Termination Date], "yyyy-MM-dd"), type text),
    #"Selected Termination Issues" = Table.SelectColumns(#"Added Termination Details", IssueColumns),

    LeaveSource = LeaveBalances_DIAGNOSTICS,
    #"Named Leave Identity Columns" = Table.RenameColumns(LeaveSource, {
        {"Employee Code", "EmployeeID"}, {"Leave Balance Role", "Role"}}),
    #"Added Leave Employee Name" = Table.AddColumn(#"Named Leave Identity Columns", "Employee Roster Name", each null, type nullable text),
    #"Added Leave Membership Label" = Table.AddColumn(#"Added Leave Employee Name", "Worker Membership", each null, type nullable text),
    #"Selected Leave Issues" = Table.SelectColumns(#"Added Leave Membership Label", IssueColumns),

    ScopeSource = OutsideAnalysisFacilities_DIAGNOSTICS,
    #"Named Scope Issues" = Table.AddColumn(ScopeSource, "Issue", each "Outside analysis facilities", type text),
    #"Added Scope Issue Details" = Table.AddColumn(#"Named Scope Issues", "Details", each
        "Non-analysis membership fields: " & [Outside Analysis Scope], type text),
    #"Selected Scope Issues" = Table.SelectColumns(#"Added Scope Issue Details", IssueColumns),

    #"Combined Blocking Issues" = Table.Combine({
        #"Selected Duplicate Issues", #"Selected Identity Issues", #"Selected Membership Issues",
        #"Selected Unmatched Issues", #"Selected Ambiguous Issues", #"Selected Termination Issues", #"Selected Leave Issues", #"Selected Scope Issues"}),
    #"Buffered Issue Output" = Table.Buffer(#"Combined Blocking Issues")
in
    #"Buffered Issue Output";

// Query: WorkerReconciliation_CHECK
// Purpose: Block publication for identity, membership, employment and analysis-location failures.
// Inputs: Existing blocking worker diagnostics.
// Output: One status, failure count and explanation for each required publication check.
shared WorkerReconciliation_CHECK = let
    Source = DuplicateWorkers_DIAGNOSTICS,
    DuplicateWorkerCount = Table.RowCount(Source),
    IdentitySource = IncompleteWorkerIdentity_DIAGNOSTICS,
    IncompleteIdentityCount = Table.RowCount(IdentitySource),
    MembershipSource = InvalidWorkerMembership_DIAGNOSTICS,
    InvalidMembershipCount = Table.RowCount(MembershipSource),
    UnmatchedSource = UnmatchedCurrentWorkers_DIAGNOSTICS,
    // Match requirements are the same as the blocking unmatched-worker rule in WorkerReconciliation_ISSUES.
    #"Kept Blocking Unmatched Workers" = Table.SelectRows(UnmatchedSource, each not [In Allocation] and not [In Leave Balances]),
    BlockingUnmatchedCount = Table.RowCount(#"Kept Blocking Unmatched Workers"),
    AmbiguousSource = AmbiguousCurrentWorkers_DIAGNOSTICS,
    AmbiguousWorkerCount = Table.RowCount(AmbiguousSource),
    TerminatedSource = LeakedTerminatedWorkers_DIAGNOSTICS,
    LeakedTerminatedCount = Table.RowCount(TerminatedSource),
    LeaveSource = LeaveBalances_DIAGNOSTICS,
    InvalidLeaveIdentityCount = Table.RowCount(LeaveSource),
    ScopeSource = OutsideAnalysisFacilities_DIAGNOSTICS,
    OutsideAnalysisFacilityCount = Table.RowCount(ScopeSource),
    CheckResult = (name as text, failures as number, details as text) as record =>
        let
            CheckStatus = if failures = 0 then "Pass" else "Fail",
            CheckRecord = [Check = name, Status = CheckStatus, Failures = failures, Details = details]
        in
            CheckRecord,
    UniqueWorkerCheck = CheckResult(
        "Unique employee and facility workers", DuplicateWorkerCount,
        "Employees_TABLE must publish one preferred row per employee and facility."),
    CompleteIdentityCheck = CheckResult(
        "Complete worker identity", IncompleteIdentityCount,
        "EmployeeID, facility and role are required; a name is optional only for leave-only workers."),
    ValidMembershipCheck = CheckResult(
        "Valid worker membership", InvalidMembershipCount,
        "Each worker must have roster, availability or leave membership."),
    CompleteWorkerMatchCheck = CheckResult(
        "Current worker match complete", BlockingUnmatchedCount,
        "Availability-only workers require a Worker's Report match; roster or leave workers may supply their identity and role independently."),
    UnambiguousWorkerCheck = CheckResult(
        "Current worker record unambiguous", AmbiguousWorkerCount,
        "Each employee may have at most one materially distinct Worker's Report row eligible at roster start."),
    NoTerminatedWorkersCheck = CheckResult(
        "No terminated workers published", LeakedTerminatedCount,
        "Workers terminated before roster start must remain diagnostic-only."),
    ValidLeaveIdentityCheck = CheckResult(
        "Valid leave balance source identity", InvalidLeaveIdentityCount,
        "Leave Table2 must have complete employee/facility identity and one unambiguous role/location identity per employee ID."),
    AnalysisFacilityCheck = CheckResult(
        "Only analysis facilities published", OutsideAnalysisFacilityCount,
        "Worker facility and roster, availability and leave membership locations must belong to AnalysisFacilityCodes; Records home location is enrichment context."),
    RequiredChecks = {
        UniqueWorkerCheck, CompleteIdentityCheck, ValidMembershipCheck, CompleteWorkerMatchCheck,
        UnambiguousWorkerCheck, NoTerminatedWorkersCheck, ValidLeaveIdentityCheck, AnalysisFacilityCheck},
    #"Created Check Table" = Table.FromRecords(RequiredChecks,
        type table [Check = text, Status = text, Failures = number, Details = text]),
    #"Buffered Check Output" = Table.Buffer(#"Created Check Table")
in
    #"Buffered Check Output";

// Query: Employees-TABLE
// Purpose: Publish validated analysis-site workers with source membership and separate employee home-location context.
// Inputs: EmployeesTABLE_Prepare and WorkerReconciliation_CHECK.
// Output: One employee/facility row; Location compares analysis memberships, Records Home Location retains the full HR label, Contracted FN Hours is last.
// Notes: Source includes the Worker's Report as Records; Location and Locations Match compare only roster, availability and leave analysis memberships.
shared #"Employees-TABLE" = let
    Source = EmployeesTABLE_Prepare,
    CheckSource = WorkerReconciliation_CHECK,
    #"Kept Failed Publication Checks" = Table.SelectRows(CheckSource, each [Status] = "Fail"),
    #"Validated Employee Population" = if Table.IsEmpty(#"Kept Failed Publication Checks") then Source
        else error Error.Record("Worker reconciliation validation", "Required worker checks failed.", #"Kept Failed Publication Checks"),
    // Internal Records membership refers to Availabilities; the published Records flag below means an employee-report match.
    #"Named Source Membership Flags" = Table.RenameColumns(#"Validated Employee Population", {
        {"In Records", "In Availabilities"},
        {"In Allocation", "In Roster"}}),
    #"Added In Records" = Table.AddColumn(#"Named Source Membership Flags", "In Records", each
        not [Current Worker Match Missing], type logical),
    #"Added Source" = Table.AddColumn(#"Added In Records", "Source", each
        let
            SourceLabels = {
                if [In Records] then "Records" else null,
                if [In Roster] then "Roster" else null,
                if [In Availabilities] then "Availabilities" else null,
                if [In Leave Balances] then "Leave Balances" else null
            },
            PresentSources = List.RemoveNulls(SourceLabels),
            SourceText = Text.Combine(PresentSources, ", ")
        in
            SourceText, type text),
    // The complete HR lookup supplies Role, contracts and employment dates without making its home site an analysis membership.
    #"Added Records Home Location" = Table.AddColumn(#"Added Source", "Records Home Location", each
        [Records Location], type nullable text),
    LocationPrefix = (location as nullable text) as nullable text =>
        let
            TrimmedLocation = if location = null then null else Text.Trim(location),
            Prefix = if TrimmedLocation = null or TrimmedLocation = "" then null else Text.Start(TrimmedLocation, 6)
        in
            Prefix,
    #"Added Analysis Location Codes" = Table.AddColumn(#"Added Records Home Location", "Analysis Location Codes", each
        let
            RosterCodes = if [In Roster] then {LocationPrefix([Roster Location])} else {},
            AvailabilityCodes = if [In Availabilities] then {LocationPrefix([Availabilities Location])} else {},
            LeaveCodes = if [In Leave Balances] then {LocationPrefix([Leave Balance Location])} else {},
            AnalysisCodes = List.Combine({RosterCodes, AvailabilityCodes, LeaveCodes})
        in
            AnalysisCodes, type list),
    #"Added Location" = Table.AddColumn(#"Added Analysis Location Codes", "Location", each
        let
            DisplayCodes = List.Transform([Analysis Location Codes], each _ ?? "(missing)"),
            LocationText = Text.Combine(DisplayCodes, ", ")
        in
            LocationText, type text),
    // Agreement compares analysis-site memberships only; an employee's different home site does not create a mismatch.
    #"Added Locations Match" = Table.AddColumn(#"Added Location", "Locations Match", each
        let
            AnalysisCodes = [Analysis Location Codes],
            SourceCount = List.Count(AnalysisCodes),
            CompleteLocations = List.NonNullCount(AnalysisCodes) = SourceCount,
            DistinctCodes = List.Distinct(AnalysisCodes, Comparer.Ordinal),
            LocationsAgree = SourceCount > 0 and CompleteLocations and List.Count(DistinctCodes) = 1
        in
            LocationsAgree, type logical),
    #"Removed Internal Fields" = Table.RemoveColumns(#"Added Locations Match", {
        "Records Location", "Roster Location", "Availabilities Location", "Leave Balance Location", "Analysis Location Codes",
        "Allocation Employee Name", "Records Employee Name", "Worker Membership", "Current Worker Match Missing", "Current Worker Record Ambiguous",
        "Current Worker Row Count", "Current Worker Eligible Row Count", "Terminated Before Roster"}),
    VisibleColumns = Table.ColumnNames(#"Removed Internal Fields"),
    OtherOutputColumns = List.RemoveItems(VisibleColumns, {"Contracted FN Hours"}),
    OutputColumnOrder = OtherOutputColumns & {"Contracted FN Hours"},
    #"Ordered Employee Columns" = Table.ReorderColumns(#"Removed Internal Fields", OutputColumnOrder),
    Output = Table.Buffer(#"Ordered Employee Columns")
in
    Output;

// Query: Employees-MergedAvail+Roster
// Purpose: Reconcile availability and roster rows for the legacy membership summary.
// Inputs: Availabilities-EmployeesLIST and RosteredEmployeesTABLE.
// Output: Full-outer worker/facility rows with one coalesced Facility-Abbrev.
shared #"Employees-MergedAvail+Roster" = let
    Source = #"Availabilities-EmployeesLIST",
    RosterSource = RosteredEmployeesTABLE,
    #"Joined Roster Workers" = Table.NestedJoin(Source, {"Payroll Code", "Facility-AbbrevZ"},
        RosterSource, {"Employee Code", "Facility-Abbrev"}, "RosteredEmployeeShifts (2)", JoinKind.FullOuter),
    #"Sorted Availability Payroll Codes" = Table.Sort(#"Joined Roster Workers", {{"Payroll Code", Order.Ascending}}),
    #"Expanded Roster Identity" = Table.ExpandTableColumn(#"Sorted Availability Payroll Codes", "RosteredEmployeeShifts (2)",
        {"Department", "Employee Roster Name", "Employment Type", "Role", "Employee Code", "Facility-Abbrev"},
        {"Department.1", "Employee Roster Name", "Employment Type", "Role", "Employee Code", "Allocation Facility-Abbrev"}),
    // Full-outer rows hold their facility in different columns depending on which membership sources are present.
    #"Added Unified Facility" = Table.AddColumn(#"Expanded Roster Identity", "Facility-Abbrev", each
        [#"Allocation Facility-Abbrev"] ?? [#"Facility-AbbrevZ"], type nullable text),
    #"Sorted Membership Rows" = Table.Sort(#"Added Unified Facility", {{"Payroll Code", Order.Ascending}})
in
    #"Sorted Membership Rows";

// Query: AvailabilitiesSTATS
// Purpose: Count reconciled availability/roster worker rows by unified facility and employment type.
// Inputs: Employees-MergedAvail+Roster.
// Output: Facility rows with one count column for each observed employment type.
shared AvailabilitiesSTATS = let
    Source = #"Employees-MergedAvail+Roster",
    #"Counted Facility Employment Types" = Table.Group(Source, {"Facility-Abbrev", "Employment Type"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Named Missing Employment Types" = Table.ReplaceValue(#"Counted Facility Employment Types", null, "Null", Replacer.ReplaceValue, {"Employment Type"}),
    EmploymentTypeValues = #"Named Missing Employment Types"[Employment Type],
    DistinctEmploymentTypes = List.Distinct(EmploymentTypeValues),
    #"Pivoted Employment Counts" = Table.Pivot(#"Named Missing Employment Types", DistinctEmploymentTypes, "Employment Type", "Count", List.Sum),
    PreferredColumns = {"Facility-Abbrev", "Casual", "Part Time", "Full Time", "Agency", "Null"},
    AvailableColumns = Table.ColumnNames(#"Pivoted Employment Counts"),
    ExistingPreferredColumns = List.Select(PreferredColumns, each List.Contains(AvailableColumns, _)),
    AdditionalColumns = List.Difference(AvailableColumns, ExistingPreferredColumns),
    SortedAdditionalColumns = List.Sort(AdditionalColumns),
    OutputColumnOrder = ExistingPreferredColumns & SortedAdditionalColumns,
    #"Ordered Summary Columns" = Table.ReorderColumns(#"Pivoted Employment Counts", OutputColumnOrder)
in
    #"Ordered Summary Columns";

// Query: Availabilities-Employees-AVAIL
// Purpose: Show available worker records with stable payroll and site sorting.
// Inputs: Availabilities-Employees, which already applies the analysis-site scope.
// Output: Available rows with Date To represented as datetime.
shared #"Availabilities-Employees-AVAIL" = let
    Source = #"Availabilities-Employees",
    #"Sorted Availability Flags" = Table.Sort(Source, {{"Available", Order.Ascending}}),
    #"Kept Available Records" = Table.SelectRows(#"Sorted Availability Flags", each [Available] = true),
    #"Sorted Payroll Codes" = Table.Sort(#"Kept Available Records", {{"Payroll Code", Order.Ascending}}),
    #"Typed Availability End" = Table.TransformColumnTypes(#"Sorted Payroll Codes", {{"Date To", type datetime}}),
    #"Sorted Site Names" = Table.Sort(#"Typed Availability End", {{"Site Name", Order.Ascending}})
in
    #"Sorted Site Names";

// Query: Availabilities-EmployeesDates
// Purpose: Show availability dates and roster role details for the existing selected employee.
// Inputs: Availabilities-Employees-AVAIL and RosteredEmployeesTABLE.
// Output: Availability rows with Date, numeric Subtraction, Employment Type and Role.
// Notes: Preserve the existing employee 13290 selection and duration-in-days Subtraction interface.
shared #"Availabilities-EmployeesDates" = let
    Source = #"Availabilities-Employees-AVAIL",
    #"Added Availability Date" = Table.AddColumn(Source, "Date", each DateTime.Date([Date From]), type date),
    #"Added Availability Duration" = Table.AddColumn(#"Added Availability Date", "Subtraction", each [Date To] - [Date From], type duration),
    #"Converted Duration To Days" = Table.TransformColumnTypes(#"Added Availability Duration", {{"Subtraction", type number}}),
    #"Kept Selected Employee" = Table.SelectRows(#"Converted Duration To Days", each [Payroll Code] = "13290"),
    RosterSource = RosteredEmployeesTABLE,
    #"Joined Roster Details" = Table.NestedJoin(#"Kept Selected Employee", {"Payroll Code"}, RosterSource, {"Employee Code"}, "RosteredEmployeesTABLE", JoinKind.LeftOuter),
    #"Expanded Roster Details" = Table.ExpandTableColumn(#"Joined Roster Details", "RosteredEmployeesTABLE", {"Employment Type", "Role"}, {"Employment Type", "Role"})
in
    #"Expanded Roster Details";

// Query: Availabilities - Not in Roster
// Purpose: Show records workers whose employee ID does not occur anywhere in the allocation roster.
// Notes: This is deliberately a worker-level diagnostic; facility is not part of the comparison grain.
shared #"Availabilities - Not in Roster" = let
    Source = #"Availabilities-Employees",
    #"Selected Records Workers" = Table.SelectColumns(Source,{"Payroll Code", "Available"}),
    #"Filtered Valid Records IDs" = Table.SelectRows(#"Selected Records Workers", each [Payroll Code] <> null),
    #"Distinct Records Workers" = Table.Distinct(#"Filtered Valid Records IDs"),
    RosterSource = RosteredEmployeesTABLE,
    #"Selected Allocation Workers" = Table.SelectColumns(RosterSource,{"Employee Code"}),
    #"Filtered Valid Allocation IDs" = Table.SelectRows(#"Selected Allocation Workers", each [Employee Code] <> null),
    #"Distinct Allocation Workers" = Table.Distinct(#"Filtered Valid Allocation IDs"),
    #"Kept Workers Missing From Allocation" = Table.NestedJoin(#"Distinct Records Workers", {"Payroll Code"}, #"Distinct Allocation Workers", {"Employee Code"}, "Allocation Match", JoinKind.LeftAnti),
    #"Removed Join Column" = Table.RemoveColumns(#"Kept Workers Missing From Allocation",{"Allocation Match"}),
    #"Buffered Availability Diagnostic" = Table.Buffer(#"Removed Join Column")
in
    #"Buffered Availability Diagnostic";