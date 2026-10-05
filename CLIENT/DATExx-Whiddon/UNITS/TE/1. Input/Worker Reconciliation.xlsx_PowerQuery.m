// Power Query from: Worker Reconciliation.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\1. Input\Worker Reconciliation.xlsx
// Extracted: 2026-10-05T06:26:31.826Z

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
    NormalizePath = (value as nullable text) as nullable text =>
        let
            SlashNormalized = if value = null then null else Text.Replace(Text.Trim(value), "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    WorkbookPathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
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
    CentriSyncPaths_Typed = Table.TransformColumnTypes(
        Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}),
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
    RelativePath = Text.TrimStart(Text.Range(FilePath, BestMatch[MatchRootLength]), "\"),
    LocalFullPath = if RelativePath = "" then BestMatch[SyncedFolderRootPath]
        else BestMatch[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase),
    InputIndex = List.PositionOf(Segments, "1. Input", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidUnitFolder = UnitsIndex >= 0 and InputIndex = UnitsIndex + 2 and InputIndex = List.Count(Segments) - 1,
    UnitRoot = if ValidUnitFolder then Text.BeforeDelimiter(WorkbookFolder, "\", {0, RelativePosition.FromEnd})
        else error "Worker Reconciliation.xlsx must be under UNITS/<unit>/1. Input.",
    PortfolioRoot = if ResidentialCareIndex > 0 then Text.Combine(List.FirstN(Segments, ResidentialCareIndex), "\")
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
        })
in
    Table.Buffer(Output);

// Query: Unit1Path
// Purpose: Expose the resolved unit folder for local ResidentialCare imports.
shared Unit1Path = PathTABLE{[Variable Name = "Unit Root"]}[Value];

// Query: PortfolioPath
// Purpose: Preserve the resolved sync portfolio folder as an existing workbook interface.
shared PortfolioPath = PathTABLE{[Variable Name = "Portfolio Root"]}[Value];

// Query: IMPORT Combined Roster
// Purpose: Read the local published-roster Combined sheet unchanged for roster-worker preparation.
shared #"IMPORT Combined Roster" = let
    SourcePath = Unit1Path & "\1. Input\Allocation\Published Roster.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "Combined" and [Kind] = "Sheet"),
    CombinedSheet = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Worker reconciliation roster import", "Expected exactly one Combined sheet.", [Matches = Table.RowCount(Matches)]),
    #"Promoted Headers" = Table.PromoteHeaders(CombinedSheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", type text}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", type text}})
in
    #"Changed Type";

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

// Query: Roster Prepare
// Purpose: Normalize the published roster and derive the base role used for reconciliation.
shared #"Roster Prepare" = let
    Source = #"IMPORT Combined Roster",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", type text}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", type text}}),
    FACILITIES = Table.SelectRows(#"Changed Type", each ([Location] <> "NR: Narrabri- Robert Young")),
    // Only a terminal AM, PM or NS token is a shift suffix. Non-shift titles must not lose their final three characters.
    #"Added Shift Suffix Flag" = Table.AddColumn(FACILITIES, "Shift", each
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
    #"Filtered BLANK AGENCYOUT" = Table.SelectRows(#"Expanded Assistant Label", each ([Employment Type] <> " " and [Employment Type] <> "Agency "))
in
    #"Filtered BLANK AGENCYOUT";

// Query: Roster Blank AGENCY
// Purpose: Review blank or agency allocation workers without failing on incomplete role or location text.
shared #"Roster Blank AGENCY" = let
    Source = #"IMPORT Combined Roster",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Employment Type] = " " or [Employment Type] = "Agency ")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Employment Type", Order.Descending}}),
    #"Removed Duplicates" = Table.Distinct(#"Sorted Rows", { "Employee Code", "Location"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Duplicates",{"Location", "Employment Type", "Role", "Employee Code"}),
    #"Extracted First Characters" = Table.TransformColumns(#"Removed Other Columns", {{"Role", each if _ = null then null else Text.Start(_, 15), type nullable text}}),
    #"Extracted First Characters1" = Table.TransformColumns(#"Extracted First Characters", {{"Location", each if _ = null then null else Text.Start(_, 2), type nullable text}}),
    #"Removed Duplicates1" = Table.Distinct(#"Extracted First Characters1"),
    #"Sorted Rows1" = Table.Sort(#"Removed Duplicates1",{{"Employee Code", Order.Descending}})
in
    #"Sorted Rows1";

// Query: RosterStartDate
// Purpose: Expose the first published-roster date used by employment eligibility and termination checks.
shared RosterStartDate = let
    RosterDates = List.RemoveNulls(#"Roster Prepare"[Date]),
    Output = if List.IsEmpty(RosterDates) then error "Published Roster contains no roster dates." else List.Min(RosterDates)
in
    Output;

shared RosteredEmployeeShiftsCOUNT = let
    Source = #"Roster Prepare",
    #"Grouped Rows" = Table.Group(Source, {"Location", "Employment Type"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", List.Distinct(#"Grouped Rows"[#"Employment Type"]), "Employment Type", "Count", List.Sum)
in
    #"Pivoted Column";

shared RosteredEmployeeOnlyShifts = let
    Source = #"Roster Prepare",
    FilterAGENCY = Table.SelectRows(Source, each ([Employment Type] <> "Agency ")),
    FilterBLANK.EMPCODE = Table.SelectRows(FilterAGENCY, each ([Employee Code] <> " "))
in
    FilterBLANK.EMPCODE;

// Query: Availabilities Prepare
// Purpose: Normalize availability/leave records and identify their facility and availability meaning.
shared #"Availabilities Prepare" = let
    Source = Table.PromoteHeaders(#"IMPORT Combined Availabilities", [PromoteAllScalars=true]),
    #"Changed Source Types" = Table.TransformColumnTypes(Source,{{"Site Name", type text}, {"Employee Name", type text}, {"Payroll Code", type text}, {"Department", type text}, {"Date From", type datetime}, {"Date To", type datetime}, {"Availability or Leave Reason", type text}}),
    #"Normalized Worker Fields" = Table.TransformColumns(#"Changed Source Types", {
        {"Employee Name", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Payroll Code", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text}
    }),
    Custom1 = Table.Buffer(#"Normalized Worker Fields"),
    #"FACILITY-ABBREV" = Table.AddColumn(Custom1, "Facility-Abbrev", each
        let Site = if [Site Name] = null then null else Text.Trim([Site Name])
        in if Site = null or Site = "" then null else Text.Upper(Text.Start(Site, 2)), type nullable text),
    #"Filtered FACILITY" = Table.SelectRows(#"FACILITY-ABBREV", each [#"Facility-Abbrev"] = null or [#"Facility-Abbrev"] <> "NR"),
    // Keep the records name so Records-only workers can still publish a complete identity.
    #"Removed Columns" = Table.RemoveColumns(#"Filtered FACILITY",{"Department"}),
    #"Added Availability Flag" = Table.AddColumn(#"Removed Columns", "Available", each
        let Reason = if [Availability or Leave Reason] = null then "" else Text.Upper(Text.From([Availability or Leave Reason]))
        in if Text.Contains(Reason, "UNAVAIL") then false else Text.Contains(Reason, "AVAIL"), type logical),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Availability Flag",{{"Payroll Code", type text}})
in
    #"Changed Type";

// Query: Availabilities - Not in Roster
// Purpose: Show records workers whose employee ID does not occur anywhere in the allocation roster.
// Notes: This is deliberately a worker-level diagnostic; facility is not part of the comparison grain.
shared #"Availabilities - Not in Roster" = let
    #"Selected Records Workers" = Table.SelectColumns(#"Availabilities-Employees",{"Payroll Code", "Available"}),
    #"Filtered Valid Records IDs" = Table.SelectRows(#"Selected Records Workers", each [Payroll Code] <> null),
    #"Distinct Records Workers" = Table.Distinct(#"Filtered Valid Records IDs"),
    #"Selected Allocation Workers" = Table.SelectColumns(RosteredEmployeesTABLE,{"Employee Code"}),
    #"Filtered Valid Allocation IDs" = Table.SelectRows(#"Selected Allocation Workers", each [Employee Code] <> null),
    #"Distinct Allocation Workers" = Table.Distinct(#"Filtered Valid Allocation IDs"),
    #"Kept Workers Missing From Allocation" = Table.NestedJoin(#"Distinct Records Workers", {"Payroll Code"}, #"Distinct Allocation Workers", {"Employee Code"}, "Allocation Match", JoinKind.LeftAnti),
    #"Removed Join Column" = Table.RemoveColumns(#"Kept Workers Missing From Allocation",{"Allocation Match"})
in
    Table.Buffer(#"Removed Join Column");

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
// Purpose: Reduce records membership to one worker/facility row while preserving a usable records-source name.
// Output: One row per Payroll Code and Facility-AbbrevZ.
shared #"Availabilities-EmployeesLIST" = let
    Source = #"Availabilities-Employees",
    #"Selected Worker Fields" = Table.SelectColumns(Source,{"Payroll Code", "Facility-Abbrev", "Employee Name"}),
    #"Renamed Facility" = Table.RenameColumns(#"Selected Worker Fields",{{"Facility-Abbrev", "Facility-AbbrevZ"}}),
    #"Grouped Worker Facility" = Table.Group(#"Renamed Facility", {"Payroll Code", "Facility-AbbrevZ"}, {
        {"Records Employee Name", each
            let
                Names = List.Sort(
                    List.Distinct(List.RemoveNulls([Employee Name]), Comparer.OrdinalIgnoreCase),
                    Comparer.OrdinalIgnoreCase)
            in
                if List.IsEmpty(Names) then null else Names{0}, type nullable text}
    }),
    // Preserve source presence independently of whether the worker identifier is complete.
    #"Added Records Source Marker" = Table.AddColumn(#"Grouped Worker Facility", "In Records Source", each true, type logical)
in
    Table.Buffer(#"Added Records Source Marker");

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
    #"Removed Location" = Table.RemoveColumns(#"Replaced ASSISTANT IN NURSING",{"Location"}),
    #"Removed Exact Duplicates" = Table.Distinct(#"Removed Location"),
    // Blank roster lines are vacancies or unassigned shifts, not workers.
    #"Filtered Identified Workers" = Table.SelectRows(#"Removed Exact Duplicates", each [Employee Code] <> null),
    // Preserve source presence independently of whether the worker identifier is complete.
    #"Added Allocation Source Marker" = Table.AddColumn(#"Filtered Identified Workers", "In Allocation Source", each true, type logical)
in
    Table.Buffer(#"Added Allocation Source Marker");

// Query: IMPORT Current_Workers
// Purpose: Read the local unfiltered worker report unchanged for employment, termination, role and contract enrichment.
// Inputs: Table1 in 1. Input/Worker's Report.xlsx.
shared #"IMPORT Current_Workers" = let
    SourcePath = Unit1Path & "\1. Input\Worker's Report.xlsx",
    Navigation = Excel.Workbook(Binary.Buffer(File.Contents(SourcePath)), null, true),
    Matches = Table.SelectRows(Navigation, each [Item] = "Table1" and [Kind] = "Table"),
    CurrentWorkers = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Current workers import", "Expected exactly one Table1 table in Worker's Report.xlsx.", [Matches = Table.RowCount(Matches)])
in
    CurrentWorkers;

// Query: CurrentWorkers_Prepare
// Purpose: Validate and normalize current-worker employment, termination and contract fields used by reconciliation.
shared CurrentWorkers_Prepare = let
    RequiredColumns = {"Employee Code", "Location Description", "Position Description", "Employee Status Status Type", "Start Date", "Termination Date", "Termination Reason Description", "Contracted FN Hours"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(#"IMPORT Current_Workers")),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(#"IMPORT Current_Workers", RequiredColumns)
        else error Error.Record("Current workers preparation", "Worker's Report.xlsx/Table1 is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected,{{"Employee Code", type text}, {"Location Description", type text}, {"Position Description", type text}, {"Employee Status Status Type", type text}, {"Start Date", type date}, {"Termination Date", type nullable date}, {"Termination Reason Description", type text}, {"Contracted FN Hours", type nullable number}}),
    #"Added Facility" = Table.AddColumn(Typed, "Facility-Abbrev", each
        let Location = if [Location Description] = null then null else Text.Trim(Text.From([Location Description]))
        in if Location = null or Location = "" then null else Text.Upper(Text.Start(Location, 2)), type nullable text),
    Normalized = Table.TransformColumns(#"Added Facility", {
        {"Employee Code", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Facility-Abbrev", each let Value = if _ = null then null else Text.Upper(Text.Trim(Text.From(_))) in if Value = "" then null else Value, type nullable text},
        {"Position Description", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text},
        {"Employee Status Status Type", each let Value = if _ = null then null else Text.Trim(Text.From(_)) in if Value = "" then null else Value, type nullable text}
    })
in
    Table.Buffer(Normalized);

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
            Table.First(Table.Sort(_, {
                {"Roster Eligibility Rank", Order.Ascending},
                {"Open Employment Rank", Order.Ascending},
                {"Termination Date", Order.Descending},
                {"Start Date", Order.Descending},
                {"Facility-Abbrev", Order.Ascending},
                {"Position Description", Order.Ascending}
            })), type record},
        {"Current Worker Row Count", each Table.RowCount(_), Int64.Type},
        // Repeated source rows with the same active employment facts are not materially ambiguous.
        {"Current Worker Eligible Row Count", each
            let
                EligibleRows = Table.SelectRows(_, each [Termination Date] = null or [Termination Date] >= RosterStartDate),
                MaterialFields = Table.SelectColumns(EligibleRows,
                    {"Facility-Abbrev", "Position Description", "Employee Status Status Type", "Termination Date", "Contracted FN Hours"}),
                DistinctEligibleRows = Table.Distinct(MaterialFields)
            in
                Table.RowCount(DistinctEligibleRows), Int64.Type}
    }),
    #"Expanded Selected Current Worker" = Table.ExpandRecordColumn(#"Grouped Worker", "Selected Current Worker",
        {"Facility-Abbrev", "Location Description", "Position Description", "Employee Status Status Type", "Start Date", "Termination Date", "Termination Reason Description", "Contracted FN Hours"},
        {"Facility-Abbrev", "Location Description", "Position Description", "Employee Status Status Type", "Start Date", "Termination Date", "Termination Reason Description", "Contracted FN Hours"})
in
    Table.Buffer(#"Expanded Selected Current Worker");

// Query: Employees-ALL
// Purpose: Reconcile allocation and records membership, then attach current-worker status and termination details.
// Output: One or more auditable candidate rows per employee/facility/role before active-worker filtering.
shared #"Employees-ALL" = let
    RosterStart = RosterStartDate,
    Source = Table.NestedJoin(#"Availabilities-EmployeesLIST", {"Payroll Code", "Facility-AbbrevZ"}, RosteredEmployeesTABLE, {"Employee Code", "Facility-Abbrev"}, "RosteredEmployeeShifts (2)", JoinKind.FullOuter),
    #"Expanded RosteredEmployeeShifts (2)" = Table.ExpandTableColumn(Source, "RosteredEmployeeShifts (2)", {"Department", "Employee Roster Name", "Employment Type", "Employee Code", "Role", "Facility-Abbrev", "In Allocation Source"}, {"Department", "Employee Roster Name", "Employment Type", "Employee Code", "RoleX", "Facility-AbbrevX", "In Allocation Source"}),
    #"Added In Records" = Table.AddColumn(#"Expanded RosteredEmployeeShifts (2)", "In Records", each [#"In Records Source"] = true, type logical),
    #"Added In Allocation" = Table.AddColumn(#"Added In Records", "In Allocation", each [#"In Allocation Source"] = true, type logical),
    #"Added Worker Membership" = Table.AddColumn(#"Added In Allocation", "Worker Membership", each
        if [In Allocation] and [In Records] then "Both"
        else if [In Allocation] then "Allocation only"
        else if [In Records] then "Records only"
        else error "A reconciled worker must appear in Allocation, Records, or both.", type text),
    #"Added Assigned Employee" = Table.AddColumn(#"Added Worker Membership", "AssignedAvailID", each [Payroll Code] ?? [Employee Code], type nullable text),
    #"Added Join Facility" = Table.AddColumn(#"Added Assigned Employee", "Join Facility-Abbrev", each [#"Facility-AbbrevX"] ?? [#"Facility-AbbrevZ"], type nullable text),
    // Employment status is employee-level; the published facility remains the Allocation/Records facility above.
    #"Merged Current Workers" = Table.NestedJoin(#"Added Join Facility", {"AssignedAvailID"}, CurrentWorkers_Reconciled, {"Employee Code"}, "Current Worker", JoinKind.LeftOuter),
    #"Expanded Current Worker" = Table.ExpandTableColumn(#"Merged Current Workers", "Current Worker",
        {"Employee Code", "Facility-Abbrev", "Position Description", "Employee Status Status Type", "Contracted FN Hours", "Start Date", "Termination Date", "Termination Reason Description", "Current Worker Row Count", "Current Worker Eligible Row Count"},
        {"Employee Code.1", "Facility-AbbrevY", "Position Description", "Employee Status Status Type", "Contracted FN Hours", "Start Date", "Termination Date", "Termination Reason Description", "Current Worker Row Count", "Current Worker Eligible Row Count"}),
    #"Added Current Worker Match Missing" = Table.AddColumn(#"Expanded Current Worker", "Current Worker Match Missing", each [Employee Code.1] = null, type logical),
    #"Added Current Worker Record Ambiguous" = Table.AddColumn(#"Added Current Worker Match Missing", "Current Worker Record Ambiguous", each
        [Current Worker Eligible Row Count] <> null and [Current Worker Eligible Row Count] > 1, type logical),
    #"Added Terminated Before Roster" = Table.AddColumn(#"Added Current Worker Record Ambiguous", "Terminated Before Roster", each
        not [Current Worker Match Missing] and [Termination Date] <> null and [Termination Date] < RosterStart, type logical)
in
    #"Added Terminated Before Roster";

// Query: EmployeesTABLE_Prepare
// Purpose: Prepare active reconciled workers with explicit Allocation/Records membership and one preferred role.
// Output: One preferred worker row per employee/facility.
shared EmployeesTABLE_Prepare = let
    Source = #"Employees-ALL",
    #"Filtered Terminated Workers" = Table.SelectRows(Source, each not [Terminated Before Roster]),
    #"Renamed Allocation Name" = Table.RenameColumns(#"Filtered Terminated Workers",{{"Employee Roster Name", "Allocation Employee Name"}}),
    #"Added Employee ID" = Table.AddColumn(#"Renamed Allocation Name", "EmployeeID", each [AssignedAvailID] ?? [Employee Code.1], type nullable text),
    #"Added Facility" = Table.AddColumn(#"Added Employee ID", "Facility-Abbrev", each [#"Join Facility-Abbrev"] ?? [#"Facility-AbbrevY"], type nullable text),
    #"Added Employee Name" = Table.AddColumn(#"Added Facility", "Employee Roster Name", each [Allocation Employee Name] ?? [Records Employee Name], type nullable text),
    #"Added Employment Type" = Table.AddColumn(#"Added Employee Name", "EmploymentType", each [Employment Type] ?? [Employee Status Status Type], type nullable text),
    #"Added Preferred Role Candidate" = Table.AddColumn(#"Added Employment Type", "Role", each [RoleX] ?? [Position Description], type nullable text),
    #"Selected Worker Fields" = Table.SelectColumns(#"Added Preferred Role Candidate",{
        "Allocation Employee Name", "Records Employee Name", "Employee Roster Name", "EmployeeID", "Facility-Abbrev", "EmploymentType", "Role",
        "In Allocation", "In Records", "Worker Membership", "Current Worker Match Missing", "Current Worker Record Ambiguous",
        "Current Worker Row Count", "Current Worker Eligible Row Count", "Position Description", "Contracted FN Hours", "Termination Date", "Terminated Before Roster"}),
    #"Sorted Rows" = Table.Sort(#"Selected Worker Fields",{{"EmployeeID", Order.Ascending}}),
    #"Replaced -" = Table.ReplaceValue(#"Sorted Rows"," ","-",Replacer.ReplaceText,{"EmploymentType"}),
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
    #"Added Role Alignment Rank" = Table.AddColumn(#"Replaced -", "Role Alignment Rank", each
        RoleAlignmentRank([Role], [Position Description]), Int64.Type),
    PreferredWorkerColumns = List.RemoveItems(Table.ColumnNames(#"Added Role Alignment Rank"), {"EmployeeID", "Facility-Abbrev"}),
    #"Grouped Preferred Workers" = Table.Group(#"Added Role Alignment Rank", {"EmployeeID", "Facility-Abbrev"}, {
        {"Preferred Worker", each Table.First(Table.Sort(_, {
            {"Role Alignment Rank", Order.Ascending},
            {"Role", Order.Ascending},
            {"Employee Roster Name", Order.Ascending},
            {"EmploymentType", Order.Ascending}
        })), type record}
    }),
    #"Expanded Preferred Worker" = Table.ExpandRecordColumn(#"Grouped Preferred Workers", "Preferred Worker", PreferredWorkerColumns, PreferredWorkerColumns),
    #"Removed Preference Fields" = Table.RemoveColumns(#"Expanded Preferred Worker",{"Role Alignment Rank", "Position Description"})
in
    Table.Buffer(#"Removed Preference Fields");

// Query: TerminatedWorkers_DIAGNOSTICS
// Purpose: Show reconciled candidates excluded because their termination date precedes the roster start.
shared TerminatedWorkers_DIAGNOSTICS = let
    Source = Table.SelectRows(#"Employees-ALL", each [Terminated Before Roster]),
    WithEmployeeID = Table.AddColumn(Source, "EmployeeID", each [AssignedAvailID] ?? [Employee Code.1], type nullable text),
    WithFacility = Table.AddColumn(WithEmployeeID, "Facility-Abbrev", each [#"Join Facility-Abbrev"] ?? [#"Facility-AbbrevY"], type nullable text),
    RenamedAllocationName = Table.RenameColumns(WithFacility,{{"Employee Roster Name", "Allocation Employee Name"}}),
    WithEmployeeName = Table.AddColumn(RenamedAllocationName, "Employee Roster Name", each [Allocation Employee Name] ?? [Records Employee Name], type nullable text),
    Selected = Table.SelectColumns(WithEmployeeName,
        {"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Termination Date", "Termination Reason Description", "In Allocation", "In Records", "Worker Membership"}),
    Output = Table.Distinct(Selected)
in
    Table.Buffer(Output);

// Query: DuplicateWorkers_DIAGNOSTICS
// Purpose: Show every employee/facility key that still has more than one preferred candidate row.
shared DuplicateWorkers_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    DistinctText = (values as list) as text =>
        Text.Combine(List.Sort(List.Distinct(List.RemoveNulls(List.Transform(values, each Text.From(_))), Comparer.OrdinalIgnoreCase), Comparer.OrdinalIgnoreCase), " | "),
    Grouped = Table.Group(Source, {"EmployeeID", "Facility-Abbrev"}, {
        {"Candidate Count", each Table.RowCount(_), Int64.Type},
        {"Employee Roster Name", each DistinctText([Employee Roster Name]), type text},
        {"Role", each DistinctText([Role]), type text},
        {"Worker Membership", each DistinctText([Worker Membership]), type text}
    }),
    Output = Table.SelectRows(Grouped, each [Candidate Count] > 1)
in
    Table.Buffer(Output);

// Query: IncompleteWorkerIdentity_DIAGNOSTICS
// Purpose: Show retained candidates missing a required identifier, facility, resolved name or role.
shared IncompleteWorkerIdentity_DIAGNOSTICS = let
    Source = EmployeesTABLE_Prepare,
    IsBlank = (value as any) as logical => value = null or (Value.Is(value, type text) and Text.Trim(value) = ""),
    Filtered = Table.SelectRows(Source, each IsBlank([EmployeeID]) or IsBlank([#"Facility-Abbrev"]) or IsBlank([Employee Roster Name]) or IsBlank([Role])),
    WithMissingFields = Table.AddColumn(Filtered, "Missing Fields", each Text.Combine(List.RemoveNulls({
        if IsBlank([EmployeeID]) then "EmployeeID" else null,
        if IsBlank([#"Facility-Abbrev"]) then "Facility-Abbrev" else null,
        if IsBlank([Employee Roster Name]) then "Employee Roster Name" else null,
        if IsBlank([Role]) then "Role" else null
    }), ", "), type text),
    Output = Table.SelectColumns(WithMissingFields,{"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Missing Fields"})
in
    Table.Buffer(Output);

// Query: UnmatchedCurrentWorkers_DIAGNOSTICS
// Purpose: Show membership candidates that could not be matched to the worker report by employee ID.
// Notes: Allocation membership is sufficient for publication; these rows receive the downstream Settings contract fallback.
shared UnmatchedCurrentWorkers_DIAGNOSTICS = let
    Source = Table.SelectRows(EmployeesTABLE_Prepare, each [Current Worker Match Missing]),
    Output = Table.SelectColumns(Source,{"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "In Allocation", "In Records"})
in
    Table.Buffer(Output);

// Query: AmbiguousCurrentWorkers_DIAGNOSTICS
// Purpose: Show workers with more than one worker-report row eligible at roster start.
shared AmbiguousCurrentWorkers_DIAGNOSTICS = let
    Source = Table.SelectRows(EmployeesTABLE_Prepare, each [Current Worker Record Ambiguous]),
    Output = Table.SelectColumns(Source,{"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Current Worker Row Count", "Current Worker Eligible Row Count"})
in
    Table.Buffer(Output);

// Query: InvalidWorkerMembership_DIAGNOSTICS
// Purpose: Show candidates whose Allocation/Records flags do not agree with the published membership label.
shared InvalidWorkerMembership_DIAGNOSTICS = let
    ValidMembershipValues = {"Allocation only", "Records only", "Both"},
    Source = Table.SelectRows(EmployeesTABLE_Prepare, each
        (not [In Allocation] and not [In Records])
            or not List.Contains(ValidMembershipValues, [Worker Membership])),
    Output = Table.SelectColumns(Source,{"EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "In Allocation", "In Records"})
in
    Table.Buffer(Output);

// Query: LeakedTerminatedWorkers_DIAGNOSTICS
// Purpose: Independently anti-check that no excluded terminated worker/facility key remains in the publishable population.
shared LeakedTerminatedWorkers_DIAGNOSTICS = let
    TerminatedKeys = Table.Distinct(Table.SelectColumns(TerminatedWorkers_DIAGNOSTICS,{"EmployeeID", "Facility-Abbrev", "Termination Date"})),
    Joined = Table.NestedJoin(EmployeesTABLE_Prepare, {"EmployeeID", "Facility-Abbrev"}, TerminatedKeys, {"EmployeeID", "Facility-Abbrev"}, "Terminated Match", JoinKind.Inner),
    Output = Table.ExpandTableColumn(Joined, "Terminated Match", {"Termination Date"}, {"Excluded Termination Date"})
in
    Table.Buffer(Output);

// Query: WorkerReconciliation_ISSUES
// Purpose: Consolidate every blocking reconciliation issue into one row-level review table.
shared WorkerReconciliation_ISSUES = let
    DuplicateBase = Table.AddColumn(DuplicateWorkers_DIAGNOSTICS, "Issue", each "Duplicate employee/facility", type text),
    DuplicateWithDetails = Table.AddColumn(DuplicateBase, "Details", each Text.From([Candidate Count]) & " candidate rows remain for this key.", type text),
    DuplicateIssues = Table.SelectColumns(DuplicateWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    IdentityBase = Table.AddColumn(IncompleteWorkerIdentity_DIAGNOSTICS, "Issue", each "Incomplete worker identity", type text),
    IdentityWithDetails = Table.AddColumn(IdentityBase, "Details", each "Missing: " & [Missing Fields], type text),
    IdentityIssues = Table.SelectColumns(IdentityWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    MembershipBase = Table.AddColumn(InvalidWorkerMembership_DIAGNOSTICS, "Issue", each "Invalid worker membership", type text),
    MembershipWithDetails = Table.AddColumn(MembershipBase, "Details", each "Allocation/Records flags do not agree with Worker Membership.", type text),
    MembershipIssues = Table.SelectColumns(MembershipWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    // Only a Records-only row requires worker-report enrichment; Allocation supplies identity and role for unmatched roster workers.
    UnmatchedBlocking = Table.SelectRows(UnmatchedCurrentWorkers_DIAGNOSTICS, each not [In Allocation]),
    UnmatchedBase = Table.AddColumn(UnmatchedBlocking, "Issue", each "Current worker match missing", type text),
    UnmatchedWithDetails = Table.AddColumn(UnmatchedBase, "Details", each "No Worker's Report row matched this Records-only employee ID.", type text),
    UnmatchedIssues = Table.SelectColumns(UnmatchedWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    AmbiguousBase = Table.AddColumn(AmbiguousCurrentWorkers_DIAGNOSTICS, "Issue", each "Current worker record ambiguous", type text),
    AmbiguousWithDetails = Table.AddColumn(AmbiguousBase, "Details", each Text.From([Current Worker Eligible Row Count]) & " Worker's Report rows are eligible at roster start for this employee.", type text),
    AmbiguousIssues = Table.SelectColumns(AmbiguousWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    TerminatedBase = Table.AddColumn(LeakedTerminatedWorkers_DIAGNOSTICS, "Issue", each "Terminated worker leaked", type text),
    TerminatedWithDetails = Table.AddColumn(TerminatedBase, "Details", each "Excluded termination date: " & Date.ToText([Excluded Termination Date], "yyyy-MM-dd"), type text),
    TerminatedIssues = Table.SelectColumns(TerminatedWithDetails,{"Issue", "EmployeeID", "Facility-Abbrev", "Employee Roster Name", "Role", "Worker Membership", "Details"}),

    Output = Table.Combine({DuplicateIssues, IdentityIssues, MembershipIssues, UnmatchedIssues, AmbiguousIssues, TerminatedIssues})
in
    Table.Buffer(Output);

// Query: WorkerReconciliation_CHECK
// Purpose: Summarize and block publication for duplicate grain, incomplete identity, invalid membership or unsafe employment joins.
shared WorkerReconciliation_CHECK = let
    CheckResult = (name as text, failures as number, details as text) as record =>
        [Check = name, Status = if failures = 0 then "Pass" else "Fail", Failures = failures, Details = details],
    Checks = {
        CheckResult("Unique employee and facility workers", Table.RowCount(DuplicateWorkers_DIAGNOSTICS), "Employees_TABLE must publish one preferred row per employee and facility."),
        CheckResult("Complete worker identity", Table.RowCount(IncompleteWorkerIdentity_DIAGNOSTICS), "EmployeeID, facility, resolved employee name and role are required."),
        CheckResult("Valid worker membership", Table.RowCount(InvalidWorkerMembership_DIAGNOSTICS), "Each worker must be Allocation only, Records only, or Both."),
        CheckResult("Current worker match complete", Table.RowCount(Table.SelectRows(UnmatchedCurrentWorkers_DIAGNOSTICS, each not [In Allocation])), "Every Records-only membership candidate must match Worker's Report by employee ID; Allocation-only workers may use the Settings contract fallback."),
        CheckResult("Current worker record unambiguous", Table.RowCount(AmbiguousCurrentWorkers_DIAGNOSTICS), "Each employee may have at most one materially distinct Worker's Report row eligible at roster start."),
        CheckResult("No terminated workers published", Table.RowCount(LeakedTerminatedWorkers_DIAGNOSTICS), "Workers terminated before roster start must remain diagnostic-only.")
    }
in
    Table.Buffer(Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = number, Details = text]));

// Query: Employees-TABLE
// Purpose: Publish the validated authoritative worker population consumed by Capacity-ShiftAvailability.
shared #"Employees-TABLE" = let
    Failures = Table.SelectRows(WorkerReconciliation_CHECK, each [Status] = "Fail"),
    Checked = if Table.IsEmpty(Failures) then EmployeesTABLE_Prepare
        else error Error.Record("Worker reconciliation validation", "Required worker checks failed.", Failures),
    Output = Table.RemoveColumns(Checked, {
        "Allocation Employee Name", "Records Employee Name", "Current Worker Match Missing", "Current Worker Record Ambiguous",
        "Current Worker Row Count", "Current Worker Eligible Row Count", "Termination Date", "Terminated Before Roster"})
in
    Table.Buffer(Output);

// Query: Employees-MergedAvail+Roster
// Purpose: Reconcile records and allocation rows for the legacy availability membership summary.
// Output: Full-outer worker/facility rows with one coalesced Facility-Abbrev.
shared #"Employees-MergedAvail+Roster" = let
    Source = Table.NestedJoin(#"Availabilities-EmployeesLIST", {"Payroll Code", "Facility-AbbrevZ"}, RosteredEmployeesTABLE, {"Employee Code", "Facility-Abbrev"}, "RosteredEmployeeShifts (2)", JoinKind.FullOuter),
    #"Sorted Rows" = Table.Sort(Source,{{"Payroll Code", Order.Ascending}}),
    #"Expanded RosteredEmployeeShifts (2)" = Table.ExpandTableColumn(#"Sorted Rows", "RosteredEmployeeShifts (2)", {"Department", "Employee Roster Name", "Employment Type", "Role", "Employee Code", "Facility-Abbrev"}, {"Department.1", "Employee Roster Name", "Employment Type", "Role", "Employee Code", "Allocation Facility-Abbrev"}),
    // Records-only and allocation-only rows hold facility in different source columns after the full outer join.
    #"Added Unified Facility" = Table.AddColumn(#"Expanded RosteredEmployeeShifts (2)", "Facility-Abbrev", each [#"Allocation Facility-Abbrev"] ?? [#"Facility-AbbrevZ"], type nullable text),
    #"Sorted Rows1" = Table.Sort(#"Added Unified Facility",{{"Payroll Code", Order.Ascending}})
in
    #"Sorted Rows1";

// Query: AvailabilitiesSTATS
// Purpose: Count reconciled records/allocation worker rows by unified facility and normalized employment type.
shared AvailabilitiesSTATS = let
    Source = #"Employees-MergedAvail+Roster",
    #"Grouped Rows" = Table.Group(Source, {"Facility-Abbrev", "Employment Type"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped Rows",null,"Null",Replacer.ReplaceValue,{"Employment Type"}),
    #"Pivoted Column" = Table.Pivot(#"Replaced Value", List.Distinct(#"Replaced Value"[#"Employment Type"]), "Employment Type", "Count", List.Sum),
    PreferredColumns = {"Facility-Abbrev", "Casual", "Part Time", "Full Time", "Agency", "Null"},
    ExistingPreferredColumns = List.Select(PreferredColumns, each List.Contains(Table.ColumnNames(#"Pivoted Column"), _)),
    OtherColumns = List.Sort(List.Difference(Table.ColumnNames(#"Pivoted Column"), ExistingPreferredColumns)),
    #"Reordered Columns" = Table.ReorderColumns(#"Pivoted Column", ExistingPreferredColumns & OtherColumns)
in
    #"Reordered Columns";

shared #"Availabilities-Employees-AVAIL" = let
    Source = #"Availabilities-Employees",
    #"Sorted Rows2" = Table.Sort(Source,{{"Available", Order.Ascending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows2", each ([Available] = true)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Payroll Code", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"Date To", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type",{{"Site Name", Order.Ascending}})
in
    #"Sorted Rows1";

shared #"Availabilities-EmployeesDates" = let
    Source = #"Availabilities-Employees-AVAIL",
    #"Inserted Date" = Table.AddColumn(Source, "Date", each DateTime.Date([Date From]), type date),
    #"Inserted Time Subtraction" = Table.AddColumn(#"Inserted Date", "Subtraction", each [Date To] - [Date From], type duration),
    #"Changed Type1" = Table.TransformColumnTypes(#"Inserted Time Subtraction",{{"Subtraction", type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Payroll Code] = "13290")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Payroll Code"}, RosteredEmployeesTABLE, {"Employee Code"}, "RosteredEmployeesTABLE", JoinKind.LeftOuter),
    #"Expanded RosteredEmployeesTABLE" = Table.ExpandTableColumn(#"Merged Queries", "RosteredEmployeesTABLE", {"Employment Type", "Role"}, {"Employment Type", "Role"})
in
    #"Expanded RosteredEmployeesTABLE";