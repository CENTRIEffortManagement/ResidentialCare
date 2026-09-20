// Power Query from: Demand-MasterRoster Manual Read.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\Demand-MasterRoster Manual Read.xlsx
// Extracted: 2026-09-16T01:29:23.598Z

section Section1;

// Query: IMPORT CentriSyncPaths
// Purpose: Read the machine's shared path mapping for portable workbook imports.
shared #"IMPORT CentriSyncPaths" =
let
    Navigation = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    Mapping = Navigation{[Item="CentriSyncPaths", Kind="Table"]}[Data]
in
    Table.Buffer(Mapping);

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook's path through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl, either the standard one-row FilePath table or this workbook's single named cell.
// Output: Variable Name / Value rows for the existing UnitL1PathTABLE worksheet table.
// Notes: The worksheet output is not a path input; reading it here would reuse stale query results.
shared UnitL1PathTABLE =
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        // Excel exposes the existing single-cell FilePathUrl name as Column1.
        // Keep the canonical FilePath column when the standard input table is installed.
        PathInput = if Table.HasColumns(Source, {"FilePath"}) then Source
            else if Table.ColumnNames(Source) = {"Column1"} then
                Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must be a one-row FilePath table or a single named cell.",
        SelectedColumns = Table.SelectColumns(PathInput, {"FilePath"}),
        ChangedType = Table.TransformColumnTypes(SelectedColumns, {{"FilePath", type text}}),
        ReplacedValue = Table.TransformColumns(ChangedType, {{"FilePath", each if _ = null then null else Text.Replace(_, "/", "\"), type text}}),
        ValidatedTable = if Table.RowCount(ReplacedValue) = 1 then ReplacedValue else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedTable)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if RawFilePathValue = null or Text.Trim(RawFilePathValue) = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else Text.Trim(RawFilePathValue),
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet.
    // Remove only the workbook brackets and sheet suffix before resolving the folder.
    WorkbookPath = if Text.Contains(RawFilePath, "[") then
        Text.BeforeDelimiter(RawFilePath, "[") & Text.BetweenDelimiters(RawFilePath, "[", "]")
        else RawFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    ValidatedWorkbookPath = if Comparer.OrdinalIgnoreCase(InputFileName, "Demand-MasterRoster Manual Read.xlsx") = 0 then
        WorkbookPath
        else error "FilePathUrl identifies another workbook. In the named cell use =CELL(""filename"",A1), then save and recalculate Demand-MasterRoster Manual Read.xlsx.",
    CentriSyncPaths_Table = #"IMPORT CentriSyncPaths",
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(ValidatedWorkbookPath),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
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
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    BestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(LocalFullPath, "[", "]") otherwise Text.AfterDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"FilePathUrl", FilePath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

// Query: Unit1Path
// Purpose: Derive the unit root from the resolved workbook folder for relative imports.
shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    WorkbookFolder = #"Filtered Rows"{0}[Value],
    UnitFolder =
        if Text.EndsWith(WorkbookFolder, "\1. Input", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\1. Input"))
        else if Text.EndsWith(WorkbookFolder, "\2. Calculations", Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length("\2. Calculations"))
        else
            error "Root Path did not end in an expected Unit1 workbook folder: " & WorkbookFolder
in
    UnitFolder;

// Query: IMPORT Settings Data
// Purpose: Read the current unit's Settings workbook for the standard-FTE duration.
shared #"IMPORT Settings Data" =
    Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
        Unit1Path & "\2. Calculations\Settings Data.xlsx")), null, true));

// Query: ShiftDuration
// Purpose: Read and validate the standard-FTE duration in hours from Settings Data.
// Inputs: IMPORT Settings Data, named table ShiftDuration, column ShiftDuration.
// Output: One positive finite duration in hours; missing or invalid settings stop publication.
// Notes: This single extraction goes directly from IMPORT to the named result query.
shared ShiftDuration =
let
    Matches = Table.SelectRows(#"IMPORT Settings Data",
        each [Item] = "ShiftDuration" and [Kind] = "Table"),
    Data = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error "Settings Data must contain exactly one ShiftDuration table.",
    RequiredColumn = if Table.HasColumns(Data, {"ShiftDuration"}) then Data
        else error "The ShiftDuration table must contain the ShiftDuration column.",
    Hours = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn{0}[ShiftDuration]
        else error "ShiftDuration must contain exactly one data row.",
    ValidatedHours = if not Value.Is(Hours, type number) then
            error "ShiftDuration must be a numeric hours value."
        else if Number.IsNaN(Hours) or Hours <= 0 or Number.Abs(Hours) = #infinity then
            error "ShiftDuration must be a positive finite number of hours."
        else Hours
in
    ValidatedHours;


shared #"IMPORT Master" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\ResidentialCare\Whiddon\260728\UNITS\Unit1\1. Input\Master Roster.xlsx"), null, true),
    Combined_Sheet = Source{[Item="Combined",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(Combined_Sheet, [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Master Template", type text}, {"Template", type text}, {"Location", type text}, {"Department", type text}, {"Role", type text}, {"Area", type text}, {"Employee Code", Int64.Type}, {"Employee Name", type text}, {"Week No", Int64.Type}, {"Week Day", type text}, {"Start Time", type time}, {"End Time", type time}, {"Roster Hours", type number}, {"Cost", type number}, {"MinRosterHours", type number}, {"MaxRosterHours", Int64.Type}, {"Event", type text}, {"Break Length", Int64.Type}, {"Break Start Time", type time}, {"Paid Break Length", Int64.Type}, {"Paid Break Start Time", type time}, {"Shift Definition", type text}, {"Shift Net Length", type number}, {"Shift Type", type text}, {"Non Attended", type logical}})
in
    #"Changed Type1";

shared #"INPUT SHiftEnd" = let
    Source = Excel.CurrentWorkbook(){[Name="Table7"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shift", type text}, {"End", type time}})
in
    #"Changed Type";

shared #"ShftEnd Prepare" = let
    Source = #"INPUT SHiftEnd",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type)
in
    #"Added Index";

// Query: INPUT MinuteWorkers
// Purpose: Reads the authoritative roster-role to direct-care-role mapping without altering workbook table rows or types.
// Inputs: Current-workbook table MatchingRosterRoleswithANACCRoles.
// Output: Roster Roles, DC Category, DC Role, and Direct Care % exactly as supplied.
shared #"INPUT MinuteWorkers" =
let
    Source = Excel.CurrentWorkbook(){[Name = "MatchingRosterRoleswithANACCRoles"]}[Content]
in
    Source;

// Query: MW Roster Role Mapping Prepare
// Purpose: Normalises the explicit roster-role mapping and derives compatibility keys without inferring relationships.
// Inputs: INPUT MinuteWorkers.
// Output: One row per supplied mapping with original labels, normalised keys, assigned DC Role, DC Category and MinuteCategory.
// Notes: NA remains visible here for validation but is excluded before historical calculations.
shared #"MW Roster Role Mapping Prepare" =
let
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
        each if [DC Role] = null then null else [DC Role],
        type nullable text
    ),
    #"Added Minute Category" = Table.AddColumn(
        #"Added DC Role Key",
        "MinuteCategory",
        each if [DC Category] = "RN" then "RN" else if [DC Category] = "OTHER" then "OTHERS" else null,
        type nullable text
    )
in
    #"Added Minute Category";

// Query: MW MinuteWorkers Prepare
// Purpose: Produces one configuration row per retained DC Role for downstream role-level joins.
// Inputs: MW Roster Role Mapping Prepare.
// Output: Distinct DC Role/Role/RoleKey, DC Category, MinuteCategory and Direct Care % combinations.
// Notes: Multiple Roster Roles may intentionally map to one DC Role. Invalid conflicts are exposed by MW Input Check.
shared #"MW MinuteWorkers Prepare" =
let
    Source = Table.SelectRows(
        #"MW Roster Role Mapping Prepare",
        each List.Contains({"RN", "OTHER"}, [DC Category])
    ),
    #"Selected DC Role Attributes" = Table.SelectColumns(
        Source,
        {"RoleKey", "DC Category", "MinuteCategory", "Direct Care %"}
    ),
    // Downstream labels use the normalised DC role so case-only source
    // differences cannot split one analytical role.
    #"Added Normalised DC Role" = Table.AddColumn(
        #"Selected DC Role Attributes",
        "DC Role",
        each [RoleKey],
        type nullable text
    ),
    #"Added Analytical Role" = Table.AddColumn(
        #"Added Normalised DC Role",
        "Role",
        each [RoleKey],
        type nullable text
    ),
    #"Distinct DC Role Attributes" = Table.Distinct(#"Added Analytical Role")
in
    #"Distinct DC Role Attributes";

// Query: MinuteWorkerRoleAssignments_TABLE
// Purpose: Audits every retained explicit Master Roster role to DC Role assignment.
// Inputs: IMPORT Master and MW Roster Role Mapping Prepare.
// Output: One row per retained Roster Roles/DC Role assignment with category, direct-care percentage, row count and facility coverage.
// Notes: NA mappings are deliberately excluded from the MinuteWorker analysis; unmatched roles fail MW Input Check.
shared MinuteWorkerRoleAssignments_TABLE =
let
    Source = Table.SelectColumns(#"IMPORT Master", {"Location", "Role"}),
    #"Normalised Master Role Fields" = Table.TransformColumns(
        Source,
        {
            {
                "Location",
                each if _ = null then null else Text.Upper(Text.Start(Text.Trim(_), 2)),
                type nullable text
            },
            {"Role", each if _ = null then null else Text.Trim(_), type nullable text}
        }
    ),
    #"Renamed Original Role" = Table.RenameColumns(
        #"Normalised Master Role Fields",
        {{"Role", "MasterRosterRole"}}
    ),
    #"Added Roster Role Key" = Table.AddColumn(
        #"Renamed Original Role",
        "RosterRoleKey",
        each if [MasterRosterRole] = null then null else Text.Upper([MasterRosterRole]),
        type nullable text
    ),
    #"Merged Explicit Role Mapping" = Table.NestedJoin(
        #"Added Roster Role Key",
        {"RosterRoleKey"},
        #"MW Roster Role Mapping Prepare",
        {"RosterRoleKey"},
        "RoleMapping",
        JoinKind.Inner
    ),
    #"Expanded Explicit Role Mapping" = Table.ExpandTableColumn(
        #"Merged Explicit Role Mapping",
        "RoleMapping",
        {"DC Role", "DC Category", "MinuteCategory", "Direct Care %"},
        {"DC Role", "DC Category", "MinuteCategory", "Direct Care %"}
    ),
    #"Excluded NA Assignments" = Table.SelectRows(
        #"Expanded Explicit Role Mapping",
        each List.Contains({"RN", "OTHER"}, [DC Category])
    ),
    #"Summarised Role Assignments" = Table.Group(
        #"Excluded NA Assignments",
        {
            "MasterRosterRole", "DC Role", "DC Category", "MinuteCategory", "Direct Care %"
        },
        {
            {"MasterRosterRowCount", each Table.RowCount(_), Int64.Type},
            {
                "FacilityCount",
                each List.Count(List.Distinct(List.RemoveNulls([Location]))),
                Int64.Type
            },
            {
                "Facilities",
                each Text.Combine(List.Sort(List.Distinct(List.RemoveNulls([Location]))), ", "),
                type text
            }
        }
    ),
    #"Sorted Role Assignments" = Table.Sort(
        #"Summarised Role Assignments",
        {
            {"DC Role", Order.Ascending},
            {"MasterRosterRole", Order.Ascending}
        }
    )
in
    #"Sorted Role Assignments";

// Query: Master Prepare
// Purpose: Uses source net shift hours, applies the explicit roster-role mapping, excludes NA roles, and assigns shifts before historical aggregation.
// Inputs: IMPORT Master, MW Roster Role Mapping Prepare, ShftEnd Prepare, and shift-boundary scalars.
// Output: Historical roster rows tagged with original role, DC Role, DC Category, MinuteCategory and Direct Care %.
// Notes: Roster Hours is the downstream compatibility name for Shift Net Length; no additional meal deduction is applied.
shared #"Master Prepare" =
let
    Source = #"IMPORT Master",
    #"Normalised Location" = Table.TransformColumns(
        Source,
        {
            {
                "Location",
                each if _ = null then null else Text.Upper(Text.Start(Text.Trim(_), 2)),
                type nullable text
            }
        }
    ),
    #"Selected Historical Columns" = Table.SelectColumns(
        #"Normalised Location",
        {"Location", "Role", "Employee Code", "Week No", "Week Day", "Start Time", "End Time", "Shift Net Length"}
    ),
    // Select only the authoritative net field, then retain the existing hours interface for all historical calculations and checks.
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
    // NA and unmatched rows do not enter care allocation. MW Input Check
    // distinguishes intentional NA mappings from missing mappings.
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
    #"Filtered Facilities" = Table.SelectRows(
        #"Added Analytical Role",
        each [Location] <> "NR"
    ),
    #"Assigned Shift" = Table.AddColumn(
        #"Filtered Facilities",
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
    #"Expanded ShftEnd Prepare";

shared LocRoleWeekDaysHours = let
    Source = #"Master Prepare",
    #"added DAYOFWEEKS#" = Table.AddColumn(Source, "DayOfWeek", each List.PositionOf(
    {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},
    [Week Day]
) + 1),
    #"Sorted Rows" = Table.Sort(#"added DAYOFWEEKS#",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Grouped Rows" = Table.Group(#"Sorted Rows", {"Location", "Week No", "Role", "DayOfWeek", "Week Day", "Shift", "ShiftIndex"}, {{"Hours", each List.Sum([Roster Hours]), type nullable number}}),
    #"Sorted Rows1" = Table.Sort(#"Grouped Rows",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Inserted Merged Column1" = Table.AddColumn(#"Sorted Rows1", "WeekDayShift", each Text.Combine({Text.From([DayOfWeek], "en-AU"), [Week Day], [Shift]}, ""), type text),
    #"Inserted Merged Column" = Table.AddColumn(#"Inserted Merged Column1", "WeekAndDay", each Text.Combine({Text.From([Week No], "en-AU"), [Week Day]}, ""), type text),
    #"Inserted Merged Column2" = Table.AddColumn(#"Inserted Merged Column", "DayShift", each Text.Combine({[Week Day], [Shift]}, ""), type text)
in
    #"Inserted Merged Column2";

shared StartPM = let
    Source = #"ShftEnd Prepare",
    End = Source{0}[End]
in
    End;

shared StartNS = let
    Source = #"ShftEnd Prepare",
    End = Source{1}[End]
in
    End;

shared StartAM = let
    Source = #"ShftEnd Prepare",
    End = Source{2}[End]
in
    End;

// Query: INPUT TargetMinutes
// Purpose: Reads RN and ALL productive-care target hours per fortnight for every facility column.
// Inputs: Current-workbook table TargetMinutes.
// Output: MinuteType plus dynamically typed numeric facility columns.
// Notes: The existing TargetMinutes table name is retained; its values are hours per fortnight, confirmed 2026-09-07.
shared #"INPUT TargetMinutes" =
let
    Source = Excel.CurrentWorkbook(){[Name = "TargetMinutes"]}[Content],
    ColumnNames = Table.ColumnNames(Source),
    RequiredSchema =
        if not List.Contains(ColumnNames, "MinuteType") then
            error "TargetMinutes must contain a MinuteType column."
        else if List.Count(ColumnNames) = 1 then
            error "TargetMinutes must contain at least one facility column."
        else
            Source,
    FacilityColumns = List.RemoveItems(ColumnNames, {"MinuteType"}),
    #"Assigned Required Types" = Table.TransformColumnTypes(
        RequiredSchema,
        {{"MinuteType", type text}} &
            List.Transform(FacilityColumns, each {_, type number})
    )
in
    #"Assigned Required Types";

shared LocRoleWeekDaysHoursMATRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "WeekDayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[WeekDayShift]), "WeekDayShift", "Hours", List.Sum),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVEMATTRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}, {"ShiftIndex", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "DayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[DayShift]), "DayShift", "Hours", List.Average),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVETABLE = let
    Source = LocRoleWeekDaysShiftAVEMATTRIX,
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Location", "Role"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

shared LocRoleTOTAL = let
    Source = LocRoleWeekDaysShiftAVETABLE,
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Hours", each List.Sum([Value]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared #"LocRoleDayShift%" = let
    Source = Table.NestedJoin(LocRoleWeekDaysShiftAVETABLE, {"Location", "Role"}, LocRoleTOTAL, {"Location", "Role"}, "LocRoleTOTAL", JoinKind.LeftOuter),
    #"Expanded LocRoleTOTAL" = Table.ExpandTableColumn(Source, "LocRoleTOTAL", {"Hours"}, {"Hours"}),
    #"Inserted Division" = Table.AddColumn(#"Expanded LocRoleTOTAL", "LocRoleDayShift%", each [Value] / [Hours], type number),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Division",{{"LocRoleDayShift%", Percentage.Type}})
in
    #"Changed Type";

shared #"LocRoleDayShift%CHECK" = let
    Source = #"LocRoleDayShift%",
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Check", each List.Sum([#"LocRoleDayShift%"]), type number}})
in
    #"Grouped Rows";

shared #"LocRoleDatShift%MATRIX" = let
    Source = #"LocRoleDayShift%",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Value", "Hours"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Attribute]), "Attribute", "LocRoleDayShift%", List.Sum)
in
    #"Pivoted Column";

shared #"LocRoleWeekDaysHours (2)" = let
    Source = LocRoleWeekDaysHours
in
    Source;

// Query: MW Check Table
// Purpose: Converts validation records to the standard MinuteWorker check-table schema.
// Inputs: A list of validation records.
// Output: Check, Severity, Facility, MinuteCategory, Role, Actual, Expected, and Message.
shared #"MW Check Table" = (Rows as list) as table =>
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
    Result;

// Query: MW TargetMinutes Prepare
// Purpose: Converts RN and ALL productive-care hours per fortnight to daily and fortnight minute targets.
// Inputs: INPUT TargetMinutes.
// Output: Two rows per facility with original RN/ALL fortnight hours and daily/fortnight minutes for RN and OTHERS.
shared #"MW TargetMinutes Prepare" =
let
    Source = #"INPUT TargetMinutes",
    TargetPeriodDays = 14,
    MinutesPerHour = 60,
    FacilityColumns = List.RemoveItems(Table.ColumnNames(Source), {"MinuteType"}),
    // Expand every facility cell explicitly: Table.Unpivot drops null cells,
    // which would hide missing targets and wholly blank facility columns.
    #"Unpivoted Facility Targets" = Table.Combine(
        List.Transform(
            FacilityColumns,
            (FacilityName as text) as table =>
                Table.AddColumn(
                    Table.RenameColumns(
                        Table.SelectColumns(Source, {"MinuteType", FacilityName}),
                        {{FacilityName, "TargetFortnightHours"}}
                    ),
                    "Facility",
                    each FacilityName,
                    type text
                )
        )
    ),
    #"Normalised Target Keys" = Table.TransformColumns(
        #"Unpivoted Facility Targets",
        {
            {"MinuteType", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text},
            {"Facility", each Text.Upper(Text.Trim(_)), type text}
        }
    ),
    #"Summarised Facility Targets" = Table.Group(
        #"Normalised Target Keys",
        {"Facility"},
        {
            {"RNCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "RN")), Int64.Type},
            {"ALLCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "ALL")), Int64.Type},
            {"RNNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "RN" and [TargetFortnightHours] = null)), Int64.Type},
            {"ALLNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "ALL" and [TargetFortnightHours] = null)), Int64.Type},
            {
                "RNFortnightTargetHours",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "RN")[TargetFortnightHours]
                    in
                        if List.IsEmpty(Values) then null else List.Sum(Values),
                type nullable number
            },
            {
                "ALLFortnightTargetHours",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "ALL")[TargetFortnightHours]
                    in
                        if List.IsEmpty(Values) then null else List.Sum(Values),
                type nullable number
            },
            {
                "UnexpectedTypeCount",
                each Table.RowCount(
                    Table.SelectRows(
                        _,
                        each [MinuteType] = null or ([MinuteType] <> "RN" and [MinuteType] <> "ALL")
                    )
                ),
                Int64.Type
            }
        }
    ),
    // Convert hours to minutes once, before any daily allocation. These are
    // productive-care targets; the role Direct Care % is applied downstream.
    #"Added RN Fortnight Target" = Table.AddColumn(
        #"Summarised Facility Targets",
        "RNFortnightTargetMinutes",
        each
            if [RNFortnightTargetHours] = null then
                null
            else
                [RNFortnightTargetHours] * MinutesPerHour,
        type nullable number
    ),
    #"Added ALL Fortnight Target" = Table.AddColumn(
        #"Added RN Fortnight Target",
        "ALLFortnightTargetMinutes",
        each
            if [ALLFortnightTargetHours] = null then
                null
            else
                [ALLFortnightTargetHours] * MinutesPerHour,
        type nullable number
    ),
    // Retain the fortnight target divided by 14 as an average daily reference.
    // Actual weekday allocations vary with whole-period historical weights.
    #"Added RN Daily Target" = Table.AddColumn(
        #"Added ALL Fortnight Target",
        "RNDailyTargetMinutes",
        each [RNFortnightTargetMinutes] / TargetPeriodDays,
        type nullable number
    ),
    #"Added ALL Daily Target" = Table.AddColumn(
        #"Added RN Daily Target",
        "ALLDailyTargetMinutes",
        each [ALLFortnightTargetMinutes] / TargetPeriodDays,
        type nullable number
    ),
    // ALL includes RN, so the productive-minute pool available to all other
    // MinuteWorker roles is ALL less RN at both the daily and fortnight grain.
    #"Added Category Target Records" = Table.AddColumn(
        #"Added ALL Daily Target",
        "CategoryTargets",
        each {
            [
                MinuteCategory = "RN",
                CategoryDailyTargetMinutes = [RNDailyTargetMinutes],
                CategoryTargetMinutes = [RNFortnightTargetMinutes]
            ],
            [
                MinuteCategory = "OTHERS",
                CategoryDailyTargetMinutes =
                    if [RNDailyTargetMinutes] = null or [ALLDailyTargetMinutes] = null then
                        null
                    else
                        [ALLDailyTargetMinutes] - [RNDailyTargetMinutes],
                CategoryTargetMinutes =
                    if [RNFortnightTargetMinutes] = null or [ALLFortnightTargetMinutes] = null then
                        null
                    else
                        [ALLFortnightTargetMinutes] - [RNFortnightTargetMinutes]
            ]
        },
        type list
    ),
    #"Expanded Category Target Rows" = Table.ExpandListColumn(
        #"Added Category Target Records",
        "CategoryTargets"
    ),
    #"Expanded Category Target Values" = Table.ExpandRecordColumn(
        #"Expanded Category Target Rows",
        "CategoryTargets",
        {"MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"},
        {"MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"}
    )
in
    #"Expanded Category Target Values";

// Query: MW Fortnight Days
// Purpose: Defines the ordered 14-day display keys without merging corresponding weekdays.
// Output: FortnightWeek 1/2, DayOfWeek 1..7, Week Day, FortnightDay and FortnightDayIndex 1..14.
shared #"MW Fortnight Days" =
let
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
    Result;

// Query: MW Historical Weeks
// Purpose: Maps the two supplied source weeks to display weeks 1 and 2 while retaining original Week No.
// Inputs: Master Prepare.
// Output: Facility, original Week No, FortnightWeek and HistoricalWeeksInPeriod.
// Notes: Sorted numeric source Week No defines order. A non-two-week period fails MW Distribution Check; no weeks are silently discarded.
shared #"MW Historical Weeks" =
let
    Source = Table.RenameColumns(Table.SelectColumns(#"Master Prepare", {"Location", "Week No"}), {{"Location", "Facility"}}),
    ValidWeeks = Table.Distinct(Table.SelectRows(Source, each [Week No] <> null)),
    FacilityWeeks = Table.Group(ValidWeeks, {"Facility"}, {
        {"Weeks", each Table.AddIndexColumn(Table.Sort(Table.SelectColumns(_, {"Week No"}), {{"Week No", Order.Ascending}}),
            "FortnightWeek", 1, 1, Int64.Type), type table},
        {"HistoricalWeeksInPeriod", each Table.RowCount(_), Int64.Type}
    }),
    Result = Table.ExpandTableColumn(FacilityWeeks, "Weeks", {"Week No", "FortnightWeek"}, {"Week No", "FortnightWeek"})
in
    Result;

// Query: MW Historical WeekDayShift
// Purpose: Preserves each historical week's net roster hours and FTE with separate fortnight-day display keys.
// Inputs: Master Prepare, MW MinuteWorkers Prepare, and ShftEnd Prepare.
// Output: One row per observed facility/role, original Week No, weekday and AM/PM/NS shift.
// Notes: Complete weeks have explicit zero cells; missing cells in incomplete weeks remain null. No target scaling is applied.
shared #"MW Historical WeekDayShift" =
let
    History = Table.Buffer(Table.AddColumn(
        Table.RenameColumns(#"Master Prepare", {{"Location", "Facility"}}),
        "DayOfWeek",
        each List.PositionOf({"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}, [#"Week Day"]) + 1,
        Int64.Type
    )),
    ValidDays = Table.SelectRows(History, each [Week No] <> null and [DayOfWeek] >= 1 and [DayOfWeek] <= 7),
    // Coverage is measured at facility/week grain before adding absent cells.
    // Week No is retained unchanged so the two Mondays can be compared directly.
    WeekCoverage = Table.Group(ValidDays, {"Facility", "Week No"}, {
        {"HistoricalDaysPresent", each List.Count(List.Distinct([DayOfWeek])), Int64.Type}
    }),
    Roles = Table.Distinct(Table.SelectColumns(History, {"Facility", "Role"})),
    RoleAttributes = Table.ExpandTableColumn(
        Table.NestedJoin(Roles, {"Role"}, #"MW MinuteWorkers Prepare", {"Role"}, "Configuration", JoinKind.Inner),
        "Configuration", {"DC Role", "DC Category", "MinuteCategory", "Direct Care %"},
        {"DC Role", "DC Category", "MinuteCategory", "Direct Care %"}
    ),
    OrderedRoleWeeks = Table.ExpandTableColumn(
        Table.NestedJoin(RoleAttributes, {"Facility"}, #"MW Historical Weeks", {"Facility"}, "WeekKeys", JoinKind.Inner),
        "WeekKeys", {"Week No", "FortnightWeek", "HistoricalWeeksInPeriod"}, {"Week No", "FortnightWeek", "HistoricalWeeksInPeriod"}
    ),
    RoleWeeks = Table.ExpandTableColumn(
        Table.NestedJoin(OrderedRoleWeeks, {"Facility", "Week No"}, WeekCoverage, {"Facility", "Week No"}, "Weeks", JoinKind.LeftOuter),
        "Weeks", {"HistoricalDaysPresent"}, {"HistoricalDaysPresent"}
    ),
    Weekdays = #table(type table [DayOfWeek = Int64.Type, #"Week Day" = text], {
        {1, "Monday"}, {2, "Tuesday"}, {3, "Wednesday"}, {4, "Thursday"},
        {5, "Friday"}, {6, "Saturday"}, {7, "Sunday"}
    }),
    RoleWeekDays = Table.ExpandTableColumn(
        Table.AddColumn(RoleWeeks, "Days", each Weekdays), "Days",
        {"DayOfWeek", "Week Day"}, {"DayOfWeek", "Week Day"}
    ),
    AddedFortnightDayIndex = Table.AddColumn(RoleWeekDays, "FortnightDayIndex",
        each ([FortnightWeek] - 1) * 7 + [DayOfWeek], Int64.Type),
    AddedFortnightDay = Table.AddColumn(AddedFortnightDayIndex, "FortnightDay",
        each Text.From([FortnightWeek], "en-AU") & "-" & [#"Week Day"], type text),
    ShiftNames = #table(type table [Shift = text], {{"AM"}, {"PM"}, {"NS"}}),
    ShiftKeys = Table.ExpandTableColumn(
        Table.NestedJoin(ShiftNames, {"Shift"}, #"ShftEnd Prepare", {"Shift"}, "ShiftOrder", JoinKind.LeftOuter),
        "ShiftOrder", {"Index"}, {"ShiftIndex"}
    ),
    CompleteGrain = Table.ExpandTableColumn(
        Table.AddColumn(AddedFortnightDay, "Shifts", each ShiftKeys), "Shifts",
        {"Shift", "ShiftIndex"}, {"Shift", "ShiftIndex"}
    ),
    ObservedCells = Table.Group(ValidDays, {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, {
        {"ObservedRosterHours", each List.Sum([Roster Hours]), type nullable number},
        {"SourceRowCount", each Table.RowCount(_), Int64.Type},
        {"InvalidRosterRowCount", each Table.RowCount(Table.SelectRows(_, each [Roster Hours] = null or [Roster Hours] < 0)), Int64.Type}
    }),
    JoinedCells = Table.ExpandTableColumn(
        Table.NestedJoin(CompleteGrain,
            {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, ObservedCells,
            {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, "Observed", JoinKind.LeftOuter),
        "Observed", {"ObservedRosterHours", "SourceRowCount", "InvalidRosterRowCount"},
        {"ObservedRosterHours", "SourceRowCount", "InvalidRosterRowCount"}
    ),
    CellCounts = Table.ReplaceValue(JoinedCells, null, 0, Replacer.ReplaceValue, {"SourceRowCount", "InvalidRosterRowCount"}),
    AddedCoverageStatus = Table.AddColumn(CellCounts, "HistoricalCoverageStatus",
        each if [HistoricalDaysPresent] = 7 and [HistoricalWeeksInPeriod] = 2 then "PASS" else "INCOMPLETE", type text),
    AddedCellStatus = Table.AddColumn(AddedCoverageStatus, "HistoricalCellStatus",
        each if [InvalidRosterRowCount] > 0 then "ERROR"
        else if [SourceRowCount] > 0 then "OBSERVED"
        else if [HistoricalDaysPresent] = 7 then "ZERO" else "MISSING", type text),
    AddedRosterHours = Table.AddColumn(AddedCellStatus, "HistoricalRosterHours",
        each if [HistoricalCellStatus] = "ERROR" or [HistoricalCellStatus] = "MISSING" then null
        else if [SourceRowCount] = 0 then 0 else [ObservedRosterHours], type nullable number),
    // Historical FTE uses source Shift Net Length through Master Prepare's Roster Hours alias, without Direct Care % or target adjustment.
    AddedRosterFTE = Table.AddColumn(AddedRosterHours, "HistoricalRosterFTE",
        each [HistoricalRosterHours] / ShiftDuration, type nullable number),
    AddedProductiveHours = Table.AddColumn(AddedRosterFTE, "HistoricalProductiveHours",
        each [HistoricalRosterHours] * [#"Direct Care %"], type nullable number),
    AddedDayShift = Table.AddColumn(AddedProductiveHours, "DayShift", each [#"Week Day"] & [Shift], type text),
    AddedFortnightDayShift = Table.AddColumn(AddedDayShift, "FortnightDayShift", each [FortnightDay] & "-" & [Shift], type text),
    Result = Table.Sort(Table.RemoveColumns(AddedFortnightDayShift, {"ObservedRosterHours"}), {
        {"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Week No", Order.Ascending},
        {"DayOfWeek", Order.Ascending}, {"ShiftIndex", Order.Ascending}
    })
in
    Result;

// Query: MW Historical DayShift
// Purpose: Calculates unaveraged role/day/shift productive-hour shares across the complete source fortnight.
// Inputs: MW Historical WeekDayShift.
// Output: Positive-history cells at facility/role/FortnightDay/shift grain, with original Week No.
// Notes: No averaging occurs. All absent valid cells remain visible in the historical fortnight table.
shared #"MW Historical DayShift" =
let
    Source = Table.SelectRows(#"MW Historical WeekDayShift", each [HistoricalRosterHours] <> null and [HistoricalRosterHours] > 0),
    History = Table.Buffer(Table.AddColumn(Source, "RoleKey", each Text.Upper(Text.Trim([Role])), type text)),
    RoleDays = Table.Group(History, {"Facility", "MinuteCategory", "RoleKey", "FortnightDayIndex"}, {
        {"RoleDayHistoricalRosterHours", each List.Sum([HistoricalRosterHours]), type number},
        {"RoleDayHistoricalProductiveHours", each List.Sum([HistoricalProductiveHours]), type number}
    }),
    CategoryDays = Table.Group(History, {"Facility", "MinuteCategory", "FortnightDayIndex"}, {
        {"CategoryDayHistoricalProductiveHours", each List.Sum([HistoricalProductiveHours]), type number}
    }),
    CategoryPeriods = Table.Group(History, {"Facility", "MinuteCategory"}, {
        {"CategoryFortnightHistoricalProductiveHours", each List.Sum([HistoricalProductiveHours]), type number}
    }),
    WithRoleDays = Table.ExpandTableColumn(
        Table.NestedJoin(History, {"Facility", "MinuteCategory", "RoleKey", "FortnightDayIndex"},
            RoleDays, {"Facility", "MinuteCategory", "RoleKey", "FortnightDayIndex"}, "RoleDay", JoinKind.LeftOuter),
        "RoleDay", {"RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours"}),
    WithCategoryDays = Table.ExpandTableColumn(
        Table.NestedJoin(WithRoleDays, {"Facility", "MinuteCategory", "FortnightDayIndex"},
            CategoryDays, {"Facility", "MinuteCategory", "FortnightDayIndex"}, "CategoryDay", JoinKind.LeftOuter),
        "CategoryDay", {"CategoryDayHistoricalProductiveHours"}),
    WithCategoryPeriods = Table.ExpandTableColumn(
        Table.NestedJoin(WithCategoryDays, {"Facility", "MinuteCategory"},
            CategoryPeriods, {"Facility", "MinuteCategory"}, "CategoryPeriod", JoinKind.LeftOuter),
        "CategoryPeriod", {"CategoryFortnightHistoricalProductiveHours"}),
    SafeShare = (Numerator as nullable number, Denominator as nullable number) as nullable number =>
        if Denominator = null or Denominator <= 0 then null else Numerator / Denominator,
    // Conditional shares explain the mix within this particular week/day, not a pooled weekday.
    RoleShare = Table.AddColumn(WithCategoryPeriods, "RoleDayHistoryDistribution%",
        each SafeShare([RoleDayHistoricalProductiveHours], [CategoryDayHistoricalProductiveHours]), Percentage.Type),
    ShiftShare = Table.AddColumn(RoleShare, "RoleDayShiftDistribution%",
        each SafeShare([HistoricalRosterHours], [RoleDayHistoricalRosterHours]), Percentage.Type),
    CategoryDayShare = Table.AddColumn(ShiftShare, "CategoryDayRoleShiftDistribution%",
        each SafeShare([HistoricalProductiveHours], [CategoryDayHistoricalProductiveHours]), Percentage.Type),
    // The denominator includes all 14 days. These cell shares sum to 100% once per facility/category.
    FortnightCellShare = Table.AddColumn(CategoryDayShare, "CategoryFortnightRoleDayShiftDistribution%",
        each SafeShare([HistoricalProductiveHours], [CategoryFortnightHistoricalProductiveHours]), Percentage.Type),
    Result = Table.AddColumn(FortnightCellShare, "CategoryFortnightDayDistribution%",
        each SafeShare([CategoryDayHistoricalProductiveHours], [CategoryFortnightHistoricalProductiveHours]), Percentage.Type)
in
    Result;

// Query: MW Role Distribution
// Purpose: Exposes each role's conditional share within an individual fortnight day.
// Inputs: MW Historical DayShift.
// Output: Facility/category/role/FortnightDay with source week and productive-hour denominators.
shared #"MW Role Distribution" =
let
    Result = Table.Distinct(Table.SelectColumns(#"MW Historical DayShift", {
        "Facility", "MinuteCategory", "Role", "RoleKey", "DC Role", "DC Category", "Direct Care %",
        "Week No", "FortnightWeek", "DayOfWeek", "Week Day", "FortnightDayIndex", "FortnightDay",
        "RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours", "CategoryDayHistoricalProductiveHours",
        "RoleDayHistoryDistribution%", "CategoryFortnightDayDistribution%"
    }))
in
    Result;

// Query: MW Category Weekday Targets
// Purpose: Allocates the complete category target to 14 distinct fortnight days.
// Inputs: MW TargetMinutes Prepare, MW Fortnight Days, MW Historical Weeks and MW Historical DayShift.
// Output: Fourteen rows per facility/category; CategoryDailyTargetMinutes is informational only.
// Notes: Query name retained for existing connections; it no longer returns a seven-day average pattern.
shared #"MW Category Weekday Targets" =
let
    Targets = Table.SelectColumns(#"MW TargetMinutes Prepare",
        {"Facility", "MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"}),
    CompleteDays = Table.ExpandTableColumn(Table.AddColumn(Targets, "Days", each #"MW Fortnight Days"),
        "Days", {"FortnightWeek", "DayOfWeek", "Week Day", "FortnightDayIndex", "FortnightDay"}),
    WithSourceWeek = Table.ExpandTableColumn(
        Table.NestedJoin(CompleteDays, {"Facility", "FortnightWeek"}, #"MW Historical Weeks",
            {"Facility", "FortnightWeek"}, "SourceWeek", JoinKind.LeftOuter), "SourceWeek", {"Week No"}),
    Shares = Table.Distinct(Table.SelectColumns(#"MW Historical DayShift",
        {"Facility", "MinuteCategory", "FortnightDayIndex", "CategoryFortnightDayDistribution%"})),
    ExpandedShares = Table.ExpandTableColumn(
        Table.NestedJoin(WithSourceWeek, {"Facility", "MinuteCategory", "FortnightDayIndex"}, Shares,
            {"Facility", "MinuteCategory", "FortnightDayIndex"}, "DayShare", JoinKind.LeftOuter),
        "DayShare", {"CategoryFortnightDayDistribution%"}),
    // A zero-history category/day has no target weight. Incomplete source periods fail the publication gate.
    FilledShares = Table.ReplaceValue(ExpandedShares, null, 0, Replacer.ReplaceValue, {"CategoryFortnightDayDistribution%"}),
    Result = Table.AddColumn(FilledShares, "CategoryWeekdayTargetMinutes",
        each [CategoryTargetMinutes] * [#"CategoryFortnightDayDistribution%"], type nullable number)
in
    Result;

// Query: MW Role Targets
// Purpose: Allocates the full fortnight target to each distinct role/week/day and calculates separate weekly subtotals.
// Inputs: MW Role Distribution and MW TargetMinutes Prepare.
// Output: RoleDailyTargetMinutes, actual RoleWeeklyTargetMinutes for this week, and RoleTargetMinutes for all 14 days.
shared #"MW Role Targets" =
let
    WithCategory = Table.ExpandTableColumn(
        Table.NestedJoin(#"MW Role Distribution", {"Facility", "MinuteCategory"},
            #"MW TargetMinutes Prepare", {"Facility", "MinuteCategory"}, "Target", JoinKind.Inner),
        "Target", {"CategoryDailyTargetMinutes", "CategoryTargetMinutes"}),
    DailyTargets = Table.Buffer(Table.AddColumn(WithCategory, "RoleDailyTargetMinutes",
        each [CategoryTargetMinutes] * [#"CategoryFortnightDayDistribution%"] * [#"RoleDayHistoryDistribution%"], type number)),
    WeeklyTargets = Table.Group(DailyTargets, {"Facility", "MinuteCategory", "RoleKey", "FortnightWeek"}, {
        {"RoleWeeklyTargetMinutes", each List.Sum([RoleDailyTargetMinutes]), type number}
    }),
    PeriodTargets = Table.Group(DailyTargets, {"Facility", "MinuteCategory", "RoleKey"}, {
        {"RoleTargetMinutes", each List.Sum([RoleDailyTargetMinutes]), type number}
    }),
    WithWeek = Table.ExpandTableColumn(
        Table.NestedJoin(DailyTargets, {"Facility", "MinuteCategory", "RoleKey", "FortnightWeek"},
            WeeklyTargets, {"Facility", "MinuteCategory", "RoleKey", "FortnightWeek"}, "WeekTarget", JoinKind.LeftOuter),
        "WeekTarget", {"RoleWeeklyTargetMinutes"}),
    WithPeriod = Table.ExpandTableColumn(
        Table.NestedJoin(WithWeek, {"Facility", "MinuteCategory", "RoleKey"},
            PeriodTargets, {"Facility", "MinuteCategory", "RoleKey"}, "PeriodTarget", JoinKind.LeftOuter),
        "PeriodTarget", {"RoleTargetMinutes"}),
    // Informational average only: never used to allocate or replace an individual day's requirement.
    Result = Table.AddColumn(WithPeriod, "RoleAverageDailyTargetMinutes", each [RoleTargetMinutes] / 14, type number)
in
    Result;

// Query: MW DayShift Allocation
// Purpose: Converts the complete fortnight productive-care budget to required roster FTE for each distinct week/day/shift.
// Inputs: MW Historical DayShift and MW Role Targets.
// Output: Facility/role/original Week No/FortnightDay/shift with unrounded configured standard-FTE equivalents.
// Notes: No corresponding-weekday averaging and no divide-by-two. WeekdayShift is now a unique fortnight-day/shift key.
shared #"MW DayShift Allocation" =
let
    WithTargets = Table.ExpandTableColumn(
        Table.NestedJoin(#"MW Historical DayShift", {"Facility", "MinuteCategory", "RoleKey", "FortnightDayIndex"},
            #"MW Role Targets", {"Facility", "MinuteCategory", "RoleKey", "FortnightDayIndex"}, "Target", JoinKind.Inner),
        "Target", {"CategoryDailyTargetMinutes", "CategoryTargetMinutes", "RoleDailyTargetMinutes",
            "RoleWeeklyTargetMinutes", "RoleAverageDailyTargetMinutes", "RoleTargetMinutes"}),
    // Shares cover the entire fortnight. Scale once by the entire fortnight target.
    AllocatedCare = Table.AddColumn(WithTargets, "WeekdayShiftTargetMinutes",
        each [CategoryTargetMinutes] * [#"CategoryFortnightRoleDayShiftDistribution%"], type number),
    TwoStageCheck = Table.AddColumn(AllocatedCare, "TwoStageAllocationVarianceMinutes",
        each [WeekdayShiftTargetMinutes] - [RoleDailyTargetMinutes] * [#"RoleDayShiftDistribution%"], type number),
    RosterMinutes = Table.AddColumn(TwoStageCheck, "WeekdayShiftRosterMinutes",
        each if [#"Direct Care %"] = null or [#"Direct Care %"] <= 0 then null
        else [WeekdayShiftTargetMinutes] / [#"Direct Care %"], type nullable number),
    // Convert allocated roster minutes to hours before dividing by the standard-FTE hours.
    RequiredFTE = Table.AddColumn(RosterMinutes, "FTE", each ([WeekdayShiftRosterMinutes] / 60) / ShiftDuration, type nullable number),
    AddedShiftKey = Table.AddColumn(RequiredFTE, "WeekdayShift", each [FortnightDay] & "-" & [Shift], type text),
    Result = Table.Sort(AddedShiftKey, {
        {"Facility", Order.Ascending}, {"MinuteCategory", Order.Ascending}, {"Role", Order.Ascending},
        {"FortnightDayIndex", Order.Ascending}, {"ShiftIndex", Order.Ascending}
    })
in
    Result;

// Query: MW FTE Profile Comparison
// Purpose: Pairs actual redistributed FTE with master-roster FTE and independently checks constant-factor profile alignment.
// Inputs: MW Historical WeekDayShift, MW DayShift Allocation and MW TargetMinutes Prepare.
// Output: One historical facility/category/role/original week/day/shift cell, preserving all original columns and zeros.
// Notes: The expected factor is category target productive hours / historical productive hours across the full fortnight.
shared #"MW FTE Profile Comparison" =
let
    History = Table.Buffer(#"MW Historical WeekDayShift"),
    // Compute the scalar independently of the allocation shares. Direct Care %
    // cancels when converting productive allocations back to roster FTE.
    HistoricalCategories = Table.Group(History, {"Facility", "MinuteCategory"}, {
        {"ProfileHistoricalProductiveHours", each List.Sum([HistoricalProductiveHours]), type nullable number}
    }),
    Targets = Table.SelectColumns(#"MW TargetMinutes Prepare", {"Facility", "MinuteCategory", "CategoryTargetMinutes"}),
    WithHistoryTotals = Table.ExpandTableColumn(
        Table.NestedJoin(History, {"Facility", "MinuteCategory"}, HistoricalCategories,
            {"Facility", "MinuteCategory"}, "HistoryTotal", JoinKind.LeftOuter),
        "HistoryTotal", {"ProfileHistoricalProductiveHours"}),
    WithTargets = Table.ExpandTableColumn(
        Table.NestedJoin(WithHistoryTotals, {"Facility", "MinuteCategory"}, Targets,
            {"Facility", "MinuteCategory"}, "Target", JoinKind.LeftOuter),
        "Target", {"CategoryTargetMinutes"}),
    // Keep week in the join so the two Tuesdays cannot be combined. Do not
    // expand a duplicate allocation join: expose its count and fail alignment.
    JoinedAllocation = Table.NestedJoin(WithTargets,
        {"Facility", "MinuteCategory", "Role", "Week No", "FortnightDayIndex", "Shift"},
        #"MW DayShift Allocation",
        {"Facility", "MinuteCategory", "Role", "Week No", "FortnightDayIndex", "Shift"},
        "Redistribution", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(JoinedAllocation, "RedistributionMatchCount", each Table.RowCount([Redistribution]), Int64.Type),
    // Only valid zero-history cells with a configured target may become zero.
    // Missing positive-history allocations and missing targets stay null.
    WithRedistributedFTE = Table.AddColumn(WithMatchCount, "RedistributedRosterFTE", each
        if [CategoryTargetMinutes] = null or [HistoricalCoverageStatus] <> "PASS" or [HistoricalRosterFTE] = null then null
        else if [RedistributionMatchCount] = 1 then [Redistribution]{0}[FTE]
        else if [RedistributionMatchCount] = 0 and [HistoricalRosterFTE] = 0 then 0
        else null, type nullable number),
    WithExpectedFactor = Table.AddColumn(WithRedistributedFTE, "ExpectedRedistributionFactor", each
        if [CategoryTargetMinutes] = null or [ProfileHistoricalProductiveHours] = null or [ProfileHistoricalProductiveHours] <= 0 then null
        else [CategoryTargetMinutes] / 60 / [ProfileHistoricalProductiveHours], type nullable number),
    WithActualFactor = Table.AddColumn(WithExpectedFactor, "RedistributionFactor", each
        if [HistoricalRosterFTE] = null or [HistoricalRosterFTE] <= 0 then null
        else [RedistributedRosterFTE] / [HistoricalRosterFTE], type nullable number),
    WithVariance = Table.AddColumn(WithActualFactor, "ProfileVarianceFTE", each
        if [HistoricalRosterFTE] = 0 and [RedistributedRosterFTE] = 0 then 0
        else [RedistributedRosterFTE] - [HistoricalRosterFTE] * [ExpectedRedistributionFactor], type nullable number),
    WithStatus = Table.AddColumn(WithVariance, "ProfileAlignmentStatus", each
        if [HistoricalCoverageStatus] <> "PASS" or [HistoricalRosterFTE] = null then "ERROR"
        else if [CategoryTargetMinutes] = null then "NO TARGET"
        else if [RedistributionMatchCount] > 1 or [RedistributedRosterFTE] = null or [ProfileVarianceFTE] = null then "ERROR"
        else if Number.Abs([ProfileVarianceFTE]) <= (0.000001 / 60) / ShiftDuration then
            if [HistoricalRosterFTE] = 0 then "PASS ZERO" else "PASS"
        else "ERROR", type text),
    Result = Table.Sort(Table.RemoveColumns(WithStatus, {"Redistribution"}), {
        {"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"FortnightDayIndex", Order.Ascending}, {"ShiftIndex", Order.Ascending}
    })
in
    Result;

// Query: MW Role Fortnight Day Totals
// Purpose: Makes all 14 role/day cells explicit and sums shifts without pooling the two weeks.
// Inputs: MW DayShift Allocation, MW Role Targets, MW Fortnight Days and MW Historical Weeks.
// Output: One row per allocated facility/category/role/FortnightDay, including zero-allocation days.
shared #"MW Role Fortnight Day Totals" =
let
    Allocation = Table.Buffer(#"MW DayShift Allocation"),
    Roles = Table.Distinct(Table.SelectColumns(Allocation,
        {"Facility", "MinuteCategory", "Role", "DC Role", "DC Category", "Direct Care %", "RoleTargetMinutes", "RoleAverageDailyTargetMinutes"})),
    Days = Table.ExpandTableColumn(Table.AddColumn(Roles, "Days", each #"MW Fortnight Days"),
        "Days", {"FortnightWeek", "DayOfWeek", "Week Day", "FortnightDayIndex", "FortnightDay"}),
    WithSourceWeek = Table.ExpandTableColumn(
        Table.NestedJoin(Days, {"Facility", "FortnightWeek"}, #"MW Historical Weeks",
            {"Facility", "FortnightWeek"}, "SourceWeek", JoinKind.LeftOuter), "SourceWeek", {"Week No"}),
    DailyAllocation = Table.Group(Allocation, {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, {
        {"AllocatedDayProductiveMinutes", each List.Sum([WeekdayShiftTargetMinutes]), type number},
        {"AllocatedDayRosterMinutes", each List.Sum([WeekdayShiftRosterMinutes]), type nullable number},
        {"DayFTEShiftTotal", each List.Sum([FTE]), type nullable number},
        {"TargetAMFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "AM")[FTE]), type nullable number},
        {"TargetPMFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "PM")[FTE]), type nullable number},
        {"TargetNSFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "NS")[FTE]), type nullable number},
        {"DailyShiftDistributionTotal%", each List.Sum([#"RoleDayShiftDistribution%"]), type nullable number},
        {"MaximumAbsoluteTwoStageVarianceMinutes", each List.Max(List.Transform([TwoStageAllocationVarianceMinutes], Number.Abs)), type nullable number},
        {"InvalidAllocationCount", each Table.RowCount(Table.SelectRows(_, each [FTE] = null or [WeekdayShiftTargetMinutes] = null)), Int64.Type}
    }),
    Measures = {"AllocatedDayProductiveMinutes", "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
        "TargetAMFTE", "TargetPMFTE", "TargetNSFTE", "DailyShiftDistributionTotal%",
        "MaximumAbsoluteTwoStageVarianceMinutes", "InvalidAllocationCount"},
    ExpandedAllocation = Table.ExpandTableColumn(
        Table.NestedJoin(WithSourceWeek, {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, DailyAllocation,
            {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, "Actual", JoinKind.LeftOuter), "Actual", Measures),
    ExpandedTargets = Table.ExpandTableColumn(
        Table.NestedJoin(ExpandedAllocation, {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, #"MW Role Targets",
            {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, "Expected", JoinKind.LeftOuter),
        "Expected", {"RoleDailyTargetMinutes"}),
    // Only absent allocation cells become zero. An invalid present row is separately counted and fails checks.
    Result = Table.ReplaceValue(ExpandedTargets, null, 0, Replacer.ReplaceValue, Measures & {"RoleDailyTargetMinutes"})
in
    Result;

// Query: MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK
// Purpose: Reconciles each distinct role/day, each actual source week and the complete fortnight.
// Inputs: MW Role Fortnight Day Totals and MW Role Targets.
// Output: Fourteen rows per allocated facility/role. Weekly fields refer to this row's FortnightWeek, not an averaged week.
shared MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK =
let
    Tolerance = 0.000001,
    Days = Table.Buffer(#"MW Role Fortnight Day Totals"),
    Weeks = Table.Group(Days, {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, {
        {"AllocatedWeeklyProductiveMinutes", each List.Sum([AllocatedDayProductiveMinutes]), type number},
        {"WeeklyFTEShiftTotal", each List.Sum([DayFTEShiftTotal]), type number}
    }),
    ExpectedWeeks = Table.Group(#"MW Role Targets", {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, {
        {"ExpectedWeeklyProductiveMinutes", each List.Sum([RoleDailyTargetMinutes]), type number}
    }),
    Periods = Table.Group(Days, {"Facility", "MinuteCategory", "Role"}, {
        {"ReconstructedFortnightProductiveMinutes", each List.Sum([AllocatedDayProductiveMinutes]), type number},
        {"FortnightFTEShiftTotal", each List.Sum([DayFTEShiftTotal]), type number}
    }),
    WithWeek = Table.ExpandTableColumn(
        Table.NestedJoin(Days, {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, Weeks,
            {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, "Week", JoinKind.LeftOuter),
        "Week", {"AllocatedWeeklyProductiveMinutes", "WeeklyFTEShiftTotal"}),
    WithExpectedWeek = Table.ExpandTableColumn(
        Table.NestedJoin(WithWeek, {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, ExpectedWeeks,
            {"Facility", "MinuteCategory", "Role", "FortnightWeek"}, "ExpectedWeek", JoinKind.LeftOuter),
        "ExpectedWeek", {"ExpectedWeeklyProductiveMinutes"}),
    FilledExpectedWeek = Table.ReplaceValue(WithExpectedWeek, null, 0, Replacer.ReplaceValue, {"ExpectedWeeklyProductiveMinutes"}),
    WithPeriod = Table.ExpandTableColumn(
        Table.NestedJoin(FilledExpectedWeek, {"Facility", "MinuteCategory", "Role"}, Periods,
            {"Facility", "MinuteCategory", "Role"}, "Period", JoinKind.LeftOuter),
        "Period", {"ReconstructedFortnightProductiveMinutes", "FortnightFTEShiftTotal"}),
    DailyVariance = Table.AddColumn(WithPeriod, "DayAllocationVarianceMinutes",
        each [AllocatedDayProductiveMinutes] - [RoleDailyTargetMinutes], type number),
    WeeklyVariance = Table.AddColumn(DailyVariance, "WeeklyVarianceMinutes",
        each [AllocatedWeeklyProductiveMinutes] - [ExpectedWeeklyProductiveMinutes], type number),
    PeriodVariance = Table.AddColumn(WeeklyVariance, "FortnightVarianceMinutes",
        each [ReconstructedFortnightProductiveMinutes] - [RoleTargetMinutes], type number),
    AddedStatus = Table.AddColumn(PeriodVariance, "Status", each
        if [InvalidAllocationCount] = 0 and Number.Abs([DayAllocationVarianceMinutes]) <= Tolerance and
            Number.Abs([WeeklyVarianceMinutes]) <= Tolerance and Number.Abs([FortnightVarianceMinutes]) <= Tolerance and
            [MaximumAbsoluteTwoStageVarianceMinutes] <= Tolerance and
            (Number.Abs([#"DailyShiftDistributionTotal%"] - 1) <= Tolerance or
                ([#"DailyShiftDistributionTotal%"] = 0 and [RoleDailyTargetMinutes] = 0))
        then "PASS" else "ERROR", type text),
    AddedMessage = Table.AddColumn(AddedStatus, "CheckMessage",
        each [FortnightDay] & ": reconcile this day, its actual week, and the sum of all 14 days; no averaging or doubling.", type text),
    Result = Table.Sort(AddedMessage, {{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"FortnightDayIndex", Order.Ascending}})
in
    Result;

// Query: MinuteWorkersFTE_CATEGORY_DAILY_CHECK
// Purpose: Reconciles weighted RN/OTHERS targets for 14 separate days, two actual weeks and the fortnight.
// Inputs: MW Category Weekday Targets and MW DayShift Allocation.
// Output: Fourteen rows per facility/category; the two weekly totals may differ.
shared MinuteWorkersFTE_CATEGORY_DAILY_CHECK =
let
    Tolerance = 0.000001,
    DayActuals = Table.Group(#"MW DayShift Allocation", {"Facility", "MinuteCategory", "FortnightDayIndex"}, {
        {"AllocatedDayProductiveMinutes", each List.Sum([WeekdayShiftTargetMinutes]), type number}
    }),
    WithActuals = Table.ExpandTableColumn(
        Table.NestedJoin(#"MW Category Weekday Targets", {"Facility", "MinuteCategory", "FortnightDayIndex"}, DayActuals,
            {"Facility", "MinuteCategory", "FortnightDayIndex"}, "Actual", JoinKind.LeftOuter),
        "Actual", {"AllocatedDayProductiveMinutes"}),
    Days = Table.Buffer(Table.ReplaceValue(WithActuals, null, 0, Replacer.ReplaceValue, {"AllocatedDayProductiveMinutes"})),
    Weeks = Table.Group(Days, {"Facility", "MinuteCategory", "FortnightWeek"}, {
        {"AllocatedWeeklyProductiveMinutes", each List.Sum([AllocatedDayProductiveMinutes]), type number},
        {"ExpectedWeeklyProductiveMinutes", each List.Sum([CategoryWeekdayTargetMinutes]), type number}
    }),
    Periods = Table.Group(Days, {"Facility", "MinuteCategory"}, {
        {"ReconstructedFortnightProductiveMinutes", each List.Sum([AllocatedDayProductiveMinutes]), type number}
    }),
    WithWeeks = Table.ExpandTableColumn(
        Table.NestedJoin(Days, {"Facility", "MinuteCategory", "FortnightWeek"}, Weeks,
            {"Facility", "MinuteCategory", "FortnightWeek"}, "Week", JoinKind.LeftOuter),
        "Week", {"AllocatedWeeklyProductiveMinutes", "ExpectedWeeklyProductiveMinutes"}),
    WithPeriod = Table.ExpandTableColumn(
        Table.NestedJoin(WithWeeks, {"Facility", "MinuteCategory"}, Periods,
            {"Facility", "MinuteCategory"}, "Period", JoinKind.LeftOuter), "Period", {"ReconstructedFortnightProductiveMinutes"}),
    DayVariance = Table.AddColumn(WithPeriod, "DailyVarianceMinutes",
        each [AllocatedDayProductiveMinutes] - [CategoryWeekdayTargetMinutes], type number),
    WeekVariance = Table.AddColumn(DayVariance, "WeeklyVarianceMinutes",
        each [AllocatedWeeklyProductiveMinutes] - [ExpectedWeeklyProductiveMinutes], type number),
    PeriodVariance = Table.AddColumn(WeekVariance, "FortnightVarianceMinutes",
        each [ReconstructedFortnightProductiveMinutes] - [CategoryTargetMinutes], type number),
    AddedStatus = Table.AddColumn(PeriodVariance, "Status", each
        if [CategoryTargetMinutes] <> null and Number.Abs([DailyVarianceMinutes]) <= Tolerance and
            Number.Abs([WeeklyVarianceMinutes]) <= Tolerance and Number.Abs([FortnightVarianceMinutes]) <= Tolerance
        then "PASS" else "ERROR", type text),
    Result = Table.Sort(AddedStatus, {{"Facility", Order.Ascending}, {"MinuteCategory", Order.Ascending}, {"FortnightDayIndex", Order.Ascending}})
in
    Result;

// Query: MinuteWorkersFTE_HISTORICAL_DAY_CHECK
// Purpose: Compares master-roster FTE with required FTE for each distinct week/day; corresponding weekdays are never averaged.
// Inputs: MW Historical WeekDayShift, MW Role Fortnight Day Totals and Master Prepare.
// Output: Historical and target daily/shift FTE plus headcount/coverage context at facility/role/FortnightDay grain.
// Notes: Prior average/across-weeks comparison fields are retired; charts must use the separate-day fields below.
shared MinuteWorkersFTE_HISTORICAL_DAY_CHECK =
let
    History = Table.Group(#"MW Historical WeekDayShift",
        {"Facility", "MinuteCategory", "Role", "Week No", "FortnightWeek", "DayOfWeek", "Week Day", "FortnightDayIndex", "FortnightDay"}, {
        {"HistoricalDailyRosterFTE", each if List.NonNullCount([HistoricalRosterFTE]) <> Table.RowCount(_) then null else List.Sum([HistoricalRosterFTE]), type nullable number},
        {"HistoricalAMRosterFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "AM")[HistoricalRosterFTE]), type nullable number},
        {"HistoricalPMRosterFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "PM")[HistoricalRosterFTE]), type nullable number},
        {"HistoricalNSRosterFTE", each List.Sum(Table.SelectRows(_, each [Shift] = "NS")[HistoricalRosterFTE]), type nullable number},
        {"HistoricalCoverageStatus", each if List.AllTrue(List.Transform([HistoricalCoverageStatus], each _ = "PASS")) then "PASS" else "INCOMPLETE", type text},
        {"InvalidRosterRowCount", each List.Sum([InvalidRosterRowCount]), Int64.Type}
    }),
    RawHeadcounts = Table.Group(#"Master Prepare", {"Location", "Role", "Week No", "Week Day"}, {
        {"HistoricalDailyDistinctWorkers", each List.Count(List.Distinct(List.RemoveNulls(Table.SelectRows(_, each [Roster Hours] > 0)[Employee Code]))), Int64.Type},
        {"HistoricalRowsMissingEmployeeCode", each Table.RowCount(Table.SelectRows(_, each [Employee Code] = null)), Int64.Type}
    }),
    WithHeadcount = Table.ExpandTableColumn(
        Table.NestedJoin(History, {"Facility", "Role", "Week No", "Week Day"}, RawHeadcounts,
            {"Location", "Role", "Week No", "Week Day"}, "Workers", JoinKind.LeftOuter),
        "Workers", {"HistoricalDailyDistinctWorkers", "HistoricalRowsMissingEmployeeCode"}),
    WithCounts = Table.ReplaceValue(WithHeadcount, null, 0, Replacer.ReplaceValue,
        {"HistoricalDailyDistinctWorkers", "HistoricalRowsMissingEmployeeCode"}),
    // The target staging has explicit zero-allocation days. A missing target join remains null (not fabricated zero).
    WithTarget = Table.ExpandTableColumn(
        Table.NestedJoin(WithCounts, {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, #"MW Role Fortnight Day Totals",
            {"Facility", "MinuteCategory", "Role", "FortnightDayIndex"}, "Target", JoinKind.LeftOuter),
        "Target", {"DayFTEShiftTotal", "TargetAMFTE", "TargetPMFTE", "TargetNSFTE", "RoleDailyTargetMinutes", "AllocatedDayProductiveMinutes"},
        {"TargetDailyFTEShiftTotal", "TargetAMFTE", "TargetPMFTE", "TargetNSFTE", "RoleDailyTargetMinutes", "TargetDailyProductiveMinutes"}),
    Variance = Table.AddColumn(WithTarget, "TargetVsHistoricalRosterFTE",
        each [TargetDailyFTEShiftTotal] - [HistoricalDailyRosterFTE], type nullable number),
    TargetStatus = Table.AddColumn(Variance, "RoleDayTargetStatus", each
        if [TargetDailyFTEShiftTotal] = null then "NO TARGET"
        else if Number.Abs([TargetDailyProductiveMinutes] - [RoleDailyTargetMinutes]) <= 0.000001 then "PASS" else "ERROR", type text),
    Result = Table.Sort(TargetStatus, {{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"FortnightDayIndex", Order.Ascending}})
in
    Result;

// Query: MW Input Check
// Purpose: Validates the explicit roster-role mapping, retained DC-role attributes, and RN/ALL target inputs.
// Inputs: INPUT MinuteWorkers, MW Roster Role Mapping Prepare, MW MinuteWorkers Prepare, and MW TargetMinutes Prepare.
// Output: Error records for invalid required inputs.
shared #"MW Input Check" =
let
    // A completely empty target table must fail before its facility columns
    // disappear from downstream grouping.
    TargetPresenceChecks = #"MW Check Table"(
        if Table.IsEmpty(#"INPUT TargetMinutes") then
            {[
                Check = "Fortnight target inputs are present",
                Severity = "Error",
                Actual = 0,
                Expected = 1,
                Message = "TargetMinutes must contain RN and ALL productive-care hours per fortnight."
            ]}
        else
            {}
    ),
    Mapping = Table.Buffer(#"MW Roster Role Mapping Prepare"),
    MappingKeyCounts = Table.Group(
        Mapping,
        {"RosterRoleKey"},
        {{"MappingCount", each Table.RowCount(_), Int64.Type}}
    ),
    InvalidMappingKeyCounts = Table.SelectRows(
        MappingKeyCounts,
        each [RosterRoleKey] = null or [RosterRoleKey] = "" or [MappingCount] <> 1
    ),
    MappingKeyChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidMappingKeyCounts), each [
        Check = "Roster role mapping key is present and unique",
        Severity = "Error", Facility = null, MinuteCategory = null, Role = [RosterRoleKey],
        Actual = Number.From([MappingCount]), Expected = 1,
        Message = "Each non-blank normalised Roster Roles value must occur exactly once, including NA mappings."
    ])),
    InvalidCategories = Table.SelectRows(
        Mapping,
        each [DC Category] = null or not List.Contains({"RN", "OTHER", "NA"}, [DC Category])
    ),
    CategoryChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidCategories), each [
        Check = "DC Category is recognised",
        Severity = "Error", Facility = null, MinuteCategory = [MinuteCategory], Role = [DC Role],
        Actual = null, Expected = null,
        Message = "Roster role '" & (if [Roster Roles] = null then "<blank>" else [Roster Roles]) &
            "' has DC Category '" & (if [DC Category] = null then "<blank>" else [DC Category]) &
            "'; expected RN, OTHER, or NA."
    ])),
    RetainedMappings = Table.SelectRows(Mapping, each List.Contains({"RN", "OTHER"}, [DC Category])),
    InvalidRetainedRoles = Table.SelectRows(
        RetainedMappings,
        each [RoleKey] = null or [RoleKey] = ""
    ),
    RetainedRoleChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidRetainedRoles), each [
        Check = "Retained mapping has a DC Role",
        Severity = "Error", Facility = null, MinuteCategory = [MinuteCategory], Role = [DC Role],
        Actual = 0, Expected = 1,
        Message = "Roster role '" & (if [Roster Roles] = null then "<blank>" else [Roster Roles]) &
            "' is retained by DC Category but has no DC Role."
    ])),
    InvalidDirectCare = Table.SelectRows(
        RetainedMappings,
        each not [DirectCareInputIsNumeric] or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1
    ),
    DirectCareChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidDirectCare), each [
        Check = "Direct Care percentage",
        Severity = "Error", Facility = null, MinuteCategory = [MinuteCategory], Role = [DC Role],
        Actual = [#"Direct Care %"],
        Expected = 1,
        Message = "Retained RN/OTHER mappings require numeric Direct Care % greater than 0 and no greater than 100%."
    ])),
    DCRoleAttributeCounts = Table.Group(
        RetainedMappings,
        {"RoleKey"},
        {{"AttributeCount", each Table.RowCount(Table.Distinct(Table.SelectColumns(_, {"DC Category", "Direct Care %"}))), Int64.Type}}
    ),
    ConflictingDCRoles = Table.SelectRows(DCRoleAttributeCounts, each [RoleKey] <> null and [AttributeCount] <> 1),
    DCRoleConsistencyChecks = #"MW Check Table"(List.Transform(Table.ToRecords(ConflictingDCRoles), each [
        Check = "DC Role attributes are consistent",
        Severity = "Error", Facility = null, MinuteCategory = null, Role = [RoleKey],
        Actual = Number.From([AttributeCount]), Expected = 1,
        Message = "Every Roster Roles row mapped to the same DC Role must use one DC Category and Direct Care %."
    ])),
    TargetSummary = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {
                "Facility", "RNCount", "ALLCount", "RNNullCount", "ALLNullCount",
                "RNFortnightTargetHours", "ALLFortnightTargetHours", "UnexpectedTypeCount"
            }
        )
    ),
    TargetChecks = #"MW Check Table"(
        List.Combine(
            List.Transform(
                Table.ToRecords(TargetSummary),
                each List.RemoveNulls({
                    if [RNCount] <> 1 then [
                        Check = "RN target row count",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = Number.From([RNCount]),
                        Expected = 1,
                        Message = "Each facility must have exactly one RN target row."
                    ] else null,
                    if [ALLCount] <> 1 then [
                        Check = "ALL target row count",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = Number.From([ALLCount]),
                        Expected = 1,
                        Message = "Each facility must have exactly one ALL target row."
                    ] else null,
                    if [UnexpectedTypeCount] <> 0 then [
                        Check = "Target minute types",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = Number.From([UnexpectedTypeCount]),
                        Expected = 0,
                        Message = "MinuteType values must be RN or ALL."
                    ] else null,
                    if [RNNullCount] <> 0 or [RNFortnightTargetHours] = null or [RNFortnightTargetHours] < 0 then [
                        Check = "RN target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNFortnightTargetHours],
                        Expected = 0,
                        Message = "RN productive-care target hours per fortnight must be numeric and non-negative."
                    ] else null,
                    if [ALLNullCount] <> 0 or [ALLFortnightTargetHours] = null or [ALLFortnightTargetHours] < 0 then [
                        Check = "ALL target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = [ALLFortnightTargetHours],
                        Expected = 0,
                        Message = "ALL productive-care target hours per fortnight must be numeric and non-negative."
                    ] else null,
                    if
                        [RNFortnightTargetHours] <> null and
                        [ALLFortnightTargetHours] <> null and
                        [RNFortnightTargetHours] > [ALLFortnightTargetHours]
                    then [
                        Check = "RN target not greater than ALL",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNFortnightTargetHours],
                        Expected = [ALLFortnightTargetHours],
                        Message = "RN fortnight target hours cannot exceed ALL fortnight target hours."
                    ] else null
                })
            )
        )
    ),
    Result = Table.Combine({
        TargetPresenceChecks, MappingKeyChecks, CategoryChecks, RetainedRoleChecks,
        DirectCareChecks, DCRoleConsistencyChecks, TargetChecks
    })
in
    Result;

// Query: MW Distribution Check
// Purpose: Validates historical source rows, complete periods, and conditional and whole-period distribution denominators.
// Inputs: Master Prepare, LocRoleWeekDaysHours, MW Historical DayShift, and MW TargetMinutes Prepare.
// Output: Pass, warning, and error records for historical allocation readiness.
shared #"MW Distribution Check" =
let
    Tolerance = 0.000001,
    HistoricalDayShift = Table.Buffer(#"MW Historical DayShift"),
    RoleDistribution = Table.Buffer(
        Table.Distinct(
            Table.SelectColumns(
                HistoricalDayShift,
                {
                    "Facility", "MinuteCategory", "Role", "RoleKey", "DC Role", "DC Category",
                    "Direct Care %", "DayOfWeek", "Week Day", "FortnightDayIndex", "FortnightDay",
                    "RoleDayHistoricalRosterHours",
                    "RoleDayHistoricalProductiveHours",
                    "CategoryDayHistoricalProductiveHours",
                    "RoleDayHistoryDistribution%"
                }
            )
        )
    ),
    TargetCategories = Table.Buffer(#"MW TargetMinutes Prepare"),
    HistoricalInput = Table.Buffer(
        Table.SelectColumns(
            LocRoleWeekDaysHours,
            {
                "Location", "Week No", "Role", "DayOfWeek", "Week Day",
                "Shift", "ShiftIndex", "Hours"
            }
        )
    ),
    Weekdays = #table(
        type table [DayOfWeek = Int64.Type, #"Week Day" = text],
        {
            {1, "Monday"},
            {2, "Tuesday"},
            {3, "Wednesday"},
            {4, "Thursday"},
            {5, "Friday"},
            {6, "Saturday"},
            {7, "Sunday"}
        }
    ),

    // Invalid source keys and negative/null aggregated hours must fail before
    // they can create negative or incomplete allocation shares.
    // Check raw hours as well: aggregation can otherwise conceal a negative
    // adjustment or ignore a null beside valid hours in the same cell.
    InvalidRawRosterRows = Table.SelectRows(#"Master Prepare", each [Roster Hours] = null or [Roster Hours] < 0),
    RawRosterHoursChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidRawRosterRows), each [
        Check = "Historical raw roster hours are valid", Severity = "Error",
        Facility = [Location], MinuteCategory = null, Role = [Role],
        Actual = [Roster Hours], Expected = 0,
        Message = "Each eligible original roster row must have non-null, non-negative hours before grouping."
    ])),
    // Verify that the completed graph table preserves source totals and row
    // counts. Zero-fill must not duplicate observed cells or lose source rows.
    HistoricalGrid = Table.Buffer(#"MW Historical WeekDayShift"),
    HistoricalGridCellCounts = Table.Group(HistoricalGrid,
        {"Facility", "Role", "Week No", "DayOfWeek", "Shift"},
        {{"CellCount", each Table.RowCount(_), Int64.Type}}),
    HistoricalGridKeyChecks = #"MW Check Table"(List.Transform(
        Table.ToRecords(Table.SelectRows(HistoricalGridCellCounts, each [CellCount] <> 1)), each [
            Check = "Historical graph table has unique cells", Severity = "Error",
            Facility = [Facility], MinuteCategory = null, Role = [Role], Actual = [CellCount], Expected = 1,
            Message = "Expected exactly one historical row per facility/role/original week/weekday/shift."
        ])),
    RawPeriodTotals = Table.Group(#"Master Prepare", {"Location", "Week No"}, {
        {"SourceHours", each List.Sum([Roster Hours]), type nullable number},
        {"SourceRows", each Table.RowCount(_), Int64.Type}
    }),
    GridPeriodTotals = Table.Group(HistoricalGrid, {"Facility", "Week No"}, {
        {"GridHours", each List.Sum([HistoricalRosterHours]), type nullable number},
        {"GridSourceRows", each List.Sum([SourceRowCount]), Int64.Type}
    }),
    ComparedPeriodTotals = Table.ExpandTableColumn(
        Table.NestedJoin(RawPeriodTotals, {"Location", "Week No"}, GridPeriodTotals,
            {"Facility", "Week No"}, "Grid", JoinKind.LeftOuter),
        "Grid", {"GridHours", "GridSourceRows"}, {"GridHours", "GridSourceRows"}),
    HistoricalGridTotalChecks = #"MW Check Table"(List.Transform(Table.ToRecords(ComparedPeriodTotals), each [
        Check = "Historical graph table preserves source hours and rows",
        Severity = if [GridHours] <> null and [SourceHours] <> null and
            Number.Abs([GridHours] - [SourceHours]) <= Tolerance and [GridSourceRows] = [SourceRows]
            then "Pass" else "Error",
        Facility = [Location], MinuteCategory = null, Role = null,
        Actual = [GridHours], Expected = [SourceHours],
        Message = "Historical graph hours and contributing row counts must match the unaveraged source for week " &
            (if [Week No] = null then "<null>" else Text.From([Week No], "en-AU")) & "."
    ])),
    InvalidHistoricalInput = Table.SelectRows(
        HistoricalInput,
        each
            [Week No] = null or
            [DayOfWeek] = null or [DayOfWeek] < 1 or [DayOfWeek] > 7 or
            not List.Contains(
                {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"},
                [#"Week Day"]
            ) or
            [Shift] = null or
            not List.Contains({"AM", "PM", "NS"}, [Shift]) or
            [ShiftIndex] = null or
            [Hours] = null or [Hours] < 0
    ),
    HistoricalInputChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(InvalidHistoricalInput),
            each [
                Check = "Historical allocation source row is valid",
                Severity = "Error",
                Facility = [Location],
                MinuteCategory = null,
                Role = [Role],
                Actual = [Hours],
                Expected = 0,
                Message =
                    "History requires a non-null Week No, a recognised weekday/shift, " &
                    "a ShiftIndex, and non-negative roster hours. Week=" &
                    (if [Week No] = null then "<null>" else Text.From([Week No], "en-AU")) &
                    ", day=" &
                    (if [#"Week Day"] = null then "<null>" else [#"Week Day"]) &
                    ", shift=" &
                    (if [Shift] = null then "<null>" else [Shift]) & "."
            ]
        )
    ),
    TargetFacilityList = Table.Distinct(
        Table.SelectColumns(TargetCategories, {"Facility"})
    ),
    HistoricalFacilityPeriods = Table.Distinct(
        Table.SelectColumns(
            Table.SelectRows(HistoricalInput, each [Week No] <> null),
            {"Location", "Week No"}
        )
    ),
    // A single fortnight requires exactly two source weeks. Never average, pick
    // a pair from a longer extract, or invent an absent week.
    WeekCounts = Table.Group(HistoricalFacilityPeriods, {"Location"}, {
        {"ObservedWeeks", each Table.RowCount(_), Int64.Type}
    }),
    TargetWeekCounts = Table.ExpandTableColumn(
        Table.NestedJoin(TargetFacilityList, {"Facility"}, WeekCounts, {"Location"}, "WeekCount", JoinKind.LeftOuter),
        "WeekCount", {"ObservedWeeks"}),
    FortnightPeriodChecks = #"MW Check Table"(List.Transform(Table.ToRecords(TargetWeekCounts), each [
        Check = "Historical target facility contains exactly two source weeks",
        Severity = if [ObservedWeeks] = 2 then "Pass" else "Error",
        Facility = [Facility], MinuteCategory = null, Role = null,
        Actual = if [ObservedWeeks] = null then 0 else [ObservedWeeks], Expected = 2,
        Message = "Supply exactly two complete, uniquely identified source weeks; no selection or averaging is performed."
    ])),
    TargetHistoricalPeriods = Table.RemoveColumns(
        Table.NestedJoin(
            HistoricalFacilityPeriods,
            {"Location"},
            TargetFacilityList,
            {"Facility"},
            "TargetFacility",
            JoinKind.Inner
        ),
        {"TargetFacility"}
    ),
    #"Added Expected Historical Weekdays" = Table.AddColumn(
        TargetHistoricalPeriods,
        "Weekdays",
        each Weekdays,
        type table [DayOfWeek = Int64.Type, #"Week Day" = text]
    ),
    #"Expanded Expected Historical Weekdays" = Table.ExpandTableColumn(
        #"Added Expected Historical Weekdays",
        "Weekdays",
        {"DayOfWeek", "Week Day"},
        {"DayOfWeek", "Week Day"}
    ),
    ObservedHistoricalWeekdays = Table.Distinct(
        Table.SelectColumns(
            Table.SelectRows(
                HistoricalInput,
                each
                    [Week No] <> null and
                    [DayOfWeek] <> null and
                    [DayOfWeek] >= 1 and [DayOfWeek] <= 7
            ),
            {"Location", "Week No", "DayOfWeek"}
        )
    ),
    MissingHistoricalWeekdays = Table.NestedJoin(
        #"Expanded Expected Historical Weekdays",
        {"Location", "Week No", "DayOfWeek"},
        ObservedHistoricalWeekdays,
        {"Location", "Week No", "DayOfWeek"},
        "ObservedDay",
        JoinKind.LeftAnti
    ),
    HistoricalCoverageChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(MissingHistoricalWeekdays),
            each [
                Check = "Historical facility period has every weekday",
                Severity = "Error",
                Facility = [Location],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message =
                    "Week " & Text.From([Week No], "en-AU") & " has no eligible " &
                    [#"Week Day"] &
                    " MinuteWorker history; complete periods are required for separate fortnight-day allocation."
            ]
        )
    ),

    BuildDistributionRangeChecks =
        (Rows as table, ColumnName as text, CheckName as text) as table =>
            let
                InvalidRows = Table.SelectRows(
                    Rows,
                    each
                        let
                            DistributionValue = Record.Field(_, ColumnName)
                        in
                            DistributionValue = null or
                            Number.IsNaN(DistributionValue) or
                            DistributionValue < -Tolerance or
                            DistributionValue > 1 + Tolerance
                ),
                Result = #"MW Check Table"(
                    List.Transform(
                        Table.ToRecords(InvalidRows),
                        each [
                            Check = CheckName,
                            Severity = "Error",
                            Facility = [Facility],
                            MinuteCategory = [MinuteCategory],
                            Role = [Role],
                            Actual = Record.Field(_, ColumnName),
                            Expected = 1,
                            Message =
                                ColumnName & " must be between 0% and 100% on " &
                                [#"FortnightDay"] & "."
                        ]
                    )
                )
            in
                Result,
    RoleDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.Distinct(
            Table.SelectColumns(
                RoleDistribution,
                {
                    "Facility", "MinuteCategory", "Role", "FortnightDayIndex", "FortnightDay",
                    "RoleDayHistoryDistribution%"
                }
            )
        ),
        "RoleDayHistoryDistribution%",
        "Role weekday history share is within range"
    ),
    ShiftDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.SelectColumns(
            HistoricalDayShift,
            {
                "Facility", "MinuteCategory", "Role", "FortnightDayIndex", "FortnightDay",
                "RoleDayShiftDistribution%"
            }
        ),
        "RoleDayShiftDistribution%",
        "Role weekday shift share is within range"
    ),
    CategoryDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.SelectColumns(
            HistoricalDayShift,
            {
                "Facility", "MinuteCategory", "Role", "FortnightDayIndex", "FortnightDay",
                "CategoryDayRoleShiftDistribution%"
            }
        ),
        "CategoryDayRoleShiftDistribution%",
        "Category weekday role/shift share is within range"
    ),
    WholePeriodDistributionRangeChecks = BuildDistributionRangeChecks(
        HistoricalDayShift,
        "CategoryFortnightRoleDayShiftDistribution%",
        "Category whole-period role/day/shift share is within range"
    ),
    WholePeriodDistributionTotals = Table.Group(HistoricalDayShift, {"Facility", "MinuteCategory"}, {
        {"Actual", each List.Sum([#"CategoryFortnightRoleDayShiftDistribution%"]), type nullable number}
    }),
    WholePeriodDistributionChecks = #"MW Check Table"(List.Transform(Table.ToRecords(WholePeriodDistributionTotals), each [
        Check = "Category whole-period distribution totals 100%",
        Severity = if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then "Pass" else "Error",
        Facility = [Facility], MinuteCategory = [MinuteCategory], Role = null,
        Actual = [Actual], Expected = 1,
        Message = "CategoryFortnightRoleDayShiftDistribution% must total 100% across all roles, weekdays and shifts in the fortnight."
    ])),

    RoleDistributionTotals = Table.Group(
        RoleDistribution,
        {"Facility", "MinuteCategory", "FortnightDayIndex", "FortnightDay"},
        {{"Actual", each List.Sum([#"RoleDayHistoryDistribution%"]), type nullable number}}
    ),
    RoleDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(RoleDistributionTotals),
            each [
                Check = "Role history distribution totals 100%",
                Severity =
                    if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [Actual],
                Expected = 1,
                Message =
                    "RoleDayHistoryDistribution% must total 100% for " &
                    [#"FortnightDay"] & " independently."
            ]
        )
    ),

    DayShiftDistributionTotals = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory", "Role", "FortnightDayIndex", "FortnightDay"},
        {{"Actual", each List.Sum([#"RoleDayShiftDistribution%"]), type nullable number}}
    ),
    DayShiftDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(DayShiftDistributionTotals),
            each [
                Check = "Role day/shift distribution totals 100%",
                Severity =
                    if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = [Role],
                Actual = [Actual],
                Expected = 1,
                Message =
                    "RoleDayShiftDistribution% must total 100% for role " &
                    [Role] & " on " & [#"FortnightDay"] & " independently."
            ]
        )
    ),

    CategoryDayDistributionTotals = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory", "FortnightDayIndex", "FortnightDay"},
        {
            {
                "Actual",
                each List.Sum([#"CategoryDayRoleShiftDistribution%"]),
                type nullable number
            }
        }
    ),
    CategoryDayDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryDayDistributionTotals),
            each [
                Check = "Category weekday role/shift distribution totals 100%",
                Severity =
                    if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [Actual],
                Expected = 1,
                Message =
                    "CategoryDayRoleShiftDistribution% must total 100% for " &
                    [#"FortnightDay"] & " independently."
            ]
        )
    ),

    CategoryHistory = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory"},
        {
            {
                "CategoryFortnightHistoricalProductiveHours",
                each List.Sum([HistoricalProductiveHours]),
                type nullable number
            }
        }
    ),
    TargetCategoryGrain = Table.Distinct(
        Table.SelectColumns(
            TargetCategories,
            {
                "Facility", "MinuteCategory", "CategoryDailyTargetMinutes",
                "CategoryTargetMinutes"
            }
        )
    ),
    // A category may legitimately have zero weight on a particular weekday.
    // Only absence of history for the entire positive-target category is fatal;
    // missing source weekdays at facility/week grain are checked separately.
    TargetHistoryJoin = Table.ExpandTableColumn(
        Table.NestedJoin(
            TargetCategoryGrain,
            {"Facility", "MinuteCategory"},
            CategoryHistory,
            {"Facility", "MinuteCategory"},
            "CategoryHistory",
            JoinKind.LeftOuter
        ),
        "CategoryHistory",
        {"CategoryFortnightHistoricalProductiveHours"},
        {"CategoryFortnightHistoricalProductiveHours"}
    ),
    MissingHistory = Table.SelectRows(
        TargetHistoryJoin,
        each
            [CategoryTargetMinutes] > 0 and
            (
                [CategoryFortnightHistoricalProductiveHours] = null or
                [CategoryFortnightHistoricalProductiveHours] <= 0
            )
    ),
    MissingHistoryChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(MissingHistory),
            each [
                Check = "Positive target has history",
                Severity = "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual =
                    if [CategoryFortnightHistoricalProductiveHours] = null then
                        0
                    else
                        [CategoryFortnightHistoricalProductiveHours],
                Expected = 1,
                Message =
                    "Facility " & [Facility] & ", category " & [MinuteCategory] &
                    " has a positive fortnight target of " &
                    Text.From([CategoryTargetMinutes], "en-AU") &
                    " minutes but no eligible productive MinuteWorker hours across the historical period."
            ]
        )
    ),

    MasterRoleKeys = Table.Distinct(
        Table.SelectColumns(
            Table.AddColumn(
                HistoricalInput,
                "RoleKey",
                each if [Role] = null then null else Text.Upper(Text.Trim([Role])),
                type nullable text
            ),
            {"RoleKey"}
        )
    ),
    UnusedDCRoles = Table.NestedJoin(
        #"MW MinuteWorkers Prepare",
        {"RoleKey"},
        MasterRoleKeys,
        {"RoleKey"},
        "HistoryRole",
        JoinKind.LeftAnti
    ),
    UnusedDCRoleChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(UnusedDCRoles),
            each [
                Check = "Configured DC Role has eligible history",
                Severity = "Warning",
                Facility = null,
                MinuteCategory = [MinuteCategory],
                Role = [RoleKey],
                Actual = 0,
                Expected = 1,
                Message =
                    "DC Role '" & (if [DC Role] = null then "<blank>" else [DC Role]) &
                    "' is configured as " & [DC Category] &
                    " but no retained Master Roster rows map to it."
            ]
        )
    ),

    TargetFacilities = Table.Distinct(Table.SelectColumns(TargetCategories, {"Facility"})),
    HistoricalFacilities = Table.Distinct(Table.SelectColumns(CategoryHistory, {"Facility"})),
    TargetOnlyFacilities = Table.NestedJoin(
        TargetFacilities,
        {"Facility"},
        HistoricalFacilities,
        {"Facility"},
        "HistoryFacility",
        JoinKind.LeftAnti
    ),
    TargetFacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(TargetOnlyFacilities),
            each [
                Check = "Target facility has MinuteWorker history",
                Severity = "Warning",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message = "The target facility has no matching MinuteWorker history."
            ]
        )
    ),
    HistoryOnlyFacilities = Table.NestedJoin(
        HistoricalFacilities,
        {"Facility"},
        TargetFacilities,
        {"Facility"},
        "TargetFacility",
        JoinKind.LeftAnti
    ),
    HistoricalFacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(HistoryOnlyFacilities),
            each [
                Check = "Historical facility has target",
                Severity = "Warning",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message = "The historical facility has no TargetMinutes column and is omitted from the FTE output."
            ]
        )
    ),

    Result = Table.Combine(
        {
            HistoricalInputChecks,
            RawRosterHoursChecks,
            HistoricalGridKeyChecks,
            HistoricalGridTotalChecks,
            HistoricalCoverageChecks,
            FortnightPeriodChecks,
            RoleDistributionRangeChecks,
            ShiftDistributionRangeChecks,
            CategoryDistributionRangeChecks,
            WholePeriodDistributionRangeChecks,
            WholePeriodDistributionChecks,
            RoleDistributionChecks,
            DayShiftDistributionChecks,
            CategoryDayDistributionChecks,
            MissingHistoryChecks,
            UnusedDCRoleChecks,
            TargetFacilityChecks,
            HistoricalFacilityChecks
        }
    )
in
    Result;

// Query: MW Daily Allocation Check
// Purpose: Gates publication on distinct-day, actual-week and complete-fortnight reconciliation.
// Inputs: MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK and MinuteWorkersFTE_CATEGORY_DAILY_CHECK.
// Output: Standard check rows; diagnostic messages identify FortnightDay.
shared #"MW Daily Allocation Check" =
let
    RoleChecks = #"MW Check Table"(List.Transform(Table.ToRecords(MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK), each [
        Check = "Role fortnight-day/week/period allocation reconciles",
        Severity = if [Status] = "PASS" then "Pass" else "Error",
        Facility = [Facility], MinuteCategory = [MinuteCategory], Role = [Role],
        Actual = [AllocatedDayProductiveMinutes], Expected = [RoleDailyTargetMinutes],
        Message = [CheckMessage]
    ])),
    CategoryChecks = #"MW Check Table"(List.Transform(Table.ToRecords(MinuteWorkersFTE_CATEGORY_DAILY_CHECK), each [
        Check = "Category fortnight-day/week/period allocation reconciles",
        Severity = if [Status] = "PASS" then "Pass" else "Error",
        Facility = [Facility], MinuteCategory = [MinuteCategory], Role = null,
        Actual = [AllocatedDayProductiveMinutes], Expected = [CategoryWeekdayTargetMinutes],
        Message = [FortnightDay] & ": weighted day target, actual week and complete fortnight must reconcile."
    ])),
    Result = Table.Combine({RoleChecks, CategoryChecks})
in
    Result;

// Query: MW PreAllocation Check
// Purpose: Combines input and historical-distribution checks that can run before allocation.
// Inputs: MW Input Check and MW Distribution Check.
// Output: Validation rows; Severity Error prevents the allocation checks from being evaluated.
shared #"MW PreAllocation Check" =
let
    Result = Table.Combine({#"MW Input Check", #"MW Distribution Check"})
in
    Result;

// Query: MW Fortnight Hours Check
// Purpose: Independently reconciles allocated FTE to the original RN and ALL hours-per-fortnight inputs.
// Inputs: INPUT TargetMinutes and MW DayShift Allocation; evaluate after MW PreAllocation Check passes.
// Output: RN and ALL checks per facility with Actual and Expected in productive-care hours per fortnight.
// Notes: Reads raw input hours directly, avoiding the minute conversion used to produce allocation targets.
shared #"MW Fortnight Hours Check" =
let
    // Match the existing 0.000001-minute tolerance, expressed in hours.
    ToleranceHours = 0.000001 / 60,
    RawTargets = Table.Buffer(#"INPUT TargetMinutes"),
    FacilityColumns = List.RemoveItems(Table.ColumnNames(RawTargets), {"MinuteType"}),
    NormalisedTargets = Table.TransformColumns(
        RawTargets,
        {{"MinuteType", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text}}
    ),
    Allocation = Table.Buffer(
        Table.SelectColumns(
            #"MW DayShift Allocation",
            {"Facility", "MinuteCategory", "Direct Care %", "FTE"}
        )
    ),
    CheckRecords = List.Combine(
        List.Transform(
            FacilityColumns,
            (FacilityColumn as text) as list =>
                List.Transform(
                    {"RN", "ALL"},
                    (TargetType as text) as record =>
                        let
                            FacilityKey = Text.Upper(Text.Trim(FacilityColumn)),
                            InputValues = Table.Column(
                                Table.SelectRows(NormalisedTargets, each [MinuteType] = TargetType),
                                FacilityColumn
                            ),
                            ExpectedHours =
                                if List.Count(InputValues) = 1 then InputValues{0} else null,
                            EligibleAllocation = Table.SelectRows(
                                Allocation,
                                each [Facility] = FacilityKey and
                                    (TargetType = "ALL" or [MinuteCategory] = "RN")
                            ),
                            // All 14 days are present once. Convert roster FTE back to
                            // productive hours without repeating or doubling any day.
                            ReconstructedHours = List.Transform(
                                Table.ToRecords(EligibleAllocation),
                                each [FTE] * ShiftDuration * [#"Direct Care %"]
                            ),
                            HasInvalidAllocation = List.NonNullCount(ReconstructedHours) <> List.Count(ReconstructedHours),
                            ActualHours =
                                if List.IsEmpty(ReconstructedHours) then 0 else List.Sum(ReconstructedHours),
                            Passed =
                                if ExpectedHours = null or HasInvalidAllocation then false
                                else Number.Abs(ActualHours - ExpectedHours) <= ToleranceHours
                        in
                            [
                                Check = "Allocated FTE reconciles to original fortnight hours",
                                Severity = if Passed then "Pass" else "Error",
                                Facility = FacilityKey,
                                MinuteCategory = if TargetType = "RN" then "RN" else null,
                                Role = null,
                                Actual = ActualHours,
                                Expected = ExpectedHours,
                                Message = TargetType &
                                    ": sum of all 14 days of FTE x configured standard-FTE hours x Direct Care % must equal original TargetMinutes hours per fortnight."
                            ]
                )
        )
    ),
    Result = #"MW Check Table"(CheckRecords)
in
    Result;

// Query: MW FTE Profile Alignment Check
// Purpose: Validates that redistributed role/day/shift FTE is a single category-wide scalar multiple of master-roster FTE.
// Inputs: MW FTE Profile Comparison; evaluate after MW PreAllocation Check passes.
// Output: Standard check records; errors block publication and missing targets remain warnings.
shared #"MW FTE Profile Alignment Check" =
let
    Result = #"MW Check Table"(List.Transform(Table.ToRecords(#"MW FTE Profile Comparison"), each [
        Check = "Redistributed FTE preserves historical profile",
        Severity = if Text.StartsWith([ProfileAlignmentStatus], "PASS") then "Pass"
            else if [ProfileAlignmentStatus] = "NO TARGET" then "Warning" else "Error",
        Facility = [Facility], MinuteCategory = [MinuteCategory], Role = [Role],
        Actual = [RedistributedRosterFTE],
        Expected = if [HistoricalRosterFTE] = 0 then 0 else [HistoricalRosterFTE] * [ExpectedRedistributionFactor],
        Message = [FortnightDay] & " / " & [Shift] & ": " & [ProfileAlignmentStatus] &
            ". Redistributed FTE must equal historical FTE times the same facility/category scalar across both weeks."
    ]))
in
    Result;

// Query: MW Publication Check
// Purpose: Combines pre-allocation and post-allocation checks for a complete publication gate.
// Inputs: MW PreAllocation Check, MW Daily Allocation Check, MW Fortnight Hours Check, and MW FTE Profile Alignment Check.
// Output: Standard validation rows; any Severity Error blocks TABLE, MATRIX and the paired historical/redistributed profile output.
shared #"MW Publication Check" =
let
    PreAllocationChecks = Table.Buffer(#"MW PreAllocation Check"),
    FatalPreAllocation = Table.SelectRows(
        PreAllocationChecks,
        each [Severity] = "Error"
    ),
    Result =
        if Table.RowCount(FatalPreAllocation) > 0 then
            PreAllocationChecks
        else
            Table.Combine({PreAllocationChecks, #"MW Daily Allocation Check", #"MW Fortnight Hours Check", #"MW FTE Profile Alignment Check"})
in
    Result;

// Query: MinuteWorkersFTE_TABLE
// Purpose: Publishes the validated 14-day MinuteWorker FTE allocation without weekday averaging.
// Inputs: MW Publication Check and MW DayShift Allocation.
// Output: One row per facility, role, FortnightDay, and shift with unrounded configured standard-FTE equivalents in FTE.
shared MinuteWorkersFTE_TABLE =
let
    RaiseValidationError = (ErrorTitle as text, ValidationRows as table) as any =>
        let
            // Include failing keys so Excel identifies the exact input,
            // historical denominator, or weekday allocation without opening Details.
            AddedErrorContext = Table.AddColumn(
                ValidationRows,
                "ErrorContext",
                each Text.Combine(
                    List.RemoveNulls(
                        {
                            [Check],
                            if [Facility] = null then null else "Facility=" & [Facility],
                            if [MinuteCategory] = null then null else "MinuteCategory=" & [MinuteCategory],
                            if [Role] = null then null else "Role=" & [Role],
                            [Message]
                        }
                    ),
                    " | "
                ),
                type text
            ),
            ErrorMessage = Text.Combine(
                List.Distinct(AddedErrorContext[ErrorContext]),
                "#(lf)"
            )
        in
            error Error.Record(ErrorTitle, ErrorMessage, ValidationRows),
    // The publication query evaluates allocation checks only after input and
    // history checks pass, and includes reconciliation to original input hours.
    PublicationValidation = Table.Buffer(#"MW Publication Check"),
    FatalValidation = Table.SelectRows(
        PublicationValidation,
        each [Severity] = "Error"
    ),
    Result =
        if Table.RowCount(FatalValidation) > 0 then
            RaiseValidationError(
                "MinuteWorkersFTE publication validation failed",
                FatalValidation
            )
        else
            #"MW DayShift Allocation"
in
    Result;

// Query: MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE
// Purpose: Publishes paired master-roster and redistributed FTE profiles for each role, shift and distinct fortnight day.
// Inputs: MW FTE Profile Comparison and MW Publication Check.
// Output: Original historical columns plus RedistributedRosterFTE, expected/actual scalar factors and alignment diagnostics.
// Notes: Existing row grain is preserved. Plot Week Day with FortnightWeek and the two FTE measures as four series.
// Notes: Historical lines dashed, redistributed lines solid; keep one role/category per scalar comparison. No averaging or normalization.
shared MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE =
let
    FatalChecks = Table.SelectRows(#"MW Publication Check", each [Severity] = "Error"),
    Result = if Table.RowCount(FatalChecks) > 0 then
        error Error.Record("FTE profile comparison validation failed",
            "Resolve publication/profile checks before plotting redistributed FTE; inspect MW Historical WeekDayShift for raw history.", FatalChecks)
        else #"MW FTE Profile Comparison"
in
    Result;

// Query: MinuteWorkersFTE_MATRIX
// Purpose: Pivots the validated table to one FTE column per distinct fortnight day and shift (up to 42 columns).
// Inputs: MinuteWorkersFTE_TABLE.
// Output: One row per facility, MinuteCategory, and role.
shared MinuteWorkersFTE_MATRIX =
let
    Source = Table.Buffer(
        Table.SelectColumns(
            MinuteWorkersFTE_TABLE,
            {
                "Facility", "MinuteCategory", "Role", "DC Role", "DC Category",
                "Direct Care %", "FortnightDayIndex", "ShiftIndex", "WeekdayShift", "FTE"
            }
        )
    ),
    MatrixSource = Table.SelectColumns(
        Source,
        {
            "Facility", "MinuteCategory", "Role", "DC Role", "DC Category",
            "Direct Care %", "WeekdayShift", "FTE"
        }
    ),
    OrderedWeekdayShifts = List.Distinct(
        Table.Sort(
            Table.SelectColumns(Source, {"FortnightDayIndex", "ShiftIndex", "WeekdayShift"}),
            {
                {"FortnightDayIndex", Order.Ascending},
                {"ShiftIndex", Order.Ascending}
            }
        )[WeekdayShift]
    ),
    Result = Table.Pivot(
        MatrixSource,
        OrderedWeekdayShifts,
        "WeekdayShift",
        "FTE",
        List.Sum
    )
in
    Result;

// Query: MinuteWorkersFTE_CHECK
// Purpose: Reports input, distribution, category, and facility reconciliation results.
// Inputs: MW PreAllocation Check, MW Daily Allocation Check, MW Fortnight Hours Check, MW FTE Profile Alignment Check, MW DayShift Allocation, and MW TargetMinutes Prepare.
// Output: Validation rows reconciling RN/OTHERS/ALL minutes plus independent RN/ALL original fortnight hours.
shared MinuteWorkersFTE_CHECK =
let
    Tolerance = 0.000001,
    PreAllocationChecks = Table.Buffer(#"MW PreAllocation Check"),
    FatalValidation = Table.SelectRows(
        PreAllocationChecks,
        each [Severity] = "Error"
    ),
    DailyAllocationChecks = #"MW Daily Allocation Check",

    // Reconstruct productive minutes from all 14 allocated days once, with no factor of two.
    AllocationWithProductiveMinutes = Table.Buffer(
        Table.AddColumn(
            Table.SelectColumns(
                #"MW DayShift Allocation",
                {
                    "Facility", "MinuteCategory", "CategoryTargetMinutes",
                    "Direct Care %", "FTE"
                }
            ),
            "ReconciledProductiveMinutes",
            each [FTE] * ShiftDuration * [#"Direct Care %"] * 60,
            type number
        )
    ),
    CategoryReconciliation = Table.Group(
        AllocationWithProductiveMinutes,
        {"Facility", "MinuteCategory"},
        {
            {
                "Actual",
                each List.Sum([ReconciledProductiveMinutes]),
                type number
            },
            {
                "Expected",
                each List.Max([CategoryTargetMinutes]),
                type number
            }
        }
    ),
    CategoryChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryReconciliation),
            each [
                Check = "Allocated productive minutes reconcile to category target",
                Severity =
                    if Number.Abs([Actual] - [Expected]) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [Actual],
                Expected = [Expected],
                Message = "Sum of all 14 days of FTE x ShiftDuration hours x Direct Care % x 60 must equal the fortnight category target."
            ]
        )
    ),

    FacilityActual = Table.Group(
        AllocationWithProductiveMinutes,
        {"Facility"},
        {
            {
                "Actual",
                each List.Sum([ReconciledProductiveMinutes]),
                type number
            }
        }
    ),
    FacilityTargets = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {"Facility", "ALLFortnightTargetMinutes"}
        )
    ),
    FacilityReconciliation = Table.ExpandTableColumn(
        Table.NestedJoin(
            FacilityTargets,
            {"Facility"},
            FacilityActual,
            {"Facility"},
            "Allocation",
            JoinKind.LeftOuter
        ),
        "Allocation",
        {"Actual"},
        {"Actual"}
    ),
    FacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(FacilityReconciliation),
            each [
                Check = "Facility productive minutes reconcile to ALL target",
                Severity =
                    if Number.Abs(
                        (if [Actual] = null then 0 else [Actual]) - [ALLFortnightTargetMinutes]
                    ) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = if [Actual] = null then 0 else [Actual],
                Expected = [ALLFortnightTargetMinutes],
                Message = "RN plus OTHERS productive minutes must equal the facility ALL fortnight target."
            ]
        )
    ),

    Result =
        if Table.RowCount(FatalValidation) > 0 then
            PreAllocationChecks
        else
            Table.Combine(
                {
                    PreAllocationChecks,
                    DailyAllocationChecks,
                    #"MW Fortnight Hours Check",
                    #"MW FTE Profile Alignment Check",
                    CategoryChecks,
                    FacilityChecks
                }
            )
in
    Result;
