// Power Query from: Settings Data.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\Settings Data.xlsx
// Extracted: 2026-09-24T08:43:06.527Z

section Section1;

// Query: IMPORT CentriSyncPaths
// Purpose: Read the fixed public-machine path mapping for workbook-relative imports.
shared #"IMPORT CentriSyncPaths" =
let
    Navigation = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    Mapping = Navigation{[Item="CentriSyncPaths", Kind="Table"]}[Data]
in
    Table.Buffer(Mapping);

// Query: UnitL1PathTABLE
// Purpose: Resolve this Settings workbook path and its unit folder through CentriSyncPaths.
// Inputs: Current-workbook FilePathUrl and IMPORT CentriSyncPaths.
// Output: Variable Name / Value rows including Root Path, Unit Root and Unit.
shared UnitL1PathTABLE =
let
    PathInput = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then Table.SelectColumns(PathInput, {"FilePath"})
        else if Table.ColumnNames(PathInput) = {"Column1"} then Table.RenameColumns(PathInput, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain one FilePath column or one named cell.",
    RawValue = if Table.RowCount(PathColumn) = 1 then PathColumn{0}[FilePath]
        else error "FilePathUrl must contain exactly one path.",
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    RawPath = NormalizePath(if RawValue = null then null else Text.From(RawValue)),
    NonBlankPath = if RawPath = null or RawPath = "" then error "FilePathUrl is blank. Save and recalculate Settings Data.xlsx." else RawPath,
    // CELL("filename") supplies folder\[workbook.xlsx]sheet; the sheet suffix is not part of the path.
    WorkbookPath = if Text.Contains(NonBlankPath, "[") then
        Text.BeforeDelimiter(NonBlankPath, "[") & Text.BetweenDelimiters(NonBlankPath, "[", "]")
        else NonBlankPath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "Settings Data.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl must identify Settings Data.xlsx.",
    Mapping = Table.TransformColumns(
        Table.SelectColumns(#"IMPORT CentriSyncPaths", {"SharepointRootUrl", "SyncedFolderRootPath"}),
        {{"SharepointRootUrl", each NormalizePath(_), type nullable text},
         {"SyncedFolderRootPath", each NormalizePath(_), type nullable text}}),
    SharePoint = Table.AddColumn(Mapping, "MatchRoot", each [SharepointRootUrl], type nullable text),
    Documents = Table.AddColumn(Mapping, "MatchRoot", each if [SharepointRootUrl] = null then null
        else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    Local = Table.AddColumn(Mapping, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    Candidates = Table.AddColumn(Table.Combine({SharePoint, Documents, Local}), "RootLength",
        each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    Matching = Table.SelectRows(Candidates, each [MatchRoot] <> null and [MatchRoot] <> ""
        and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(FilePath) = [RootLength] or Text.Range(FilePath, [RootLength], 1) = "\")),
    Ordered = Table.Sort(Matching, {{"RootLength", Order.Descending}}),
    Longest = if Table.IsEmpty(Ordered) then error "FilePathUrl has no CentriSyncPaths mapping." else Ordered{0},
    SameLength = Table.SelectRows(Ordered, each [RootLength] = Longest[RootLength]),
    Best = if List.Count(List.Distinct(SameLength[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase)) = 1 then Longest
        else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    Relative = Text.TrimStart(Text.Range(FilePath, Best[RootLength]), "\"),
    LocalWorkbook = if Relative = "" then Best[SyncedFolderRootPath]
        else Best[SyncedFolderRootPath] & "\" & Relative,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbook, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    CalcIndex = List.PositionOf(Segments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidFolder = UnitsIndex >= 0 and CalcIndex = UnitsIndex + 2 and List.Count(Segments) = CalcIndex + 1,
    RootPath = if ValidFolder then WorkbookFolder else error "Settings Data.xlsx must be under UNITS/<unit>/2. Calculations.",
    UnitRoot = Text.BeforeDelimiter(RootPath, "\", {0, RelativePosition.FromEnd}),
    Unit = Segments{UnitsIndex + 1},
    Output = #table({"Variable Name", "Value"}, {{"Root Path", RootPath}, {"Unit Root", UnitRoot},
        {"FilePathUrl", FilePath}, {"Unit", Unit}, {"FileName", InputFileName}})
in
    Table.Buffer(Output);

// Query: Unit1Path
// Purpose: Supply the resolved unit root to the AllocationExtracted import.
shared Unit1Path = UnitL1PathTABLE{[Variable Name="Unit Root"]}[Value];

// Query: Unit
// Purpose: Expose the folder-derived unit name used by the facility filters.
shared Unit = UnitL1PathTABLE{[Variable Name="Unit"]}[Value];

// Query: IMPORT Whiddon Lists
// Purpose: Read the Whiddon SharePoint navigation for facility and analysis references.
shared #"IMPORT Whiddon Lists" =
    SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation=null, ApiVersion=15]);

// Query: LINK Facilities
// Purpose: Map the imported allocation Location to the canonical Facility-Abbrev.
shared #"LINK Facilities" =
let
    FacilityItems = #"IMPORT Whiddon Lists"{[Id="1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items],
    SelectedColumns = Table.SelectColumns(FacilityItems, {"Title", "field_1"}),
    NamedColumns = Table.RenameColumns(SelectedColumns, {{"field_1", "Facility-Abbrev"}})
in
    NamedColumns;

// Query: LINK Facility Analysis
// Purpose: Select the same Facility Analysis fields used by AllocationExtracted.
shared #"LINK Facility Analysis" =
let
    AnalysisItems = #"IMPORT Whiddon Lists"{[Id="17ed8e30-5707-4af2-8e21-d68eb8184287"]}[Items],
    SelectedColumns = Table.SelectColumns(AnalysisItems, {"Title", "EffortManagementAnalysis", "LeaveBalanceAnalysis"})
in
    SelectedColumns;

// Query: FacilityAnalysisTABLE
// Purpose: Identify whether the folder-derived Unit is enabled for Effort Management Analysis.
shared FacilityAnalysisTABLE =
let
    Source = #"LINK Facility Analysis",
    UnitRows = Table.SelectRows(Source, each [Title] <> null and Comparer.OrdinalIgnoreCase(Text.Trim([Title]), Unit) = 0),
    ValidatedUnit = if Table.RowCount(UnitRows) = 1 then UnitRows else error "Facility Analysis must identify exactly one row for the current Unit.",
    EnabledUnit = Table.SelectRows(ValidatedUnit, each [EffortManagementAnalysis] = true),
    Result = Table.RemoveColumns(EnabledUnit, {"LeaveBalanceAnalysis"})
in
    Result;

// Query: Allocation Unit Rows
// Purpose: Limit Settings dates and roles to allocation locations mapped to the active Unit.
shared #"Allocation Unit Rows" =
let
    Source = #"IMPORT AllocationExtracted",
    WithoutPriorAbbrev = Table.RemoveColumns(Source, {"Facility-Abbrev"}, MissingField.Ignore),
    JoinedFacilities = Table.NestedJoin(WithoutPriorAbbrev, {"Location"}, #"LINK Facilities", {"Title"}, "FacilityLookup", JoinKind.LeftOuter),
    ExpandedFacilities = Table.ExpandTableColumn(JoinedFacilities, "FacilityLookup", {"Facility-Abbrev"}, {"Facility-Abbrev"}),
    AnalysisEnabled = not Table.IsEmpty(FacilityAnalysisTABLE),
    FilteredUnit = Table.SelectRows(ExpandedFacilities, each [#"Facility-Abbrev"] = Unit and AnalysisEnabled)
in
    FilteredUnit;

shared #"Dates Listed" = let
    Source = DateRoleShiftAllocation,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Allocation"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Date", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared #"Min Date" = let
    Source = #"Dates Listed",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] >= DateFrom)),
    Custom1 = Table.Min(#"Filtered Rows","Date"),
    Date = Custom1[Date]
in
    Date;

shared DateList = let
    Source = #"Dates Listed",
    #"Changed Type4" = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Custom1 = #"Changed Type4"[Date],
    Custom2 = List.Max(#"Changed Type4" [Date]),
    #"Converted to Table" = #table(1, {{Custom2}}),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Max"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Min", each #"Min Date"),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom1",{{"Max", type date}, {"Min", type date}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"Max", Int64.Type}, {"Min", Int64.Type}}),
    #"Added Custom3" = Table.AddColumn(#"Changed Type2", "Custom", each {[Min]..[Max]}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom3",{{"Max", type date}, {"Min", type date}}),
    #"Expanded Custom" = Table.ExpandListColumn(#"Changed Type1", "Custom"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Custom",{"Max", "Min"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Custom", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"Custom", "Date"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared PermutationDimensions = let
    Source = #"DateList",
    #"Added SHIFTS" = Table.AddColumn(Source, "Shifts", each Shifts),
    #"Expanded Shifts" = Table.ExpandListColumn(#"Added SHIFTS", "Shifts"),
    #"Added Index" = Table.AddIndexColumn(#"Expanded Shifts", "Period", 1, 1, Int64.Type),
    #"ADD ROLES" = Table.AddColumn(#"Added Index", "RolesList", each Roles),
    #"Expanded RolesList" = Table.ExpandListColumn(#"ADD ROLES", "RolesList"),
    #"Sorted Rows" = Table.Sort(#"Expanded RolesList",{{"Date", Order.Ascending}, {"Period", Order.Ascending}, {"RolesList", Order.Ascending}})
in
    #"Sorted Rows";

shared Roles = let
    Source = #"Allocation Unit Rows",
    Role = Source[Role],
    #"Removed Duplicates" = List.Distinct(Role)
in
    #"Removed Duplicates";

shared DateNameRoleShiftAllocation = let
    Source = #"Allocation Unit Rows",
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Custom", each (if [End]<[Start]
then (1-[Start])+[End]
else [End]-[Start])*24),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each ([End] <> null)),
    #"Added Custom2" = Table.AddColumn(#"Filtered Rows", "TimeWorked", each if [Custom] >6 then [Custom] - .5
else [Custom]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom2", "Shift", each if [Start] < 0.604 then "AM" else if [Start] > 0.89 then "NIGHT" else "PM"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom1",{ "Start", "End", "Custom"})
in
    #"Removed Columns1";

shared DateRoleShiftAllocation = let
    Source = DateNameRoleShiftAllocation,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Role", "Shift"}, {{"Allocation", each List.Sum([TimeWorked]), type number}})
in
    #"Grouped Rows";

// Query: IMPORT AllocationExtracted
// Purpose: Read the current unit's already-filtered allocation output.
shared #"IMPORT AllocationExtracted" = let
    Source = Excel.Workbook(File.Contents(Unit1Path & "\1. Input\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Date", type date}, {"Start", type time}, {"End", type time}, {"Break", Int64.Type}, {"Hours", type number}, {"Name", type text}, {"Code", Int64.Type}, {"Role", type text}})
in
    #"Changed Type";

shared Shifts = let
    Source = Excel.CurrentWorkbook(){[Name="Shifts"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}}),
    #"Added Index1" = Table.AddIndexColumn(#"Changed Type", "Index.1", 1, 1, Int64.Type),
    Shifts1 = #"Added Index1"[Shifts]
in
    Shifts1;

shared DateFrom = let
    Source = Excel.CurrentWorkbook(){[Name="DateFrom"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateFrom", type date}}),
    DateFrom1 = #"Changed Type"{0}[DateFrom]
in
    DateFrom1;