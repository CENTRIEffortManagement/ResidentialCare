// Power Query from: Capacity.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\Capacity.xlsx
// Extracted: 2026-10-07T06:55:17.023Z

section Section1;

shared #"IMPORT C#TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C#"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RosteredPeriodStatus"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityCapped", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C##TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C###TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&Role2&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

// Query: LINKRole
// Purpose: Land Role Analysis items unchanged for enabled-role preparation.
shared LINKRole = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"]),
    Items = Source{[Id="bad3beb8-7064-4e11-9edd-d62dac5d702d"]}[Items]
in
    Items;

// Query: Role1
// Purpose: Retain the AIN compatibility imports without using a SharePoint row position.
// Notes: Business outputs enumerate every enabled role independently of these three legacy slots.
shared Role1 = "AIN";

// Query: Role2
// Purpose: Retain the RN compatibility imports without using a SharePoint row position.
shared Role2 = "RN";

// Query: Role3
// Purpose: Retain the current EN nursing-lane compatibility imports without using a SharePoint row position.
shared Role3 = "EN";

shared Folder = let
    Source = Unit1Path
in
    Source;

shared Facility = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT Role1C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role1C##TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role1C###TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder &"\" & #"Role1"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

// Query: Capacity
// Purpose: Publish final C### capacity for every enabled role with its unique Settings date.
// Notes: Retain the complete B publication, including its existing additional fields.
shared Capacity = let
    Source = #"Capacity Final Inputs Prepare",
    #"Expanded PermutationDimensions" = #"Capacity Calendar Match Prepare"(Source, true),
    #"Sorted Rows" = Table.Sort(#"Expanded PermutationDimensions",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows"
   ;

// Query: AvailabilityDeveloped
// Purpose: Publish the existing availability stages with reporting columns only.
// Inputs: Original availability from A.1, C# from A.2, and C##/C### from B for the configured roles.
// Output: Role, Resource, Period, Availability, AvailabilityType, Facility, ResAvailability and ResMaxAvail.
// Notes: Calculation and diagnostic metadata must not become row keys in downstream reporting pivots.
shared AvailabilityDeveloped = let
    Source = #"Capacity Reporting Inputs Prepare",
    #"Added Custom" = Table.AddColumn(Source, "Facility", each Facility),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",0,null,Replacer.ReplaceValue,{"Availability"}),
    #"Removed Columns" = Table.RemoveColumns(#"Replaced Value",{"OriginalAvailability"}),
    // Preserve the extracted removal above, then restrict the reporting interface to its business columns.
    ReportingColumns = Table.SelectColumns(#"Removed Columns", {"Role", "Resource", "Period", "Availability", "AvailabilityType", "Facility", "ResAvailability", "ResMaxAvail"})
in
    ReportingColumns;

shared AvailabilityDevelopedMATRIX = let
    Source = AvailabilityDeveloped,
    #"Removed Columns" = Table.RemoveColumns(Source,{"ResAvailability", "ResMaxAvail"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[AvailabilityType]), "AvailabilityType", "Availability", List.Sum)
in
    #"Pivoted Column";

shared PermutationDimensions = let
    Source = Excel.Workbook(File.Contents(Folder&"\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared #"IMPORT Role3C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role3C##TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role3C###TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE3" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        PathInput = if Table.HasColumns(Source, {"FilePath"}) then Table.SelectColumns(Source, {"FilePath"})
            else if Table.ColumnNames(Source) = {"Column1"} then Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must contain a FilePath column or be a single named cell.",
        ValidatedRows = if Table.RowCount(PathInput) = 1 then PathInput
            else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedRows)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if not Value.Is(RawFilePathValue, type text) then
        error "FilePathUrl[FilePath] must contain the saved workbook path as text."
        else if Text.Trim(RawFilePathValue) = "" then
        error "FilePathUrl[FilePath] is blank. Save this workbook and recalculate its CELL filename formula."
        else Text.Trim(RawFilePathValue),
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(RawFilePath),
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
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: Capacity Enabled Roles Prepare
// Purpose: Read every enabled role label without relying on list order or three positional slots.
shared #"Capacity Enabled Roles Prepare" = let
    Selected = Table.SelectColumns(LINKRole, {"Roles", "Effort Management Analysis"}),
    Enabled = Table.SelectRows(Selected, each [Effort Management Analysis] = true),
    Named = Table.RenameColumns(Table.SelectColumns(Enabled, {"Roles"}), {{"Roles", "RoleName"}}),
    Normalized = Table.TransformColumns(Named, {{"RoleName", each
        if _ = null then null else Text.Trim(Text.From(_)), type nullable text}}),
    InvalidNames = Table.SelectRows(Normalized, each [RoleName] = null or [RoleName] = ""
        or List.Contains({".", ".."}, [RoleName]) or Text.Contains([RoleName], "\") or Text.Contains([RoleName], "/")),
    WithKey = Table.AddColumn(Normalized, "RoleKey", each
        if [RoleName] = null then null else Text.Upper([RoleName]), type nullable text),
    Counts = Table.Group(WithKey, {"RoleKey"}, {{"Rows", each Table.RowCount(_), Int64.Type}}),
    Validated = if Table.IsEmpty(Normalized) or not Table.IsEmpty(InvalidNames)
            or not Table.IsEmpty(Table.SelectRows(Counts, each [Rows] <> 1)) then
            error "Capacity requires unique enabled role labels, each identifying one role folder."
        else Normalized,
    Sorted = Table.Sort(Validated, {{"RoleName", Order.Ascending}})
in
    Table.Buffer(Sorted);

// Query: IMPORT Role A1
// Purpose: Land one role's A.1 workbook navigation unchanged.
shared #"IMPORT Role A1" = (RoleName as text) as table =>
let
    SourceBinary = Binary.Buffer(File.Contents(Folder & "\" & RoleName & "\CapacityDistrib(A.1)-shifts.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true)
in
    Navigation;

// Query: IMPORT Role A2
// Purpose: Land one role's A.2 workbook navigation unchanged.
shared #"IMPORT Role A2" = (RoleName as text) as table =>
let
    SourceBinary = Binary.Buffer(File.Contents(Folder & "\" & RoleName & "\CapacityDistrib(A.2)-shifts.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true)
in
    Navigation;

// Query: IMPORT Role B
// Purpose: Land one role's B workbook navigation unchanged for both B publications within an evaluation.
shared #"IMPORT Role B" = (RoleName as text) as table =>
let
    SourceBinary = Binary.Buffer(File.Contents(Folder & "\" & RoleName & "\CapacityDistrib(B)-shifts.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true)
in
    Table.Buffer(Navigation);

// Query: EXTRACT Role Capacity Table
// Purpose: Land the exact saved table unchanged; missing or ambiguous bindings remain errors.
shared #"EXTRACT Role Capacity Table" = (Navigation as table, TableName as text) as table =>
    Navigation{[Item = TableName, Kind = "Table"]}[Data];

// Query: Capacity Stage Prepare
// Purpose: Preserve existing capacity units and stage values while validating role-folder ownership and cell keys.
// Notes: Reporting prunes diagnostic width before buffering. The final C### path retains every B output column.
shared #"Capacity Stage Prepare" = (Raw as table, RoleName as text, StageName as text, Reporting as logical) as table =>
let
    ValueColumn = if StageName = "OriginalAvailability" then "Availability"
        else if StageName = "C#" then "AvailabilityCapped"
        else if StageName = "C##" then "C##"
        else if StageName = "C###" then "C###"
        else error "Unknown capacity stage.",
    RequiredColumns = {"Role", "Resource", "Period", ValueColumn}
        & (if StageName = "C###" then {"ResAvailability", "ResMaxAvail"} else {}),
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Raw)),
    CheckedColumns = if List.IsEmpty(MissingColumns) then Raw
        else error Error.Record("Capacity stage input", RoleName & " " & StageName & " is missing required columns.", [MissingColumns = MissingColumns]),
    ReportingInputs = List.Distinct(RequiredColumns & {"ResAvailability", "ResMaxAvail", "OriginalAvailability"}),
    ScopedColumns = if Reporting then Table.SelectColumns(CheckedColumns, ReportingInputs, MissingField.UseNull) else CheckedColumns,
    Types = {{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {ValueColumn, Int64.Type}}
        & (if StageName = "C###" then {{"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}} else {}),
    Typed = Table.TransformColumnTypes(ScopedColumns, Types),
    Snapshot = Table.Buffer(Typed),
    WrongRoles = Table.SelectRows(Snapshot, each [Role] = null or [Role] <> RoleName),
    ValidRole = if Table.IsEmpty(WrongRoles) then Snapshot
        else error Error.Record("Capacity role input", "Saved " & RoleName & " " & StageName
            & " rows belong to another or a blank role. Refresh and save that role's own workbook before consolidating.",
            [ExpectedRole = RoleName, Stage = StageName]),
    MissingKeys = Table.SelectRows(ValidRole, each [Resource] = null or [Period] = null),
    KeyCounts = Table.Group(Table.SelectColumns(ValidRole, {"Resource", "Period"}), {"Resource", "Period"},
        {{"Rows", each Table.RowCount(_), Int64.Type}}),
    DuplicateKeys = Table.SelectRows(KeyCounts, each [Rows] <> 1),
    Validated = if not Table.IsEmpty(MissingKeys) or not Table.IsEmpty(DuplicateKeys) then
            error Error.Record("Capacity stage keys", RoleName & " " & StageName & " requires one row per Resource and Period.",
                [MissingKeys = Table.RowCount(MissingKeys), DuplicateKeys = Table.RowCount(DuplicateKeys)])
        else ValidRole,
    WithStage = Table.AddColumn(Validated, "AvailabilityType", each StageName, type text),
    NamedAvailability = if ValueColumn = "Availability" then WithStage
        else Table.RenameColumns(WithStage, {{ValueColumn, "Availability"}}),
    ReportingColumns = {"Role", "Resource", "Period", "Availability", "AvailabilityType", "ResAvailability", "ResMaxAvail", "OriginalAvailability"},
    Result = if Reporting then Table.SelectColumns(NamedAvailability, ReportingColumns) else NamedAvailability
in
    Result;

// Query: Capacity Role Inputs Prepare
// Purpose: Prepare one enabled role's final B input or its four reporting stages.
// Notes: Main Capacity reads B only. Reporting binds B once and reuses its navigation for C## and C###.
shared #"Capacity Role Inputs Prepare" = (RoleName as text, AllStages as logical) as table =>
let
    BWorkbook = #"IMPORT Role B"(RoleName),
    BFinalTable = #"EXTRACT Role Capacity Table"(BWorkbook, "C___TABLE_B"),
    Final = #"Capacity Stage Prepare"(BFinalTable, RoleName, "C###", AllStages),
    Result = if not AllStages then Final else
        let
            A1Workbook = #"IMPORT Role A1"(RoleName),
            A2Workbook = #"IMPORT Role A2"(RoleName),
            OriginalTable = #"EXTRACT Role Capacity Table"(A1Workbook, "ResPeriodAvailabilityTABLE"),
            CappedTable = #"EXTRACT Role Capacity Table"(A2Workbook, "ResPeriodAvailabilityCapped_C__TABLE"),
            BIntermediateTable = #"EXTRACT Role Capacity Table"(BWorkbook, "C__TABLE"),
            Original = #"Capacity Stage Prepare"(OriginalTable, RoleName, "OriginalAvailability", true),
            Capped = #"Capacity Stage Prepare"(CappedTable, RoleName, "C#", true),
            Intermediate = #"Capacity Stage Prepare"(BIntermediateTable, RoleName, "C##", true),
            Stages = Table.Combine({Original, Capped, Intermediate, Final})
        in
            Stages
in
    Result;

// Query: Capacity Calendar Prepare
// Purpose: Require unique Settings dates for enabled role/period keys before joining capacity.
shared #"Capacity Calendar Prepare" = let
    Roles = List.Buffer(#"Capacity Enabled Roles Prepare"[RoleName]),
    Selected = Table.SelectColumns(PermutationDimensions, {"Period", "RolesList", "Date"}),
    Scoped = Table.SelectRows(Selected, each List.Contains(Roles, [RolesList])),
    Invalid = Table.SelectRows(Scoped, each [Period] = null or not Value.Is([Date], type date)),
    Counts = Table.Group(Scoped, {"Period", "RolesList"}, {{"Rows", each Table.RowCount(_), Int64.Type}}),
    Duplicates = Table.SelectRows(Counts, each [Rows] <> 1),
    Validated = if not Table.IsEmpty(Invalid) or not Table.IsEmpty(Duplicates) then
            error "Capacity requires one nonblank Settings date per enabled role and Period."
        else Scoped
in
    Table.Buffer(Validated);

// Query: Capacity Calendar Match Prepare
// Purpose: Match every capacity stage cell to one Settings period without multiplying or silently losing rows.
shared #"Capacity Calendar Match Prepare" = (Rows as table, IncludeDate as logical) as table =>
let
    Joined = Table.NestedJoin(Rows, {"Period", "Role"}, #"Capacity Calendar Prepare",
        {"Period", "RolesList"}, "CapacityDate", JoinKind.LeftOuter),
    // Validation and publication consume the same projected rows; do not enumerate role-workbook functions twice.
    Snapshot = Table.Buffer(Joined),
    Invalid = Table.SelectRows(Snapshot, each Table.RowCount([CapacityDate]) <> 1),
    Validated = if Table.IsEmpty(Invalid) then Snapshot
        else error "Capacity contains role/Period cells missing from the saved Settings calendar.",
    Result = if IncludeDate then Table.ExpandTableColumn(Validated, "CapacityDate", {"Date"}, {"Date"})
        else Table.RemoveColumns(Validated, {"CapacityDate"})
in
    Result;

// Query: Capacity Final Inputs Prepare
// Purpose: Collect final C### rows from every enabled role, retaining B's existing schema and values.
shared #"Capacity Final Inputs Prepare" = let
    Roles = #"Capacity Enabled Roles Prepare"[RoleName],
    Tables = List.Transform(Roles, each #"Capacity Role Inputs Prepare"(_, false)),
    Combined = Table.Combine(Tables)
in
    Combined;

// Query: Capacity Reporting Inputs Prepare
// Purpose: Collect the four reporting stages for every enabled role on the existing reporting grain.
shared #"Capacity Reporting Inputs Prepare" = let
    Roles = #"Capacity Enabled Roles Prepare"[RoleName],
    Tables = List.Transform(Roles, each #"Capacity Role Inputs Prepare"(_, true)),
    Combined = Table.Combine(Tables),
    Checked = #"Capacity Calendar Match Prepare"(Combined, false)
in
    Checked;
