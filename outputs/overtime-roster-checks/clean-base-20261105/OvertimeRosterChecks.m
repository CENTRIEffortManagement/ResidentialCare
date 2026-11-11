section Section1;

shared #"IMPORT OvertimeRosterChecks CentriSyncPaths" = let
    Navigation = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    Mapping = Navigation{[Item = "CentriSyncPaths", Kind = "Table"]}[Data]
in
    Mapping;

shared #"INPUT OvertimeRosterChecks FilePathUrl" = Excel.CurrentWorkbook(){[Name = "FilePathUrl"]}[Content];

shared OvertimeRosterChecks_Path = let
    NormalizePath = (value as any) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(Text.From(value)), "/", "\"), "\"),
    PathInput = #"INPUT OvertimeRosterChecks FilePathUrl",
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then PathInput[FilePath]
        else if Table.ColumnNames(PathInput) = {"Column1"} then PathInput[Column1]
        else error "FilePathUrl must have a FilePath column or be one named cell.",
    RawPath = if List.Count(PathColumn) = 1 then NormalizePath(PathColumn{0}) else error "FilePathUrl requires exactly one value.",
    NonBlankPath = if RawPath = null or RawPath = "" then error "Save and recalculate OvertimeRosterChecks.xlsx before resolving FilePathUrl." else RawPath,
    WorkbookPath = if Text.Contains(NonBlankPath, "[") then Text.BeforeDelimiter(NonBlankPath, "[") & Text.BetweenDelimiters(NonBlankPath, "[", "]") else NonBlankPath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NamedWorkbookPath = if Comparer.OrdinalIgnoreCase(FileName, "OvertimeRosterChecks.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
    MappingColumns = Table.SelectColumns(#"IMPORT OvertimeRosterChecks CentriSyncPaths", {"SharepointRootUrl", "SyncedFolderRootPath"}),
    NormalizedMappings = Table.TransformColumns(MappingColumns, {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    BufferedMappings = Table.Buffer(NormalizedMappings),
    SharePointRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    CandidateRoots = Table.Combine({SharePointRoots, DocumentRoots, LocalRoots}),
    RootLengths = Table.AddColumn(CandidateRoots, "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRoots = Table.SelectRows(RootLengths, each [RootLength] > 0 and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(NamedWorkbookPath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(NamedWorkbookPath) = [RootLength] or Text.Range(NamedWorkbookPath, [RootLength], 1) = "\")),
    OrderedRoots = Table.Sort(MatchingRoots, {{"RootLength", Order.Descending}}),
    LongestRoot = if Table.IsEmpty(OrderedRoots) then error "FilePathUrl has no CentriSyncPaths mapping." else OrderedRoots{0},
    EqualRoots = Table.SelectRows(OrderedRoots, each [RootLength] = LongestRoot[RootLength]),
    Destinations = List.Distinct(EqualRoots[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    SelectedRoot = if List.Count(Destinations) = 1 then LongestRoot else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    RelativePath = Text.TrimStart(Text.Range(NamedWorkbookPath, SelectedRoot[RootLength]), "\"),
    LocalWorkbookPath = SelectedRoot[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NestedSuffix = "\2. Calculations\OvertimeRosterChecks",
    ClientPath = if Text.EndsWith(WorkbookFolder, NestedSuffix, Comparer.OrdinalIgnoreCase) then Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(NestedSuffix))
        else error "OvertimeRosterChecks.xlsx must be in the organisation's 2. Calculations/OvertimeRosterChecks folder.",
    ClientSegments = Text.Split(ClientPath, "\"),
    OrganisationPath = if List.PositionOf(ClientSegments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase) < 0 then ClientPath
        else error "OvertimeRosterChecks.xlsx is organisation-level and cannot be located beneath UNITS.",
    Result = [ClientPath = OrganisationPath, WorkbookPath = LocalWorkbookPath]
in
    Result;

shared #"IMPORT Published Roster" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\1. Input\Allocation\Published Roster.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Roster Data", Kind = "Sheet"]}[Data]
in
    SourceObject;

shared #"IMPORT Workers" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\1. Input\Worker's Report.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Sheet1", Kind = "Sheet"]}[Data]
in
    SourceObject;

shared #"IMPORT Settings Data" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\2. Calculations\Settings Data.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true)
in
    Navigation;

shared #"IMPORT OvertimeRosterChecks Whiddon Lists" = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation = "2.0", ViewMode = "All"]);

shared #"LINK OvertimeRosterChecks Roles" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items];

shared #"LINK OvertimeRosterChecks Facilities" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items];

shared #"LINK OvertimeRosterThreshold" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Title = "OvertimeRosterThreshold"]}[Items];