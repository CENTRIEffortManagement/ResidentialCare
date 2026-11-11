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
            error "FilePathUrl[FilePath] is blank. Save and recalculate OvertimeRosterChecks.xlsx."
        else NormalizePath(RawFilePath),
    InputWorkbookFolder = Text.BeforeDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputWorkbookSuffix = Text.AfterDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputFileName = if Text.StartsWith(InputWorkbookSuffix, "[") then Text.BetweenDelimiters(InputWorkbookSuffix, "[", "]") else InputWorkbookSuffix,
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "OvertimeRosterChecks.xlsx") = 0 then InputWorkbookFolder & "\" & InputFileName
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
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
    CalculationIndex = List.PositionOf(Segments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidUnitFolder = UnitsIndex >= 0 and CalculationIndex = UnitsIndex + 2 and CalculationIndex = List.Count(Segments) - 1,
    UnitRoot = if ValidUnitFolder then Text.BeforeDelimiter(WorkbookFolder, "\", {0, RelativePosition.FromEnd})
        else error "OvertimeRosterChecks.xlsx must be under UNITS/<unit>/2. Calculations.",
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

// Query: IMPORT Published Roster
// Purpose: Land one configured unit's published Roster Data sheet unchanged.
shared #"IMPORT Published Roster" = (unit as text) as table =>
let
    UnitsRoot = Text.BeforeDelimiter(Unit1Path, "\", {0, RelativePosition.FromEnd}),
    SourcePath = UnitsRoot & "\" & unit & "\1. Input\Allocation\Published Roster.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Roster Data", Kind = "Sheet"]}[Data]
in
    SourceObject;

// Query: IMPORT Workers
// Purpose: Land one configured unit's worker-report sheet unchanged.
shared #"IMPORT Workers" = (unit as text) as table =>
let
    UnitsRoot = Text.BeforeDelimiter(Unit1Path, "\", {0, RelativePosition.FromEnd}),
    SourcePath = UnitsRoot & "\" & unit & "\1. Input\Worker's Report.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Sheet1", Kind = "Sheet"]}[Data]
in
    SourceObject;

// Query: IMPORT Settings Data
// Purpose: Land one configured unit's Settings Data workbook navigation unchanged.
shared #"IMPORT Settings Data" = (unit as text) as table =>
let
    UnitsRoot = Text.BeforeDelimiter(Unit1Path, "\", {0, RelativePosition.FromEnd}),
    SourcePath = UnitsRoot & "\" & unit & "\2. Calculations\Settings Data.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true)
in
    Navigation;

// Query: IMPORT OvertimeRosterChecks Whiddon Lists
// Purpose: Land the Whiddon SharePoint navigation unchanged.
shared #"IMPORT OvertimeRosterChecks Whiddon Lists" = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation = "2.0", ViewMode = "All"]);
// Query: LINK OvertimeRosterChecks Roles
// Purpose: Land the existing role mapping unchanged.
shared #"LINK OvertimeRosterChecks Roles" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items];
// Query: LINK OvertimeRosterChecks Facilities
// Purpose: Land the existing facility mapping unchanged.
shared #"LINK OvertimeRosterChecks Facilities" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items];
// Query: LINK OvertimeRosterThreshold
// Purpose: Land the named threshold list unchanged.
shared #"LINK OvertimeRosterThreshold" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Title = "OvertimeRosterThreshold"]}[Items];
