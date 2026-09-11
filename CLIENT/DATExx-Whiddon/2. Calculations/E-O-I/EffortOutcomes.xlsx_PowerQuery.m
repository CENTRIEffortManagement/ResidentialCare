// Power Query from: EffortOutcomes.xlsx
// Pathname: CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/EffortOutcomes.xlsx
// Extracted: 2026-05-21T00:47:28.279Z
// Source: Existing M version approved for this filepath amendment; no workbook extraction or sync performed.

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve this client-level workbook through the standard CentriSyncPaths mapping.
// Inputs: Canonical FilePathUrl input with one FilePath value; a single Column1 named cell is also supported.
// Output: Existing Variable Name / Value interface, plus client and units roots for imports.
shared UnitL1PathTABLE = let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    PathInput = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
    // Retain the single named-cell shape supported by the previous resolver, but reject unrelated columns.
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then Table.SelectColumns(PathInput, {"FilePath"})
        else if Table.ColumnNames(PathInput) = {"Column1"} then Table.RenameColumns(PathInput, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or a single Column1 named cell.",
    SinglePath = if Table.RowCount(PathColumn) = 1 then PathColumn{0}[FilePath]
        else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = if not Value.Is(SinglePath, type text) then error "FilePathUrl must contain a saved workbook path as text."
        else if Text.Trim(SinglePath) = "" then error "FilePathUrl is blank. Save this workbook and recalculate its filename formula."
        else NormalizePath(SinglePath),
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; strip the sheet suffix.
    WorkbookPath = if Text.Contains(RawFilePath, "[") then
        Text.BeforeDelimiter(RawFilePath, "[") & Text.BetweenDelimiters(RawFilePath, "[", "]")
        else RawFilePath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    ValidatedPath = if Comparer.OrdinalIgnoreCase(FileName, "EffortOutcomes.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Save EffortOutcomes.xlsx and recalculate its filename formula.",
    // Fixed public-machine configuration location; no user-specific source paths.
    MappingWorkbook = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    MappingTable = MappingWorkbook{[Item="CentriSyncPaths", Kind="Table"]}[Data],
    MappingTypes = Table.TransformColumnTypes(
        Table.SelectColumns(MappingTable, {"SharepointRootUrl", "SyncedFolderRootPath"}),
        {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizedMappings = Table.TransformColumns(MappingTypes,
        {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    SiteCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot",
        each if [SharepointRootUrl] = null or [SharepointRootUrl] = "" then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    Candidates = Table.AddColumn(Table.Combine({SiteCandidates, DocumentCandidates, LocalCandidates}),
        "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    // Match complete path segments, so similarly named sites or local folders cannot capture this workbook.
    Matches = Table.SelectRows(Candidates, each [RootLength] > 0
        and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and (Comparer.OrdinalIgnoreCase(ValidatedPath, [MatchRoot]) = 0
            or Text.StartsWith(ValidatedPath, [MatchRoot] & "\", Comparer.OrdinalIgnoreCase))),
    SortedMatches = Table.Sort(Matches, {{"RootLength", Order.Descending}}),
    LongestMatches = if Table.IsEmpty(SortedMatches) then error "FilePathUrl did not match any CentriSyncPaths root: " & ValidatedPath
        else Table.SelectRows(SortedMatches, each [RootLength] = SortedMatches{0}[RootLength]),
    // Conflicting mappings at equal specificity must fail instead of depending on spreadsheet row order.
    Destinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch = if List.Count(Destinations) = 1 then LongestMatches{0}
        else error "CentriSyncPaths contains conflicting mappings for this workbook.",
    MappedRoot = BestMatch[SyncedFolderRootPath],
    LocalRoot = if (Text.Length(MappedRoot) >= 3 and Text.Range(MappedRoot, 1, 2) = ":\")
        or Text.StartsWith(MappedRoot, "\\") then MappedRoot
        else error "CentriSyncPaths SyncedFolderRootPath must be an absolute local or UNC path.",
    LocalFullPath = LocalRoot & Text.Range(ValidatedPath, BestMatch[RootLength]),
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    // E-O-I is under the client, outside UNITS. Derive the client root from its exact folder suffix.
    CalculationSuffix = "\2. Calculations\E-O-I",
    ClientRoot = if Text.EndsWith(WorkbookFolder, CalculationSuffix, Comparer.OrdinalIgnoreCase) then
        Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(CalculationSuffix))
        else error "EffortOutcomes.xlsx must be saved under the client's 2. Calculations\E-O-I folder.",
    Segments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(WorkbookFolder, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    // Buffer the small path table because multiple scalar paths consume the same mapping result.
    PathTable = #table(type table [#"Variable Name" = text, Value = nullable text], {
        {"UserName", UserName}, {"Root Path", WorkbookFolder}, {"FilePathUrl", RawFilePath},
        {"Client", Client}, {"Date", Date}, {"Unit", null}, {"FileName", FileName},
        {"Client Path", ClientRoot}, {"Units Path", ClientRoot & "\UNITS"}
    })
in
    Table.Buffer(PathTable);

// Query: Folder !!
// Purpose: Expose the resolved client root without a trailing separator.
shared #"Folder !!" = UnitL1PathTABLE{[#"Variable Name"="Client Path"]}[Value];

// Query: IMPORT Table_EffortAllMatrixAG1_1D !!
// Purpose: Import the existing effort-matrix table from Effort-All under the resolved client root.
shared #"IMPORT Table_EffortAllMatrixAG1_1D !!" = let
    Source = Excel.Workbook(File.Contents(#"Folder !!"&"\2. Calculations\E-O-I\Effort-All.xlsx"), null, true),
    EffortAllMatrixAG1_1D_2_Table = Source{[Item="EffortAllMatrixAG1_1D_2",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_2_Table,{{"Facility", type text}})
in
    #"Changed Type";

shared #"EffortOutcomesAG1-1DayShift" = let
    Source = #"IMPORT Table_EffortAllMatrixAG1_1D !!",
    Apn = Table.AddColumn(Source, "Apn", each if [Demand]=0 then null 
else [Capacity]/[Demand]),
    Apx = Table.AddColumn(Apn, "Apx", each if [Demand]=0 then null 
else [CapacityX]/[Demand]),
    Ai = Table.AddColumn(Apx, "Ai", each if [Demand]=0 then null 
else [Allocation]/[Demand]),
    Epn = Table.AddColumn(Ai, "Epn", each if [CapacityX]=0 then null 
else [Demand]/[Capacity]),
    Epx = Table.AddColumn(Epn, "EpX", each if [Capacity]=0 then null 
else [Demand]/[CapacityX]),
    Ein = Table.AddColumn(Epx, "Ein", each if [Capacity]=0 then null 
else 
[Allocation]/
[Capacity]),
    Ipn = Table.AddColumn(Ein, "Ipn", each if [Allocation]=0 then null 
else [Capacity] / [Allocation]),
    Ipx = Table.AddColumn(Ipn, "Ipx", each if [Allocation]=0 then null 
else [CapacityX]/[Allocation]),
    Iin = Table.AddColumn(Ipx, "Iin", each if [Allocation]=0 then null 
else [Demand]/[Allocation]),
    #"Filtered Rows" = Table.SelectRows(Iin, each true),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Date", type date}})
in
    #"Changed Type";
