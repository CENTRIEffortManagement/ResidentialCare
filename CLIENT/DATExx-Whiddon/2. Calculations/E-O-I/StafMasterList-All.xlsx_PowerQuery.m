// Power Query from: StafMasterList-All.xlsx
// Pathname: CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/StafMasterList-All.xlsx
// Extracted: 2026-05-21T00:47:34.382Z
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
    ValidatedPath = if Comparer.OrdinalIgnoreCase(FileName, "StafMasterList-All.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Save StafMasterList-All.xlsx and recalculate its filename formula.",
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
        else error "StafMasterList-All.xlsx must be saved under the client's 2. Calculations\E-O-I folder.",
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

// Query: Folder
// Purpose: Expose the resolved client root without a trailing separator.
shared Folder = UnitL1PathTABLE{[#"Variable Name"="Client Path"]}[Value];

// Query: UnitsPath
// Purpose: Expose the client's UNITS root for configured facility imports.
shared UnitsPath = UnitL1PathTABLE{[#"Variable Name"="Units Path"]}[Value];

// Query: ResolveFacilityPath
// Purpose: Resolve a configured unit folder beneath the client's UNITS root.
// Inputs: Existing Table3 Facility1/Facility2 values, such as Unit1 or UNITS\Unit1.
shared ResolveFacilityPath = (facility as any) as text => let
    Normalized = if Value.Is(facility, type text) then Text.Trim(Text.Replace(Text.Trim(facility), "/", "\"), "\")
        else error "Table3 facility paths must be nonblank text identifying a unit folder.",
    UnitName = if Text.StartsWith(Normalized, "UNITS\", Comparer.OrdinalIgnoreCase) then Text.Range(Normalized, 6) else Normalized,
    // A single folder name prevents absolute paths and traversal outside the resolved client.
    ValidatedUnit = if UnitName = "" or List.Contains({".", ".."}, UnitName)
        or Text.Contains(UnitName, "\") or Text.Contains(UnitName, ":") then
            error "Table3 Facility1/Facility2 must identify a unit folder, for example Unit1 or UNITS\Unit1."
        else UnitName
in
    UnitsPath & "\" & ValidatedUnit;

// Query: IMPORTRootPath
// Purpose: Expose the resolved client root through the existing legacy interface.
shared IMPORTRootPath = Folder;

// Query: INPUT FilePath
// Purpose: Preserve the single String column using the canonical resolver input.
shared #"INPUT FilePath" = #table(type table [String = text], {{UnitL1PathTABLE{[#"Variable Name"="FilePathUrl"]}[Value]}});

// Query: Client
// Purpose: Read client metadata from the authoritative path table.
shared Client = UnitL1PathTABLE{[#"Variable Name"="Client"]}[Value];

// Query: Date
// Purpose: Read the client-folder metadata from the authoritative path table.
shared Date = UnitL1PathTABLE{[#"Variable Name"="Date"]}[Value];

// Query: Facility
// Purpose: Preserve the legacy metadata interface; this client-level workbook has no single unit.
shared Facility = UnitL1PathTABLE{[#"Variable Name"="Unit"]}[Value];

// Query: FilePath-Role
// Purpose: Expose the current workbook folder; client-level staff aggregation has no role subfolder.
shared #"FilePath-Role" = UnitL1PathTABLE{[#"Variable Name"="Root Path"]}[Value];

// Query: Role
// Purpose: Preserve the legacy metadata interface; this client-level workbook is not role-specific.
shared Role = null;

// Query: FileName
// Purpose: Read the validated workbook filename from the authoritative path table.
shared FileName = UnitL1PathTABLE{[#"Variable Name"="FileName"]}[Value];

// Query: File Path Data
// Purpose: Preserve the existing metadata table using aliases to the standard resolver.
shared #"File Path Data" = let
   

    // Create the table with variable names and their corresponding values
    Source = #table(
        {"Variable Name", "Value"},
        {
            {"Root Path", #"IMPORTRootPath"},
            {"FilePath", #"FilePath-Role"},
            {"Client", Client},
            {"Date", Date},
            {"Facility", Facility},
            {"Role", Role},
            {"FileName", FileName}
            
            
            
            
            
            
        }
    )
in
    Source;

shared Facilities = let
    Source = Excel.CurrentWorkbook(){[Name="Table3"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Facility1", type text}, {"Facility2", type text}})
in
    #"Changed Type";

shared Facility1 = let
    Source = Facilities,
    Facility1 = Source{0}[Facility1]
in
    Facility1;

shared Facility2 = let
    Source = Facilities,
    Facility1 = Source{0}[Facility2]
in
    Facility1;

// Query: Facility1 - StaffList
// Purpose: Import staff from the unit selected by Facility1 beneath the resolved UNITS root.
shared #"Facility1 - StaffList" = let
        Source = Excel.Workbook(File.Contents(ResolveFacilityPath(Facility1) & "\2. Calculations\StaffListMaster.xlsx"), null, true),
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(Table_Masterlist_Table,{{"Role", type text}, {"Name", type text}, {"Resource", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type2", "Facility", each Facility1)
in
    #"Added Custom";

// Query: Facility2 - StaffList
// Purpose: Import staff from the unit selected by Facility2 beneath the resolved UNITS root.
shared #"Facility2 - StaffList" = let
        Source = Excel.Workbook(File.Contents(ResolveFacilityPath(Facility2) & "\2. Calculations\StaffListMaster.xlsx"), null, true),
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(Table_Masterlist_Table,{{"Role", type text}, {"Name", type text}, {"Resource", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type2", "Facility", each Facility2)
in
    #"Added Custom";

shared #"MasterStaffList-All" = let
    Source = Table.Combine({#"Facility1 - StaffList", #"Facility2 - StaffList"})
in
    Source;

shared Table_Masterlist = let
    Source = #"MasterStaffList-All"
in
    Source;
