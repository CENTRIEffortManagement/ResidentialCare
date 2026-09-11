// Power Query from: Tableau Connection.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\Tableau Connection.xlsx
// Extracted: 2026-09-11T06:59:22.093Z

section Section1;

shared DateLevelPathRecord = // Version 25.03 CentriSyncPaths Date-level ResidentialCare
let
    FilePathUrl =
    let
        Source = try Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content] otherwise Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        FirstColumn = Table.ColumnNames(Source){0},
        RenamedColumns = if FirstColumn = "FilePath" then Source else Table.RenameColumns(Source, {{FirstColumn, "FilePath"}}, MissingField.Ignore),
        ReplacedValue = Table.TransformColumns(RenamedColumns, {{"FilePath", each Text.Replace(Text.From(_), "/", "\"), type text}}),
        BufferedTable = Table.Buffer(ReplacedValue)
    in
        BufferedTable,

    RawFilePath = FilePathUrl{0}[FilePath],
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    IsNonBlank = (value as nullable text) as logical => if value = null then false else Text.Trim(value) <> "",
    NormalizedFilePath = NormalizePath(RawFilePath),
    FilePath = if IsNonBlank(NormalizedFilePath) then NormalizedFilePath else error "FilePathUrl is blank.",
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_Selected = Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}, MissingField.Error),
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(CentriSyncPaths_Selected, {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
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
    MatchCandidates = Table.Distinct(
        Table.SelectRows(
            Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
            each IsNonBlank([MatchRoot]) and IsNonBlank([SyncedFolderRootPath])
        ),
        {"MatchRoot", "SyncedFolderRootPath"}
    ),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(
        MatchCandidates_WithLength,
        each Text.Upper(FilePath) = Text.Upper([MatchRoot]) or Text.StartsWith(FilePath, [MatchRoot] & "\", Comparer.OrdinalIgnoreCase)
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    MaxMatchLength = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0}[MatchRootLength] else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    BestRows = Table.SelectRows(SortedMatches, each [MatchRootLength] = MaxMatchLength),
    DistinctBestRows = Table.Distinct(Table.SelectColumns(BestRows, {"MatchRoot", "SyncedFolderRootPath"})),
    BestMatch = if Table.RowCount(DistinctBestRows) = 1 then DistinctBestRows{0} else error "FilePathUrl matched multiple CentriSyncPaths roots with the same length: " & FilePath,
    RelativePath = Text.Range(FilePath, Text.Length(BestMatch[MatchRoot])),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    CurrentWorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(CurrentWorkbookFolder, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    DateRootSegments = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then List.FirstN(Segments, ResidentialCareIndex + 3) else error "Current workbook path does not contain ResidentialCare/CLIENT/DATExx: " & LocalFullPath,
    DateRootPath = Text.Combine(DateRootSegments, "\"),
    UnitsRootPath = DateRootPath & "\UNITS",
    OrgCalculationsPath = DateRootPath & "\2. Calculations",
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(CurrentWorkbookFolder, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else error "Current workbook path does not include a ResidentialCare client segment: " & LocalFullPath,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else error "Current workbook path does not include a ResidentialCare date segment: " & LocalFullPath,
    CurrentWorkbookFileName = try Text.BetweenDelimiters(LocalFullPath, "[", "]") otherwise Text.AfterDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    OUTPUT = [
        Username = UserName,
        CurrentWorkbookFolder = CurrentWorkbookFolder,
        DateRootPath = DateRootPath,
        UnitsRootPath = UnitsRootPath,
        OrgCalculationsPath = OrgCalculationsPath,
        FilePathUrl = RawFilePath,
        Client = Client,
        Date = Date,
        CurrentWorkbookFileName = CurrentWorkbookFileName
    ]
in
    OUTPUT;

shared DateRootPath = DateLevelPathRecord[DateRootPath];

shared UnitsRootPath = DateLevelPathRecord[UnitsRootPath];

shared OrgCalculationsPath = DateLevelPathRecord[OrgCalculationsPath];

shared CurrentWorkbookFileName = DateLevelPathRecord[CurrentWorkbookFileName];

shared DateAlignment = let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateAlignment", type any}}),
    DateAlignment1 = #"Changed Type"{0}[DateAlignment]
in
    DateAlignment1;

shared #"IMPORT InefficienciesAGL2Day3210" = let
    Source = Excel.Workbook(File.Contents(OrgCalculationsPath & "\E-O-I\Inefficiencies.xlsx"), null, true),
    InefficienciesAG1_1Level321_Table = Source{[Item="InefficienciesAG1_1Level3210",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(InefficienciesAG1_1Level321_Table,{{"Level", type text}, {"Date", type date}, {"Role", type text}, {"Shift", type text}, {"D", type number}, {"Cs", type number}, {"CX", type number}, {"A", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}, {"AB EXCESS OVER-ALLOCATION", type number}, {"AB SPARE STRETCH", type number}, {"AB SPARE SLACK", type number}, {"AB SPARE STRETCH ALLOCATED", type number}, {"AB SPARE SLACK ALLOCATED", type number}, {"AB EXCESS STRETCH (ALLOCATED)", type number}, {"AB EXCESS ALLOCATED SLACK", type number}, {"AB EXCESS STRETCH", type number}, {"AB EXCESS SLACK", type number}, {"AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"AB POTENTIAL SHORTFALL", type number}, {"AB WASTED STRETCH", type number}, {"AB WASTED SLACK", type number}, {"AB ALLOCATED STRETCH", type number}, {"AB ALLOCATED SLACK", type number}, {"AB UNALLOCATED SLACK", type number}, {"AB LATENT", type number}, {"AB LATENT.ALLOCATED", type number}, {"AB  LATENT.OVERALLOCATED", type number}, {"AB Potential", type number}, {"Epn", type number}, {"EpX", type number}, {"Ein", type number}, {"EF EXCESS OVER-ALLOCATION", type number}, {"EF SPARE STRETCH", type number}, {"EF SPARE SLACK", type number}, {"EF SPARE STRETCH ALLOCATED", type number}, {"EF SPARE SLACK ALLOCATED", type number}, {"EF EXCESS STRETCH (ALLOCATED)", type number}, {"EF EXCESS ALLOCATED SLACK", type number}, {"EF EXCESS STRETCH", type number}, {"EF EXCESS SLACK", type number}, {"EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"EF POTENTIAL SHORTFALL", type number}, {"EF WASTED STRETCH", type number}, {"EF WASTED SLACK", type number}, {"EF ALLOCATED STRETCH", type number}, {"EF ALLOCATED SLACK", type number}, {"EF UNALLOCATED SLACK", type number}, {"EF LATENT", type number}, {"EF LATENT.ALLOCATED", type number}, {"EF LATENT.OVERALLOCATED", type number}, {"EF Potential", type number}, {"Ipn", type number}, {"Ipx", type number}, {"Iin", type number}, {"IN EXCESS OVER-ALLOCATION", type number}, {"IN SPARE STRETCH", type number}, {"IN SPARE SLACK", type number}, {"IN SPARE STRETCH ALLOCATED", type number}, {"IN SPARE SLACK ALLOCATED", type number}, {"IN EXCESS STRETCH (ALLOCATED)", type number}, {"IN EXCESS ALLOCATED SLACK", type number}, {"IN EXCESS STRETCH", type number}, {"IN EXCESS SLACK", type number}, {"IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"IN POTENTIAL SHORTFALL", type number}, {"IN WASTED STRETCH", type number}, {"IN WASTED SLACK", type number}, {"IN ALLOCATED STRETCH", type number}, {"IN ALLOCATED SLACK", type number}, {"IN UNALLOCATED SLACK", type number}, {"IN LATENT", type number}, {"IN LATENT.ALLOCATED", type number}, {"IN LATENT.OVERALLOCATED", type number}, {"IN Potential", type number}})
in
    #"Changed Type1";

shared InefficienciesAGL2Day3210 = let
    Source1 = #"IMPORT InefficienciesAGL2Day3210",
    #"Duplicated Column" = Table.DuplicateColumn(Source1, "Role", "Role ID"),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Duplicated Column", "Role ID", "Role Abvrev"),
    #"OLD TABLEAU COLS" = Table.RenameColumns(#"Duplicated Column1",{{"Apn", "Ability PoN AG"}, {"Apx", "Ability PoX AG"}, {"Ai", "Ability Pl AG"}, {"Epn", "Efficiency PoN  AG"}, {"EpX", "Efficiency PoX  AG"}, {"Ein", "Efficiency Pl AG"}, {"Ipn", "Workload PoN  AG"}, {"Ipx", "Workload PoX AG"}, {"Iin", "Workload Pl AG"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"OLD TABLEAU COLS",{"Level", "Date", "Role", "Role ID", "Role Abvrev", "Shift", "D", "Cs", "CX", "A", "Ability PoN AG", "Ability PoX AG", "Ability Pl AG", "AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential", "Efficiency PoN  AG", "Efficiency PoX  AG", "Efficiency Pl AG", "EF EXCESS OVER-ALLOCATION", "EF SPARE STRETCH", "EF SPARE SLACK", "EF SPARE STRETCH ALLOCATED", "EF SPARE SLACK ALLOCATED", "EF EXCESS STRETCH (ALLOCATED)", "EF EXCESS ALLOCATED SLACK", "EF EXCESS STRETCH", "EF EXCESS SLACK", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL", "EF WASTED STRETCH", "EF WASTED SLACK", "EF ALLOCATED STRETCH", "EF ALLOCATED SLACK", "EF UNALLOCATED SLACK", "EF LATENT", "EF LATENT.ALLOCATED", "EF LATENT.OVERALLOCATED", "EF Potential", "Workload PoN  AG", "Workload PoX AG", "Workload Pl AG", "IN EXCESS OVER-ALLOCATION", "IN SPARE STRETCH", "IN SPARE SLACK", "IN SPARE STRETCH ALLOCATED", "IN SPARE SLACK ALLOCATED", "IN EXCESS STRETCH (ALLOCATED)", "IN EXCESS ALLOCATED SLACK", "IN EXCESS STRETCH", "IN EXCESS SLACK", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL", "IN WASTED STRETCH", "IN WASTED SLACK", "IN ALLOCATED STRETCH", "IN ALLOCATED SLACK", "IN UNALLOCATED SLACK", "IN LATENT", "IN LATENT.ALLOCATED", "IN LATENT.OVERALLOCATED", "IN Potential"}),
    #"IN TO WL" = Table.RenameColumns(#"Reordered Columns",{{"IN EXCESS OVER-ALLOCATION", "WL EXCESS OVER-ALLOCATION"}, {"IN SPARE STRETCH", "WL SPARE STRETCH"}, {"IN SPARE SLACK", "WL SPARE SLACK"}, {"IN SPARE STRETCH ALLOCATED", "WL SPARE STRETCH ALLOCATED"}, {"IN SPARE SLACK ALLOCATED", "WL SPARE SLACK ALLOCATED"}, {"IN EXCESS STRETCH (ALLOCATED)", "WL EXCESS STRETCH (ALLOCATED)"}, {"IN EXCESS ALLOCATED SLACK", "WL EXCESS ALLOCATED SLACK"}, {"IN EXCESS STRETCH", "WL EXCESS STRETCH"}, {"IN EXCESS SLACK", "WL EXCESS SLACK"}, {"IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "WL POTENTIAL SHORTFALL (OVER-ALLOCATED)"}, {"IN POTENTIAL SHORTFALL", "WL POTENTIAL SHORTFALL"}, {"IN WASTED STRETCH", "WL WASTED STRETCH"}, {"IN WASTED SLACK", "WL WASTED SLACK"}, {"IN ALLOCATED STRETCH", "WL ALLOCATED STRETCH"}, {"IN ALLOCATED SLACK", "WL ALLOCATED SLACK"}, {"IN UNALLOCATED SLACK", "WL UNALLOCATED SLACK"}, {"IN LATENT", "WL  LATENT"}, {"IN LATENT.ALLOCATED", "WL LATENT.ALLOCATED"}, {"IN LATENT.OVERALLOCATED", "WL LATENT.OVERALLOCATED"}, {"IN Potential", "WL Potential"}}),
    #"Sorted Rows" = Table.Sort(#"IN TO WL",{{"Facility", Order.Ascending}}),
    #"Removed Duplicates" = Table.Distinct(#"Sorted Rows")
in
    #"Removed Duplicates";