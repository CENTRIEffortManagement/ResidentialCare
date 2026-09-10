// Power Query from: AvailabilitiesExtracted.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\AvailabilitiesExtracted.xlsx
// Extracted: 2026-09-10T09:03:55.810Z

section Section1;

shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
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

shared #"FilePath-2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"IMPORT AvailabilityDayShiftMATRIXRaw" = let

    Source = Excel.Workbook(File.Contents(#"FilePath-1Input" & "\whiddon_availability_leave_extraction.xlsx"), null, true),
    #"Combined Output_Sheet" = Source{[Item="Combined Output",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(#"Combined Output_Sheet", [PromoteAllScalars=true]),
    #"Changed Type2" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Site Name", type text}, {"Employee Name", type text}, {"Payroll Code", Int64.Type}, {"Department", type text}, {"Date From", type datetime}, {"Date To", type datetime}, {"Availability or Leave Reason", type text}})
in
    #"Changed Type2";

shared #"FilePath-1Input" = let
    Source = UnitL1PathTABLE,
    #"Replaced Value" = Table.ReplaceValue(Source,"2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{1}[Value]
in
    Value;

shared #"IMPORT Combined Availabilities" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-1Input" & "\whiddon_availability_leave_extraction.xlsx"), null, true),
    #"Combined Output_Sheet" = Source{[Item="Combined Output",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Combined Output_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Site Name", type text}, {"Employee Name", type text}, {"Payroll Code", type text}, {"Department", type text}, {"Date From", type datetime}, {"Date To", type datetime}, {"Availability or Leave Reason", type text}})
in
    #"Changed Type";

shared #"Availabilities Prepare" = let
    Source = #"IMPORT Combined Availabilities",
    Custom1 = Table.Buffer(Source),
    #"FACILITY-ABBREV" = Table.AddColumn(Custom1, "Facility-Abbrev", each Text.Start([Site Name], 2), type text),
    #"Filtered FACILITY" = Table.SelectRows(#"FACILITY-ABBREV", each ([#"Facility-Abbrev"] <> "NR")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered FACILITY",{"Employee Name", "Department"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each true),
    #"Added Conditional Column" = Table.AddColumn(#"Filtered Rows1", "Available", each if Text.Contains([Availability or Leave Reason], "UNAVAIL") then false else if Text.Contains([Availability or Leave Reason], "AVAIL") then true else false, type logical),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Conditional Column",{{"Payroll Code", type text}})
in
    #"Changed Type";

shared #"Availabilities-Employees" = let
    Source = #"Availabilities Prepare",
    #"Replaced Errors" = Table.ReplaceErrorValues(Source, {{"Payroll Code", "999999"}}),
    #"Filtered ERROR CODE" = Table.SelectRows(#"Replaced Errors", each ([Payroll Code] <> "999999"))
in
    #"Filtered ERROR CODE";