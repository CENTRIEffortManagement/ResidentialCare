// Power Query from: 1-AllocationExtracted.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\1. Input\1-AllocationExtracted.xlsx
// Extracted: 2026-08-26T09:58:02.786Z

section Section1;

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"Filter AGENCY" = Table.SelectRows(Source, each ([Employment Type] <> " " and [Employment Type] <> "Agency ")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filter AGENCY",{"Location", "Department", "Area", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Employee_Code"}),
    #"Extracted First Characters" = Table.TransformColumns(#"Removed Other Columns", {{"Location", each Text.Start(_, 2), type text}}),
    #"Renamed Columns5" = Table.RenameColumns(#"Extracted First Characters",{{"Role", "RoleShift"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns5",{{"Employee_Code", "Code"}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Renamed Columns",{{"Employee Roster Name", "Name"}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Renamed Columns1",{{"Start Time", "Start"}}),
    #"Renamed Columns3" = Table.RenameColumns(#"Renamed Columns2",{{"End Time", "End"}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Renamed Columns3",{{"End", "Finish"}, {"Break Length Minutes", "Break"}, {"Shift Net Length", "Hours"}}),
    Custom1 = Table.AddColumn(#"Renamed Columns4", "Role", each if [Department] <> "Care" and  [Department] <> "Catering" 
         then [RoleShift]
         else 
            Text.Start([RoleShift], Text.Length([RoleShift])-3)),
    #"Reordered Columns" = Table.ReorderColumns(Custom1,{"Code", "Name", "Date", "Start", "Finish", "Break", "Hours", "Location", "Department", "Area", "Role"}),
    #"Filtered ROLES" = Table.SelectRows(#"Reordered Columns", each ([Role] = "Asst in Nursing" or [Role] = "Asst in Nursing " or [Role] = "Asst in Nursing Med Comp" or [Role] = "Enrolled Nurse" or [Role] = "REGN" or [Role] = "REGN -  In Charge" or [Role] = "REGN - In Charge")),
    BUFFER = Table.Buffer(#"Filtered ROLES")
in
    BUFFER;

shared #"Roles-Raw" = let
    Source = #"IMPORT Roster",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Role"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared RosteredDays = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", Int64.Type}}),
    #"Calculated Distinct Count" = List.NonNullCount(List.Distinct(#"Changed Type"[Date]))
in
    #"Calculated Distinct Count";

shared RosterStart = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Calculated Maximum" = List.Min(#"Removed Other Columns"[Date])
in
    #"Calculated Maximum";

shared AllocationExtraction = let
    Source = #"Allocation Prepare",
    #"Renamed Columns2" = Table.RenameColumns(Source,{{"Finish", "End"}}),
    #"Renamed Columns" = Table.RenameColumns(#"Renamed Columns2",{{"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Role", each if Text.Contains([RoleX], "REGN") then "RN" else if [RoleX] = "Asst in Nursing Med Comp" then "AINC4" else if Text.Contains([RoleX], "Asst") then "AIN" else if [RoleX] = "Enrolled Nurse" then "EN" else [RoleX]),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Role] = "AIN" or [Role] = "AINC4" or [Role] = "Asst in Nursing " or [Role] = "Enrolled Nurse" or [Role] = "RN")),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "Unit", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RoleX"})
in
    #"Removed Columns";

shared AllocationExtracted = let
    Source = AllocationExtraction,

    #"Merged ResDoubleRoles" = Table.NestedJoin(
        Source,
        {"Name"},
        ResDoubleRoles,
        {"Name"},
        "ResDoubleRoles",
        JoinKind.LeftOuter
    ),

    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(
        #"Merged ResDoubleRoles",
        "ResDoubleRoles",
        {"MultiRole"},
        {"MultiRole"}
    ),

    #"Renamed Columns" = Table.RenameColumns(
        #"Expanded ResDoubleRoles",
        {{"Name", "NameX"}}
    ),

    #"Inserted MULTIROLENAMES" = Table.AddColumn(
        #"Renamed Columns",
        "Name",
        each
            if [MultiRole] <> null then
                [NameX] & " (" & [Role] & ")"
            else
                [NameX],
        type text
    ),

    #"Removed Columns" = Table.RemoveColumns(
        #"Inserted MULTIROLENAMES",
        {"NameX", "MultiRole"}
    )

in
    #"Removed Columns";

shared PathTABLE = // Version 25.00    250107 
// Updated 250107

let

    // PART A - Define FilePathUrl and Replaced Value
    FilePathUrl = 
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        #"Renamed Columns" = Table.RenameColumns(Source, {{"Column1", "FilePath"}}),
        ReplacedValue = Table.ReplaceValue(#"Renamed Columns", "/", "\", Replacer.ReplaceText, {"FilePath"}),
        BufferedTable = Table.Buffer(ReplacedValue) // Buffer the table for better performance
    in 
        BufferedTable,

    // PART B - RootPath
    IMPORTRootPath =
    let
        // Step 1: Import CentriSyncPaths
        CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),

        CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
        CentriSyncPaths_ChangedType = Table.TransformColumnTypes(CentriSyncPaths_Table, {{"User", type text}, {"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),

        // Step 2: UrlSite - Use buffered FilePathUrl
        UrlSite_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "sites\"), type text}}),
        UrlSite_ExtractedTextBeforeDelimiter = Table.TransformColumns(UrlSite_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\"), type text}}),
        RenamedColumns1 = Table.RenameColumns(UrlSite_ExtractedTextBeforeDelimiter, {{"FilePath", "Site"}}),

        // Step 3: Prefix
        Prefix_NestedJoin = Table.NestedJoin(RenamedColumns1, {"Site"}, CentriSyncPaths_ChangedType, {"Site"}, "CentriSyncPaths", JoinKind.LeftOuter),
        Prefix_Expanded = Table.ExpandTableColumn(Prefix_NestedJoin, "CentriSyncPaths", {"SyncedFolderRootPath"}, {"Prefix"}),
        Prefix = Prefix_Expanded{0}[Prefix],

        // Step 4: Core
        Core_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "Shared Documents\"), type text}}),
        Core_ExtractedTextBeforeDelimiter = Table.TransformColumns(Core_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\", {1, RelativePosition.FromEnd}), type text}}),
        Core = Core_ExtractedTextBeforeDelimiter{0}[FilePath],

        // Step 5: FilePath
        FilePath = Prefix & "\" & Core,
        ConvertedToTable = #table(1, {{FilePath}}),
        RenamedColumns = Table.RenameColumns(ConvertedToTable, {{"Column1", "FilePath"}})
    in  
        RenamedColumns,

    // PART C - Dimensions
    // Extract UserName directly
    UserName = 
    let
        Source = IMPORTRootPath,
        Extracted = Text.BeforeDelimiter(Text.AfterDelimiter(Source[FilePath]{0}, "\", 1), "\")
    in
        Extracted,

    // Extract Client directly
    Client = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "HomeCare\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\", {0, RelativePosition.FromEnd}),
                        "\",{3, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,

    // Extract Unit directly
    Date = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "HomeCare\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\", {1, RelativePosition.FromEnd}),
                        "\",{0, RelativePosition.FromEnd}), 
                    "\")
    in
        Extracted,

    // Extract Filename
    FileName = 
    let
        Source = FilePathUrl,
        #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"FilePath", each Text.AfterDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
        #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"FilePath", each Text.BeforeDelimiter(_, "]"), type text}}),
        #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"FilePath"}),
        String = #"Replaced Value"{0}[FilePath]
    in 
        String,

    // PART D - Table
    // Create the table with variable names and their corresponding values
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", IMPORTRootPath{0}[FilePath]},
            {"FilePathUrl", FilePathUrl{0}[FilePath]},
            {"Client", Client},
            {"Date", Date},
            {"FileName", FileName}
        }
    )
in
    TABLE;

shared Path = let
    Source = #"PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","\2. Calculations","",Replacer.ReplaceText,{"Folder"}),
    Folder = #"Replaced Value"{0}[Folder]
in
    Folder;

shared DayAdjustment = let
    Source = Excel.CurrentWorkbook(){[Name="DayAdjustment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayAdjustment", Int64.Type}}),
    DayAdjustment1 = #"Changed Type"{0}[DayAdjustment]
in
    DayAdjustment1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents(Unit1Path & "\Allocation\BD,TE,JE,RY Published Roster Data 20.07.26-16.08.26.xlsx"), null, true),
    Beaudesert_Sheet = Source{[Item="Beaudesert",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Beaudesert_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Employee Roster Name", type text}, {"Employment Type", type text}, {"Role", type text}, {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start Time", type time}, {"End Time", type time}, {"Break Length Minutes", Int64.Type}, {"Shift Length", type number}, {"Shift Net Length", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Employee_Code", Int64.Type}})
in
    #"Changed Type";

shared AllocationExtractedCheck = let
    Source = AllocationExtraction,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom2" = Table.AddColumn(#"Changed Type1", "Subtraction", each if [End] < [Start]
then ([End]) + (1-[Start])
else  [End]-[Start]),
    #"Filtered Rows3" = Table.SelectRows(#"Added Custom2", each true),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows3", each true),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows2",{{"Subtraction", type number}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Role"}, {{"Days", each List.Sum([Subtraction]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "AllocationTime", each [Days]*24/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "AllocatedEffort", each [AllocationTime]*0.9366)
in
    #"Added Custom1";

shared AllocatedStaffList = let
    Source = AllocationExtracted,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Role", Order.Ascending}, {"Name", Order.Ascending}})
in
    #"Sorted Rows";

[ Description = "BY PASSED" ]
shared ResDoubleRoles = let
    Source = AllocationExtraction,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Count] = 2)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "MultiRole", each "MultiRole")
in
    #"Added Custom";

shared ResRolesProportion = let
    Source = Table.NestedJoin(AllocationExtraction, {"Name"}, ResDoubleRoles, {"Name"}, "ResDoubleRoles", JoinKind.LeftOuter),
    #"Expanded ResDoubleRoles" = Table.ExpandTableColumn(Source, "ResDoubleRoles", {"Name"}, {"Name.1"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDoubleRoles", each ([Name.1] <> null)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Name", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Name", "Date", "Start", "Hours", "Role"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"Name", "Role"}, {{"Hours", each List.Sum([Hours]), type nullable number}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", List.Distinct(#"Grouped Rows"[Role]), "Role", "Hours", List.Sum),
    #"Inserted Addition" = Table.AddColumn(#"Pivoted Column", "Addition", each [AIN] + [AINC4], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Addition",{{"Addition", "Total"}}),
    #"Inserted Division" = Table.AddColumn(#"Renamed Columns", "AINC4HrsAvilPref", each [AINC4] / [Total], type number),
    #"Rounded Off" = Table.TransformColumns(#"Inserted Division",{{"AINC4HrsAvilPref", each Number.Round(_, 1), type number}})
in
    #"Rounded Off";

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

shared #"UnitL1PathTABLE (2)" = // Version 25.02 flexible ResidentialCare
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

shared Unit1Path = let
    Source = #"UnitL1PathTABLE (2)",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","\2. Calculations","",Replacer.ReplaceText,{"Folder"}),
    Folder = #"Replaced Value"{0}[Folder]
in
    Folder;

shared Unit = let
    Source = #"UnitL1PathTABLE (2)",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;