// Power Query from: Allocation.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\Allocation.xlsx
// Extracted: 2026-09-24T05:33:12.855Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve this workbook's calculation folder through canonical FilePathUrl and CentriSyncPaths.
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

shared #"FilePath - 2Calculations" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared AllocationExtracted1 = let
   

    Source = Excel.Workbook(File.Contents( #"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Code", Int64.Type}, {"Date", type date}, {"Start", type datetime}, {"End", type datetime}, {"Break", Int64.Type},  {"Hours", type number}, {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Unit", type any}, {"Name", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Start", Order.Ascending}})
in
    #"Sorted Rows";

shared PermutationDimensions = let


    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" &  "\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Shifts", type text}, {"RolesList", type text}})
in
    #"Changed Type";

shared Table_RoleShiftAllocation = let
  

    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\AllocationByShiftAverage.xlsx"), null, true),
    Table_RoleShiftIntervalAllocation_Table = Source{[Item="Table_RoleShiftIntervalAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_RoleShiftIntervalAllocation_Table,{{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"Role", type text}, {"RoleShiftEffort", type number}, {"RoleShiftFTE", type number}})
in
    #"Changed Type";

shared Shifts = let


    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    Shifts_Table = Source{[Item="Shifts",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Shifts_Table,{{"Shifts", type text}})
in
    #"Changed Type";

shared ResourceShiftAllocation___less_Intervals = let

    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" &  "\AllocationByShiftAverage.xlsx"), null, true),
    ResourceShiftAllocation___less_Intervals_Table = Source{[Item="ResourceShiftAllocation___less_Intervals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResourceShiftAllocation___less_Intervals_Table,{{"ShiftPeriod", type text}, {"Name", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}})
in
    #"Changed Type";

shared #"IMPORT MealBreak" = let
    
    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" &  "\Settings Data.xlsx"), null, true),
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared #"Table_RoleShiftAllocation (2)" = let
    Source = Table_RoleShiftAllocation,
    #"Grouped Rows" = Table.Group(Source, {"ShiftDate"}, {{"Count", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared #"PermutationDimensions DATESHIFTROLE" = let
    Source = PermutationDimensions,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Shifts", "RolesList"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"})
in
    #"Removed Columns";

shared StaffList = let
    Source = AllocationExtracted1,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Date", "End", "Unit", "Start"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates";

shared Allocation = let
    Source = Table_RoleShiftAllocation,
    #"Merged Queries" = Table.NestedJoin(Source, {"ShiftDate", "ShiftPeriod", "Role"}, #"PermutationDimensions DATESHIFTROLE", {"Date", "Shifts", "RolesList"}, "PermutationDimensions DATESHIFTROLE", JoinKind.RightOuter),
    #"Expanded PermutationDimensions DATESHIFTROLE" = Table.ExpandTableColumn(#"Merged Queries", "PermutationDimensions DATESHIFTROLE", {"Date", "RolesList", "Shifts"}, {"PermutationDimensions DATESHIFTROLE.Date", "PermutationDimensions DATESHIFTROLE.RolesList", "PermutationDimensions DATESHIFTROLE.Shifts"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded PermutationDimensions DATESHIFTROLE",{"ShiftDate", "ShiftPeriod", "Role"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"PermutationDimensions DATESHIFTROLE.Date", "ShiftDate"}, {"PermutationDimensions DATESHIFTROLE.RolesList", "Role"}, {"PermutationDimensions DATESHIFTROLE.Shifts", "ShiftPeriod"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"ShiftDate", "Role", "ShiftPeriod", "RoleShiftEffort", "RoleShiftFTE"}),
    #"Replaced Value" = Table.ReplaceValue(#"Reordered Columns",null,0,Replacer.ReplaceValue,{"RoleShiftEffort", "RoleShiftFTE"})
in
    #"Replaced Value";

shared MealBreakTime = let
    Source = #"IMPORT MealBreak",
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared MealBreakStart = let
    Source = #"IMPORT MealBreak",
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MinShiftGap = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap;

// Query: ShiftStart
// Purpose: Read shift start settings through the authoritative calculation-folder path.
shared ShiftStart = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    ShiftStart_Table = Source{[Item="ShiftStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftStart_Table,{{"StartAM", type number}, {"StartPM", type number}, {"StartNIGHT", type number}}),
    #"Rounded Off" = Table.TransformColumns(#"Changed Type",{{"StartAM", each Number.Round(_, 3), type number}, {"StartPM", each Number.Round(_, 3), type number}, {"StartNIGHT", each Number.Round(_, 3), type number}})
in
    #"Rounded Off";

shared Start_AM = let
    Source = ShiftStart,
    StartAM = Source{0}[StartAM]
in
    StartAM;

shared StartPM = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartPM]
in
    StartPM1;

shared StartNIGHT = let
    Source = ShiftStart,
    StartNIGHT1 = Source{0}[StartNIGHT]
in
    StartNIGHT1;

shared ShiftLength = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    ShiftLent_Table = Source{[Item="ShiftLent",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftLent_Table,{{"MinHrs", Int64.Type}, {"ShortHrs", Int64.Type}, {"StdHrs", Int64.Type}, {"1.5OT", type number}, {"2.0OT", Int64.Type}})
in
    #"Changed Type";

shared MinHrs = let
    Source = ShiftLength,
    MinHrs = Source{0}[MinHrs]
in
    MinHrs;

shared ShortHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[ShortHrs]
in
    ShortHrs;

shared StdHrs = let
    Source = ShiftLength,
    ShortHrs = Source{0}[StdHrs]
in
    ShortHrs;

shared #"15OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[1.5OT]
in
    ShortHrs;

shared #"20OT" = let
    Source = ShiftLength,
    ShortHrs = Source{0}[2.0OT]
in
    ShortHrs;

shared ResourceShiftAllocation = let
    Source = ResourceShiftAllocation___less_Intervals,
    #"Filtered Rows" = Table.SelectRows(Source, each ([ResShiftFTE] > FullShiftThreshold))
in
    #"Filtered Rows";

shared FullShiftThreshold = 0 meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

shared #"FilePath - 1Input" = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;
