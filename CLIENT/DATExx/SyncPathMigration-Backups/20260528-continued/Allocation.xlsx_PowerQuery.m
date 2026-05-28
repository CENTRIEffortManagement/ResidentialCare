// Power Query from: Allocation.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Allocation.xlsx
// Extracted: 2026-05-18T06:14:05.148Z

section Section1;

shared UnitL1PathTABLE = // Version 25.01 local-first ResidentialCare
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
    WorkbookFolder = Text.BeforeDelimiter(RawFilePath, "\", {0, RelativePosition.FromEnd}),
    IsLocalPath = (Text.Contains(WorkbookFolder, ":\") or Text.StartsWith(WorkbookFolder, "\\")) and not Text.StartsWith(Text.Lower(WorkbookFolder), "http"),
    RootPath = if IsLocalPath then WorkbookFolder else error "FilePathUrl did not resolve to a local path. Open/save this workbook from the active local sync folder before refreshing.",
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(RawFilePath, "[", "]") otherwise Text.AfterDelimiter(RawFilePath, "\", {0, RelativePosition.FromEnd}),
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"FilePathUrl", RawFilePath},
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
    Value = Source{1}[Value]
in
    Value;

shared AllocationExtracted1 = let
   

    Source = Excel.Workbook(File.Contents( #"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Table = Source{[Item="AllocationExtracted",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AllocationExtracted_Table,{{"Code", Int64.Type}, {"Date", type date}, {"Start", type datetime}, {"End", type datetime}, {"Break", Int64.Type}, {"Break Time", type datetime}, {"Hours", type number}, {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Unit", type any}, {"Name", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Date] = #date(2024, 10, 10))),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Start", Order.Ascending}})
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
    Table_RoleShiftAllocation_Table = Source{[Item="Table_RoleShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_RoleShiftAllocation_Table,{{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"Role", type text}, {"RoleShiftEffort", type number}, {"RoleShiftFTE", type number}})
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
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Resrodel\Customers - Documents\IndoChinese Aged Care\4. Analysis\221223\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftGap_Table = Source{[Item="ShiftGap",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftGap_Table,{{"MinGap", Int64.Type}}),
    MinGap = #"Changed Type"{0}[MinGap]
in
    MinGap;

shared ShiftStart = let
        Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder1 = Source1{0}[Folder],

    Source = Excel.Workbook(File.Contents(Folder1 & "\2. Calculations\Settings Data.xlsx"), null, true),
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
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Resrodel\Customers - Documents\IndoChinese Aged Care\4. Analysis\221223\2. Calculations\SETTINGS.xlsx"), null, true),
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
    #"Filtered Rows" = Table.SelectRows(Source, each [Variable Name] = "Root Path"),
    #"Replaced Value1" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value1"{0}[Value]
in
    Value;
