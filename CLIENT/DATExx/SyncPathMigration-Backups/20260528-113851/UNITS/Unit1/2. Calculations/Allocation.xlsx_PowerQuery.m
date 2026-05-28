// Power Query from: Allocation.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Allocation.xlsx
// Extracted: 2026-05-18T06:14:05.148Z

section Section1;

shared UnitL1PathTABLE = // Version 25.00 250107   
// Updated 250107

let

    
    // PART A - Define FilePathUrl and Replaced Value
    FilePathUrl = 
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content],
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
        Core_ExtractedTextBeforeDelimiter = Table.TransformColumns(Core_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
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
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{3, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,

    // Extract Unit directly
    Unit = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{0, RelativePosition.FromEnd}), 
                    "\")
    in
        Extracted,

    // Extract Date directly
    Date = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{2, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
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