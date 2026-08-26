// Power Query from: MutliRoleCheck.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\MutliRoleCheck.xlsx
// Extracted: 2026-05-18T06:14:26.566Z

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

shared FilePath = let
    Source = UnitL1PathTABLE,
    Value = Source{1}[Value]
in
    Value;

shared #"IMPORT MultiRoles" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Capacity-ShiftAvailability.xlsx"), null, true),
    MultiRoles_Table = Source{[Item="MultiRoles",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(MultiRoles_Table,{{"Name", type text}, {"AINC4HrsAvilPref", type number}, {"MultiRole", type text}})
in
    #"Changed Type";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared #"IMPORT NAMES_INTERVALS" = let
    Source = Excel.Workbook(File.Contents(FilePath & "\Shifts.xlsx"), null, true),
    NAMES_INTERVALS_Table = Source{[Item="NAMES_INTERVALS",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(NAMES_INTERVALS_Table,{{"DayDate", type date}, {"Role", type text}, {"IntervalStart", type datetime}, {"TimeType", type text}, {"IntervalEnd", type datetime}, {"Name", type text}, {"IntervalID.1", Int64.Type}, {"AllocatedIntervals", Int64.Type}, {"IntervalDurationTemp", type number}, {"ShiftPeriod", type text}, {"Double Shift", type any}, {"Effective Duration", type number}, {"RealDuration", type number}, {"ShiftType", type text}})
in
    #"Changed Type";

shared #"IMPORT MultiRoles (2)" = let
    Source = #"IMPORT MultiRoles"
in
    Source;

shared MultiRoles = let
    Source = #"IMPORT MultiRoles",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns")
in
    #"Removed Duplicates";

shared MultiRoleMATRIX = let
    Source = #"IMPORT NAMES_INTERVALS",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"DayDate", "Role", "Name", "ShiftPeriod"}),
    BUFFER = Table.Buffer(#"Removed Other Columns"),
    #"Filtered Rows" = Table.SelectRows(BUFFER, each ([Name] <> "x")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "(")),
    #"Merged Queries1" = Table.NestedJoin(#"Filtered Rows1", {"Name"}, MultiRoles, {"Name"}, "MultiRoles", JoinKind.LeftOuter),
    #"Expanded MultiRoles" = Table.ExpandTableColumn(#"Merged Queries1", "MultiRoles", {"Name"}, {"Name.1"}),
    #"Added Custom" = Table.AddColumn(#"Expanded MultiRoles", "Custom", each "X"),
    #"Merged Queries" = Table.NestedJoin(#"Added Custom", {"DayDate", "Role", "ShiftPeriod"}, #"IMPORT PermutationDimensions", {"Date", "RolesList", "Shifts"}, "IMPORT PermutationDimensions", JoinKind.RightOuter),
    #"Expanded IMPORT PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions", {"Period"}, {"Period"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded IMPORT PermutationDimensions",{"DayDate", "ShiftPeriod"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Period", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Custom"),
    #"Sorted Rows1" = Table.Sort(#"Pivoted Column",{{"Name", Order.Ascending}})
in
    #"Sorted Rows1";