// Power Query from: StafMasterList-All.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\E-O-I\StafMasterList-All.xlsx
// Extracted: 2026-05-21T00:47:34.382Z

section Section1;

shared IMPORTRootPath = let
    Source = Table.FromColumns({Lines.FromBinary(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\RootPath.txt"), null, null, 1252)}),
    Column1 = Source{0}[Column1]
in
    Column1;

shared #"INPUT FilePath" = let
    Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Column1", "String"}})
in
    #"Renamed Columns";

shared Client = let
    Source = #"INPUT FilePath",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 7), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared Date = let
    Source = #"INPUT FilePath",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 8), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared Facility = let
    Source = #"INPUT FilePath",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 10), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared #"FilePath-Role" = let
    Source = IMPORTRootPath&"\"&Client&"\"&Date&"\"&"FACILITIES"&"\"&Facility
in
    Source;

shared Role = let
    Source = #"INPUT FilePath",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", {1, RelativePosition.FromEnd}), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared FileName = let
    Source = #"INPUT FilePath",
    #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", {0, RelativePosition.FromEnd}), type text}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"String", each Text.BeforeDelimiter(_, "]"), type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"String"}),
    String = #"Replaced Value"{0}[String]
in
    String;

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

shared Folder = let
    Source = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Folder", type text}}),
    Folder1 = #"Changed Type"{0}[Folder]
in
    Folder1;

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

shared #"Facility1 - StaffList" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility1 & "\2. Calculations\StaffListMaster.xlsx"), null, true),
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(Table_Masterlist_Table,{{"Role", type text}, {"Name", type text}, {"Resource", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type2", "Facility", each Facility1)
in
    #"Added Custom";

shared #"Facility2 - StaffList" = let
        Source = Excel.Workbook(File.Contents(Folder & Facility2 & "\2. Calculations\StaffListMaster.xlsx"), null, true),
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

shared UnitL1PathTABLE = let
  
    
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