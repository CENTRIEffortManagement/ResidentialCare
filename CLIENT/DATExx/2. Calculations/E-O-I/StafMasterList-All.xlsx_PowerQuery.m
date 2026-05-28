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


