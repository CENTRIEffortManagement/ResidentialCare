// Power Query from: EffortOutcomes.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\E-O-I\EffortOutcomes.xlsx
// Extracted: 2026-05-21T00:47:28.279Z

section Section1;

shared #"EffortOutcomesAG1-1DayShift" = let
    Source = #"IMPORT Table_EffortAllMatrixAG1_1D !!",
    Apn = Table.AddColumn(Source, "Apn", each if [Demand]=0 then null 
else [Capacity]/[Demand]),
    Apx = Table.AddColumn(Apn, "Apx", each if [Demand]=0 then null 
else [CapacityX]/[Demand]),
    Ai = Table.AddColumn(Apx, "Ai", each if [Demand]=0 then null 
else [Allocation]/[Demand]),
    Epn = Table.AddColumn(Ai, "Epn", each if [CapacityX]=0 then null 
else [Demand]/[Capacity]),
    Epx = Table.AddColumn(Epn, "EpX", each if [Capacity]=0 then null 
else [Demand]/[CapacityX]),
    Ein = Table.AddColumn(Epx, "Ein", each if [Capacity]=0 then null 
else 
[Allocation]/
[Capacity]),
    Ipn = Table.AddColumn(Ein, "Ipn", each if [Allocation]=0 then null 
else [Capacity] / [Allocation]),
    Ipx = Table.AddColumn(Ipn, "Ipx", each if [Allocation]=0 then null 
else [CapacityX]/[Allocation]),
    Iin = Table.AddColumn(Ipx, "Iin", each if [Allocation]=0 then null 
else [Demand]/[Allocation]),
    #"Filtered Rows" = Table.SelectRows(Iin, each true),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Date", type date}})
in
    #"Changed Type";

shared #"Folder !!" = let
    Source = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    #"Extracted Text Before Delimiter" = Table.TransformColumns(Source, {{"Folder", each Text.BeforeDelimiter(_, "\BERR"), type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter",{{"Folder", type text}}),
    Folder1 = #"Changed Type"{0}[Folder]
in
    Folder1;

shared #"IMPORT Table_EffortAllMatrixAG1_1D !!" = let
      
    Source = Excel.Workbook(File.Contents(#"Folder !!"&"\2. Calculations\E-O-I\Effort-All.xlsx"), null, true),
    EffortAllMatrixAG1_1D_2_Table = Source{[Item="EffortAllMatrixAG1_1D_2",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_2_Table,{{"Facility", type text}})
in
    #"Changed Type";

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


