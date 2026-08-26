// Power Query from: 2-DemandExtracted-Master-.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\1. Input\2-DemandExtracted-Master-.xlsx
// Extracted: 2026-05-21T00:46:37.250Z

section Section1;

shared Facility = let
    Source = #"Table001 (Page 1)",
    Column3 = Source{2}[Column3]
in
    Column3;

shared MasterRosterJoined = let
    Source = Table.Combine({#"Table001 (Page 1)", #"Table002 (Page 2-5)", #"Table003 (Page 6-10)"}),
    #"Removed Top Rows" = Table.Skip(Source,3),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Top Rows",{{"Column1", "Role"}, {"Column2", "D1"}, {"Column3", "D2"}, {"Column4", "D3"}, {"Column5", "D4"}, {"Column6", "D5"}, {"Column7", "D6"}, {"Column8", "D7"}}),
    #"Filled Down" = Table.FillDown(#"Renamed Columns",{"Role"}),
    #"Filtered ROLE" = Table.SelectRows(#"Filled Down", each ([Role] = "Assistant in Nursing" or [Role] = "Assistant In Nursing" or [Role] = "Assistant in Nursing (Certificate IV)" or [Role] = "Registered Nurse")),
    #"Filtered Rows" = Table.SelectRows(#"Filtered ROLE", each ([Role] <> "Roster")),
    #"Added Index" = Table.AddIndexColumn(#"Filtered Rows", "Index", 1, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 0, 1, Int64.Type),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Added Index1", {"Index","Index.1", "Role"}, "Attribute", "Value"),
    #"Merged Queries1" = Table.NestedJoin(#"Unpivoted Other Columns", {"Index", "Attribute"}, #"Unpivoted Other Columns", {"Index.1", "Attribute"}, "Unpivoted Other Columns", JoinKind.LeftOuter),
    #"Expanded Unpivoted Other Columns" = Table.ExpandTableColumn(#"Merged Queries1", "Unpivoted Other Columns", {"Value"}, {"Value.1"}),
    #"Filtered Rows1" = Table.SelectRows(#"Expanded Unpivoted Other Columns", each Text.Contains([Value.1], ":"))
in
    #"Filtered Rows1";

shared #"MasterRosterJoined (2)" = let
    Source = MasterRosterJoined,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Attribute", "Value.1"}, {{"Count", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared #"Allocation Prepare" = let
    Source = #"IMPORT Roster",
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Code", Int64.Type}, {"Name", type text}, {"Date", type date}, {"Start", type time}, {"Finish", type time}, {"Break", Int64.Type}, {"Break Time", type time}, {"Hours", type number}, {"Cost", type text}, {"Location", type text}, {"Department", type text}, {"Area", type text}, {"Role", type text}, {"Period", type text}, {"Event", type text}, {"Function", type text}, {"Comments", type text}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type1", {"Date"}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Removed Errors", {{"Code", 9999}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Replaced Errors",{"Code", "Name", "Date", "Start", "Finish", "Break", "Break Time", "Hours", "Location", "Department", "Area", "Role"})
in
    #"Removed Other Columns";

shared AllocationExtracted = let
    Source = #"Allocation Prepare",
    #"Renamed Columns2" = Table.RenameColumns(Source,{{"Finish", "End"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns2", each ([Role] = "Assistant In Nursing" or [Role] = "Assistant in Nursing (Certificate IV)" or [Role] = "Registered Nurse")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","Registered Nurse","RN",Replacer.ReplaceText,{"Role"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Assistant in Nursing (Certificate IV)","AINC4",Replacer.ReplaceText,{"Role"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Assistant In Nursing","AIN",Replacer.ReplaceText,{"Role"}),
    #"Filtered Rows1" = Table.SelectRows(#"Replaced Value2", each true)
in
    #"Filtered Rows1";

shared RosteredDays = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", Int64.Type}}),
    #"Calculated Distinct Count" = List.NonNullCount(List.Distinct(#"Changed Type"[Date]))
in
    #"Calculated Distinct Count";

shared DayAdjustment = let
    Source = Excel.CurrentWorkbook(){[Name="DayAdjustment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayAdjustment", Int64.Type}}),
    DayAdjustment1 = #"Changed Type"{0}[DayAdjustment]
in
    DayAdjustment1;

shared #"IMPORT Roster" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.2 Christadelphian Homes\08 PoC\03 Incoming\240807\Ashburn working roster 15 to 28 July 24.xls"), null, true),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Start", type time}, {"Date", type date}})
in
    #"Changed Type";

shared #"04062023Carer" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\03. Incoming\230613 Maitland Data Emailed\Roster  - 22.5 -4.6.23xlsx.xlsx"), null, true),
    #"04062023_Sheet" = Source{[Item="04062023",Kind="Sheet"]}[Data],
    #"Filtered Rows" = Table.SelectRows(#"04062023_Sheet", each true)
in
    #"Filtered Rows";

shared #"RN Master 23-5-2022 (4)" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\03. Incoming\230613 Maitland Data Emailed\Roster  - 22.5 -4.6.23xlsx.xlsx"), null, true),
    #"RN Master 23-5-2022 (4)_Sheet" = Source{[Item="RN Master 23-5-2022 (4)",Kind="Sheet"]}[Data]
in
    #"RN Master 23-5-2022 (4)_Sheet";

shared AllocationExtractedCheck = let
    Source = #"AllocationExtracted",
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
    Source = #"AllocationExtracted",
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"})
in
    #"Removed Columns";

shared #"Table001 (Page 1)" = let
    Source = Pdf.Tables(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.2 Christadelphian Homes\08 PoC\03 Incoming\240807\Ashburn House Master Roster.pdf"), [Implementation="1.3"]),
    Table001 = Source{[Id="Table001"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table001,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}})
in
    #"Changed Type";

shared #"Table002 (Page 2-5)" = let
    Source = Pdf.Tables(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.2 Christadelphian Homes\08 PoC\03 Incoming\240807\Ashburn House Master Roster.pdf"), [Implementation="1.3"]),
    Table002 = Source{[Id="Table002"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table002,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type text}, {"Column14", type text}})
in
    #"Changed Type";

shared #"Table003 (Page 6-10)" = let
    Source = Pdf.Tables(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.2 Christadelphian Homes\08 PoC\03 Incoming\240807\Ashburn House Master Roster.pdf"), [Implementation="1.3"]),
    Table003 = Source{[Id="Table003"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table003,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}})
in
    #"Changed Type";

shared IMPORTRootPath = let
    Source = Table.FromColumns({Lines.FromBinary(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\RootPath.txt"), null, null, 1252)}),
    Column1 = Source{0}[Column1]
in
    Column1;

[ Description = "BUFFER" ]
shared #"INPUT FilePath B" = let
    Source = Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content],
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Column1", "String"}}),
    BUFFER = Table.Buffer(#"Renamed Columns")
in
    BUFFER;

shared FileName = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", {0, RelativePosition.FromEnd}), type text}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"String", each Text.BeforeDelimiter(_, "]"), type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"String"}),
    String = #"Replaced Value"{0}[String]
in
    String;

shared #"FilePath-Role" = let
    Source = IMPORTRootPath&"\"&Client&"\"&Date&"\"&"FACILITIES"&"\"&#"Facility (2)"
in
    Source;

shared Client = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 7), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared Date = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 8), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared #"Facility (2)" = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 10), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared RosterStart = let
    Source = #"Allocation Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Calculated Maximum" = List.Min(#"Removed Other Columns"[Date])
in
    #"Calculated Maximum";

shared #"File Path Data" = let
   

    // Create the table with variable names and their corresponding values
    Source = #table(
        {"Variable Name", "Value"},
        {
            {"Root Path", #"IMPORTRootPath"},
            {"FilePath", #"FilePath-Role"},
            {"Client", Client},
            {"Date", Date},
            {"Facility", #"Facility (2)"},
            {"FileName", FileName}
            
            
            
            
            
            
        }
    )
in
    Source;