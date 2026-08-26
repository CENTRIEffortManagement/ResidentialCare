// Power Query from: 2DRead.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\EOW\2DRead.xlsx
// Extracted: 2026-05-21T00:47:37.301Z

section Section1;

shared Boundaries = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\EOW\Grid Thresholds.xlsx"), null, true),
    Boundaries_Table = Source{[Item="Boundaries",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Boundaries_Table,{{"Point", type text}, {"Type", type text}, {"Threshold", type text}, {"Position", type text}, {"D", type any}, {"C", type any}, {"A", type any}, {"Ap", type any}, {"Ei", type any}, {"Ii", type any}, {"Ai", type any}, {"Ep", type any}, {"Ip", type any}, {"X'", type any}, {"Y'", type any}, {"Date", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"D", "Demand"}, {"C", "Capacity"}, {"A", "Allocation"}, {"Ap", "Apn"}, {"Ei", "Ein"}, {"Ii", "Iin"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Ai", "Ep", "Ip"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"X'", "Y'", "Point", "Type", "Threshold", "Position", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin"})
in
    #"Reordered Columns";

shared #"Effort+OutcomesXYImport" = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\EOW\EffortOutcomeLogXY.xlsx"), null, true),
    EffortOutcomesAG0_Table = Source{[Item="EffortOutcomesAG0",Kind="Table"]}[Data],
    #"Removed Columns1" = Table.RemoveColumns(EffortOutcomesAG0_Table,{"X", "Y"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns1",{"Date", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin", "Apx", "Ai", "Epn", "EpX", "Ipn", "Ipx", "Role", "WARDS", "Shifts", "Column1", "Column12"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns1",{"Xp", "Yp", "Zp", "KPx", "Kpy", "KPz", "KPx'i", "KPx'j", "KPx'k", "Kpy'i", "KPy'j", "KPy'k", "Column2"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"X'", "Y'", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin", "WARDS", "Role", "Shifts"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Date", type date}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Changed Type",{"X'", "Y'", "Date", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin", "Apx", "Ai", "Epn", "EpX", "Ipn", "Ipx", "WARDS", "Role", "Shifts", "Column1", "Column12", "Extreme"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Reordered Columns2",{"Column1", "Column12"})
in
    #"Removed Columns2";

shared #"Effort+OutcomesXYImportAG1-1DS" = let
    Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\EOW\EffortOutcomeLogXY.xlsx"), null, true),
    #"Effort+OutcomesXYAG1-1DS_Sheet" = Source{[Item="Effort+OutcomesXYAG1-1DS",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Effort+OutcomesXYAG1-1DS_Sheet", [PromoteAllScalars=true]),
    #"Removed Columns1" = Table.RemoveColumns(#"Promoted Headers",{"X", "Y"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns1",{"Date", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin", "Apx", "Ai", "Epn", "EpX", "Ipn", "Ipx", "Role",  "Shift", "Column1"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns1",{"Xp", "Yp", "Zp", "KPx", "Kpy", "KPz", "KPx'i", "KPx'j", "KPx'k", "Kpy'i", "KPy'j", "KPy'k", "Column2"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"X'", "Y'", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin",  "Role", "Shift"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Date", type date}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Changed Type",{"X'", "Y'", "Date", "Point", "Demand", "Capacity", "Allocation", "Apn", "Ein", "Iin", "Apx", "Ai", "Epn", "EpX", "Ipn", "Ipx", "Role", "Shift", "Column1",  "Extreme"})
in
    #"Reordered Columns2";

shared #"OutcomeXYall- Out" = let
    Source = Table.Combine({Boundaries, #"Effort+OutcomesXYImport"})
in
    Source;

shared #"Effort+OutcomesXYImport Check" = let
    Source = #"Effort+OutcomesXYImport",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"X'", "Y'", "Date", "Demand", "WARDS", "Role", "Shifts"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Other Columns", each ([Date] <> null)),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Filtered Rows1", {{"Date", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Filtered Rows1", {{"Date", type text}}, "en-AU")[Date]), "Date", "Demand", List.Sum),
    #"Filtered Rows" = Table.SelectRows(#"Pivoted Column", each ([WARDS] <> "Other"))
in
    #"Filtered Rows";

shared #"OutcomeXYall- OutAG1-1DS" = let
    Source = Table.Combine({Boundaries, #"Effort+OutcomesXYImportAG1-1DS"})
in
    Source;