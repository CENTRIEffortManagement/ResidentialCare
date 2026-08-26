// Power Query from: ResidualAllocation-ALL.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\Change\SS\ResidualAllocation-ALL.xlsx
// Extracted: 2026-05-21T00:48:13.720Z

section Section1;

shared Table_MinResidualHRSResDayShiftMAITRaw = let
     Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"), null, true),
    Table_MinResidualHRSResDayShift_Table = Source{[Item="Table_MinResidualHRSResDayShift",Kind="Table"]}[Data]
in
    Table_MinResidualHRSResDayShift_Table;

shared Table_MinResidualHRSResDayShiftMAIT = let
     Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"), null, true),
    Table_MinResidualHRSResDayShift_Table = Source{[Item="Table_MinResidualHRSResDayShift",Kind="Table"]}[Data],
    #"Renamed Columns" = Table.RenameColumns(Table_MinResidualHRSResDayShift_Table,{{"Date", "DateX"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX],MAIT_date_adjusted_for_example)),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"DateX"})
in
    #"Removed Columns";

shared EffortAllMatrixAG1_1D1 = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\2. Calculations\E-O-I\Effort.xlsx"), null, true),
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text}, {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

shared Shifts = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\2. Calculations\SETTINGS.xlsx"), null, true),
    Shifts_Table = Source{[Item="Shifts",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Shifts_Table,{{"Shifts", type text}}),
    #"Added Index" = Table.AddIndexColumn(#"Changed Type", "Index", 1, 1, Int64.Type)
in
    #"Added Index";

shared MAIT_date_adjusted_for_example = let
    Source = Excel.CurrentWorkbook(){[Name="MAIT_date_adjusted_for_example"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", Int64.Type}}),
    Column1 = #"Changed Type"{0}[Column1]
in
    Column1;

shared ShiftDuration = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\MAIT\2. Calculations\SETTINGS.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared EffortTypeOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcknNTcxLUdJRMlSK1YlWcszJyU9OLMnMzwMKGYGFPPMySzITcxRcHBXcEwuAwsZKsbEA", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [EffortType = _t, Index = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"EffortType", type text}, {"Index", Int64.Type}})
in
    #"Changed Type";

shared FACILITY = let
    Source = Excel.CurrentWorkbook(){[Name="Table9"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", type text}}),
    Column2 = #"Changed Type"{0}[Column2]
in
    Column2;

shared Table_Master_Avail_Alloc_ResidualMAIT = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"), null, true),
    Table_Master_Avail_Alloc_Residual_Table = Source{[Item="Table_Master_Avail_Alloc_Residual",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_Master_Avail_Alloc_Residual_Table,{{"Name", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text}, {"Status", type text}, {"AllocationHRS", type number}, {"ResidualHRS", type number}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}})
in
    #"Sorted Rows";

shared DateADJUSTEDShiftRole = let
    Source = PermutationDimensions,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date", "Shifts"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Added ROLES" = Table.AddColumn(#"Removed Duplicates", "Roles", each ROLES),
    #"Renamed Columns" = Table.RenameColumns(#"Added ROLES",{{"Date", "DateX"}}),
    #"Expanded Roles" = Table.ExpandTableColumn(#"Renamed Columns", "Roles", {"Role"}, {"Role"}),
    #"Add DAYS" = Table.AddColumn(#"Expanded Roles", "Date", each Date.AddDays([DateX],MAIT_date_adjusted_for_example)),
    #"Renamed Columns1" = Table.RenameColumns(#"Add DAYS",{{"Shifts", "Shift"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns1",{{"Date", type date}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Shift", Order.Ascending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Role"}),
    #"Removed Duplicates1" = Table.Distinct(#"Removed Columns")
in
    #"Removed Duplicates1";

shared #"Residual-RosteredLIST" = let
    Source = DateADJUSTEDShiftRole,
    #"Merged Queries2" = Table.NestedJoin(Source, {"DateX", "Shift"}, Table_Master_Avail_Alloc_ResidualMAIT, {"Date", "Shift"}, "Table_Master_Avail_Alloc_ResidualMAIT", JoinKind.LeftOuter),
    #"Expanded Table_Master_Avail_Alloc_ResidualMAIT" = Table.ExpandTableColumn(#"Merged Queries2", "Table_Master_Avail_Alloc_ResidualMAIT", {"Name", "Role", "Status", "AllocationHRS", "ResidualHRS", "Facility"}, {"Name", "Role", "Status", "AllocationHRS", "ResidualHRS", "Facility"}),
    #"Filled Down" = Table.FillDown(#"Expanded Table_Master_Avail_Alloc_ResidualMAIT",{"Name", "Role", "Facility"}),
    #"Replaced Value" = Table.ReplaceValue(#"Filled Down",null,"Unavailable/Unallocated",Replacer.ReplaceValue,{"Status"}),
    BUFFER1 = Table.Buffer(#"Replaced Value"),
    #"Filtered FACILITY" = Table.SelectRows(BUFFER1, each Text.Contains([Facility], FACILITY)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered FACILITY",{"DateX"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"AllocationHRS", type text}, {"ResidualHRS", type text}}),
    #"Combine RESIDUAL/ALLOC" = Table.AddColumn(#"Changed Type", "Residual/Allocation", each [ResidualHRS] & "/ " & [#"AllocationHRS"]),
    #"Merged Queries" = Table.NestedJoin(#"Combine RESIDUAL/ALLOC", {"Shift"}, Shifts, {"Shifts"}, "Shifts", JoinKind.LeftOuter),
    #"Expanded Shifts1" = Table.ExpandTableColumn(#"Merged Queries", "Shifts", {"Index"}, {"Index"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Shifts1",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"Index", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Index"}),
    BUFFER = Table.Buffer(#"Removed Columns1"),
    #"Changed Type1" = Table.TransformColumnTypes(BUFFER,{{"AllocationHRS", type number}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Changed Type1",{"ResidualHRS"})
in
    #"Removed Columns2";

shared ResRosteredHRS = let
    Source = #"Residual-RosteredLIST",
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Name", "Role"}, {{"Rostered HRS", each List.Sum([AllocationHRS]), type nullable text}})
in
    #"Grouped Rows";

shared #"Residual-RosteredMATRIX" = let
    Source = #"Residual-RosteredLIST",
    #"Duplicated Column" = Table.DuplicateColumn(Source, "Date", "Date - Copy"),
    #"Changed Type" = Table.TransformColumnTypes(#"Duplicated Column",{{"Date - Copy", type text}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Changed Type", {{"Date - Copy", each Text.BeforeDelimiter(_, "/", {0, RelativePosition.FromEnd}), type text}}),
    #"Merged Columns" = Table.CombineColumns(#"Extracted Text Before Delimiter",{"Date - Copy", "Shift"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"DateShift"),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged Columns",{"Date", "AllocationHRS", "Status"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns1", List.Distinct(#"Removed Columns1"[DateShift]), "DateShift", "Residual/Allocation"),
    #"Merged Queries" = Table.NestedJoin(#"Pivoted Column", {"Facility", "Name", "Role"}, ResRosteredHRS, {"Facility", "Name", "Role"}, "ResRosteredHRS", JoinKind.LeftOuter),
    #"Expanded ResRosteredHRS" = Table.ExpandTableColumn(#"Merged Queries", "ResRosteredHRS", {"Rostered HRS"}, {"Rostered HRS"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded ResRosteredHRS",{"Name", "Role", "Rostered HRS", "Facility", "19/06 AM", "19/06 PM", "19/06 NIGHT", "20/06 AM", "20/06 PM", "20/06 NIGHT", "21/06 AM", "21/06 PM", "21/06 NIGHT", "22/06 AM", "22/06 PM", "22/06 NIGHT", "23/06 AM", "23/06 PM", "23/06 NIGHT", "24/06 AM", "24/06 PM", "24/06 NIGHT", "25/06 AM", "25/06 PM", "25/06 NIGHT", "26/06 AM", "26/06 PM", "26/06 NIGHT", "27/06 AM", "27/06 PM", "27/06 NIGHT", "28/06 AM", "28/06 PM", "28/06 NIGHT", "29/06 AM", "29/06 PM", "29/06 NIGHT", "30/06 AM", "30/06 PM", "30/06 NIGHT", "1/07 AM", "1/07 PM", "1/07 NIGHT", "2/07 AM", "2/07 PM", "2/07 NIGHT"})
in
    #"Reordered Columns";

shared Header = let
    Source = EffortAllMatrixAG1_1D1,
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Facility], FACILITY)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{ "Capacity", "CapacityX", "CapacityMaxHC", "DATESHIFT"}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Removed Columns", "Date", "Date - Copy"),
    #"Changed Type1" = Table.TransformColumnTypes(#"Duplicated Column",{{"Date - Copy", type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Changed Type1",{"Facility", "Role", "Date", "Shift", "Demand", "Allocation", "Date - Copy"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns1",{{"Date - Copy", type text}}),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Changed Type", "Shift", "Shift - Copy"),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Duplicated Column1", {{"Date - Copy", each Text.BeforeDelimiter(_, "/", {0, RelativePosition.FromEnd}), type text}}),
    #"Merged Queries" = Table.NestedJoin(#"Extracted Text Before Delimiter", {"Shift"}, Shifts, {"Shifts"}, "Shifts", JoinKind.LeftOuter),
    #"Expanded Shifts" = Table.ExpandTableColumn(#"Merged Queries", "Shifts", {"Index"}, {"Shifts.Index"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Shifts",{{"Facility", Order.Ascending}, {"Date", Order.Ascending},{"Role", Order.Ascending}, {"Shifts.Index", Order.Ascending}}),
    #"Merged DATESHIFT" = Table.CombineColumns(#"Sorted Rows",{"Date - Copy", "Shift - Copy"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Date | Shift"),
    #"Multiplied Column" = Table.TransformColumns(#"Merged DATESHIFT", {{"Demand", each _ * ShiftDuration, type number}}),
    #"Multiplied Column1" = Table.TransformColumns(#"Multiplied Column", {{"Allocation", each _ * ShiftDuration, type number}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Multiplied Column1", "Subtraction", each  [Allocation] -[Demand], type number),
    #"Rounded Off1" = Table.TransformColumns(#"Inserted Subtraction",{{"Demand", each Number.Round(_, 1), type number}, {"Allocation", each Number.Round(_, 1), type number}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Rounded Off1",{{"Subtraction", "Initial DA Gap"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns1",{"Shifts.Index", "Date", "Shift"}),
    #"Unpivoted EFFORTS" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Role", "Date | Shift","Facility"}, "Attribute", "Value"),
    #"Rounded Off" = Table.TransformColumns(#"Unpivoted EFFORTS",{{"Value", each Number.Round(_, 1), type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Rounded Off",{{"Attribute", "EffortType"}, {"Value", "Effort"}}),
    #"Pivoted Column" = Table.Pivot(#"Renamed Columns", List.Distinct(#"Renamed Columns"[#"Date | Shift"]), "Date | Shift", "Effort", List.Sum),
    #"Reordered Columns" = Table.ReorderColumns(#"Pivoted Column",{"Facility", "EffortType", "Role", "19/06 AM", "19/06 PM", "19/06 NIGHT", "20/06 AM", "20/06 PM", "20/06 NIGHT", "21/06 AM", "21/06 PM", "21/06 NIGHT", "22/06 AM", "22/06 PM", "22/06 NIGHT", "23/06 AM", "23/06 PM", "23/06 NIGHT", "24/06 AM", "24/06 PM", "24/06 NIGHT", "25/06 AM", "25/06 PM", "25/06 NIGHT", "26/06 AM", "26/06 PM", "26/06 NIGHT", "27/06 AM", "27/06 PM", "27/06 NIGHT", "28/06 AM", "28/06 PM", "28/06 NIGHT", "29/06 AM", "29/06 PM", "29/06 NIGHT", "30/06 AM", "30/06 PM", "30/06 NIGHT", "1/07 AM", "1/07 PM", "1/07 NIGHT", "2/07 AM", "2/07 PM", "2/07 NIGHT"}),
    #"Sorted Rows1" = Table.Sort(#"Reordered Columns",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"EffortType", Order.Descending}}),
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows1", {"EffortType"}, EffortTypeOrder, {"EffortType"}, "EffortTypeOrder", JoinKind.LeftOuter),
    #"Expanded EffortTypeOrder" = Table.ExpandTableColumn(#"Merged Queries1", "EffortTypeOrder", {"Index"}, {"EffortTypeOrder.Index"}),
    #"Sorted Rows2" = Table.Sort(#"Expanded EffortTypeOrder",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"EffortTypeOrder.Index", Order.Ascending}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Sorted Rows2",{"EffortTypeOrder.Index"})
in
    #"Removed Columns2";

shared PermutationDimensions = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\Centri\2. Accounts - 2.0 RFBI\04.1 PoC\04 Analysis\230714\MAIT\2. Calculations\SETTINGS.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Shifts", type text}, {"Unit", type text}, {"RolesList", type text}})
in
    #"Changed Type";

shared ROLES = let
    Source = Table_Master_Avail_Alloc_ResidualMAIT,
    #"Grouped Rows" = Table.Group(Source, {"Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"})
in
    #"Removed Columns";

shared #"Residual-RosteredMATRIX-RN" = let
    Source = #"Residual-RosteredMATRIX",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Role] = "Registered Nurse"))
in
    #"Filtered Rows";

shared #"Residual-RosteredMATRIX-CARER" = let
    Source = #"Residual-RosteredMATRIX",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Role] = "Carer"))
in
    #"Filtered Rows";

shared #"Header-RN" = let
    Source = Header,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Role] = "Registered Nurse"))
in
    #"Filtered Rows";

shared #"Header-CARER" = let
    Source = Header,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Role] = "Carer"))
in
    #"Filtered Rows";

shared #"DateADJUSTEDShiftRole (2)" = let
    Source = PermutationDimensions,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date", "Shifts"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns"),
    #"Added ROLES" = Table.AddColumn(#"Removed Duplicates", "Roles", each ROLES),
    #"Renamed Columns" = Table.RenameColumns(#"Added ROLES",{{"Date", "DateX"}}),
    #"Expanded Roles" = Table.ExpandTableColumn(#"Renamed Columns", "Roles", {"Role"}, {"Role"}),
    #"Add DAYS" = Table.AddColumn(#"Expanded Roles", "Date", each Date.AddDays([DateX],MAIT_date_adjusted_for_example)),
    #"Renamed Columns1" = Table.RenameColumns(#"Add DAYS",{{"Shifts", "Shift"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns1",{{"Date", type date}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Shift", Order.Ascending}})
in
    #"Sorted Rows";