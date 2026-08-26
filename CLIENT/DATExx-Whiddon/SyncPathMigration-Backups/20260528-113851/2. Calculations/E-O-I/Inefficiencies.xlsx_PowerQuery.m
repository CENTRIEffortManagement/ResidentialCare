// Power Query from: Inefficiencies.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\E-O-I\Inefficiencies.xlsx
// Extracted: 2026-05-21T00:47:31.450Z

section Section1;

shared Folder = let
    Source = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Folder", type text}}),
    Folder1 = #"Changed Type"{0}[Folder]
in
    Folder1;

shared #"EffortOutcomesAG1_1DayShiftIMPORT !!" = let
        Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\E-O-I\EffortOutcomes.xlsx"), null, true),
    EffortOutcomesAG1_1DayShift_Table = Source{[Item="EffortOutcomesAG1_1DayShift",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortOutcomesAG1_1DayShift_Table,{{"Date", type date}, {"Role", type text}, {"Shift", type text}, {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"Allocation", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}, {"Epn", type number}, {"EpX", type number}, {"Ein", type number}, {"Ipn", type number}, {"Ipx", type number}, {"Iin", type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Role", "L1.1"}, {"Shift", "L1.2"}}),
    #"Added Version X !!" = Table.AddColumn(#"Renamed Columns", "Version", each "X"),
    #"Filtered DATE !!" = Table.SelectRows(#"Added Version X !!", each ([Date] <> #date(2023, 6, 18)))
in
    #"Filtered DATE !!";

shared EffortOutcomesAG1_1DayShiftAB = let
    Source = #"EffortOutcomesAG1_1DayShiftIMPORT !!",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Demand", "D"}, {"Capacity", "Cs"}, {"CapacityX", "CX"}, {"Allocation", "A"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Epn", "EpX", "Ein", "Ipn", "Ipx", "Iin"})
in
    #"Removed Columns";

shared EffortOutcomesAG1_1DayShiftEF = let
    Source = #"EffortOutcomesAG1_1DayShiftIMPORT !!",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Demand", "D"}, {"Capacity", "Cs"}, {"CapacityX", "CX"}, {"Allocation", "A"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Apn", "Apx", "Ai", "Ipn", "Ipx", "Iin"})
in
    #"Removed Columns";

shared EffortOutcomesAG1_1DayShiftIN = let
    Source = #"EffortOutcomesAG1_1DayShiftIMPORT !!",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Demand", "D"}, {"Capacity", "Cs"}, {"CapacityX", "CX"}, {"Allocation", "A"}})
in
    #"Renamed Columns";

shared InefficienciesAG1_1DayShiftAB = let
    Source = 
          // Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content]
          
            Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content]
            ,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Level", type text}, {"Potential", Int64.Type}, {"Planned", Int64.Type}, {"Degree", Int64.Type}, {"Spare", Int64.Type}, {"Spare Allocated", Int64.Type}, {"Block", Int64.Type}, {"EXCESS OVER-ALLOCATION", type number}, {"SPARE STRETCH", type number}, {"SPARE SLACK", type number}, {"SPARE STRETCH ALLOCATED", type number}, {"SPARE SLACK ALLOCATED", type number}, {"EXCESS STRETCH (ALLOCATED)", type number}, {"EXCESS ALLOCATED SLACK", type number}, {"EXCESS STRETCH", type number}, {"EXCESS SLACK", type number}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"POTENTIAL SHORTFALL", type number}, {"WASTED STRETCH", type number}, {"WASTED SLACK", type number}, {"ALLOCATED STRETCH", type number}, {"ALLOCATED SLACK", type number}, {"UNALLOCATED SLACK", type number}, {"LATENT", Int64.Type}, {"LATENT.ALLOCATED", type number}, {"LATENT.OVERALLOCATED", type number}, {"AB Potential", type number}, {"Check", type number}, {"wAB EXCESS OVER-ALLOCATION", type number}, {"wAB SPARE STRETCH", type number}, {"wAB SPARE SLACK", type number}, {"wAB SPARE STRETCH ALLOCATED", type number}, {"wAB SPARE SLACK ALLOCATED", type number}, {"wAB EXCESS STRETCH (ALLOCATED)", type number}, {"wAB EXCESS ALLOCATED SLACK", type number}, {"wAB EXCESS STRETCH", type number}, {"wAB EXCESS SLACK", type number}, {"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"wAB POTENTIAL SHORTFALL", type number}, {"wAB WASTED STRETCH", type number}, {"wAB WASTED SLACK", type number}, {"wAB ALLOCATED STRETCH", type number}, {"wAB ALLOCATED SLACK", type number}, {"wAB UNALLOCATED SLACK", type number}, {"wAB LATENT", Int64.Type}, {"wAB LATENT.ALLOCATED", Int64.Type}, {"wAB LATENT.OVERALLOCATED", Int64.Type}, {"wPotential", Int64.Type}, {"Column1", type any}, {"Column2", type any}, {"Date", type datetime}, {"L1.1", type text}, {"L1.2", type text}, {"D", type number}, {"Cs", type number}, {"CX", type number}, {"A", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2", "D", "Cs", "CX", "A",  "Apn", "Apx", "Ai"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Date", "L1.2", "L1.1",  "EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns",{{"Date", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"EXCESS OVER-ALLOCATION", "AB EXCESS OVER-ALLOCATION"}, {"SPARE STRETCH", "AB SPARE STRETCH"}, {"SPARE SLACK", "AB SPARE SLACK"}, {"SPARE STRETCH ALLOCATED", "AB SPARE STRETCH ALLOCATED"}, {"SPARE SLACK ALLOCATED", "AB SPARE SLACK ALLOCATED"}, {"EXCESS STRETCH (ALLOCATED)", "AB EXCESS STRETCH (ALLOCATED)"}, {"EXCESS ALLOCATED SLACK", "AB EXCESS ALLOCATED SLACK"}, {"EXCESS STRETCH", "AB EXCESS STRETCH"}, {"EXCESS SLACK", "AB EXCESS SLACK"}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)"}, {"POTENTIAL SHORTFALL", "AB POTENTIAL SHORTFALL"}, {"WASTED STRETCH", "AB WASTED STRETCH"}, {"WASTED SLACK", "AB WASTED SLACK"}, {"ALLOCATED STRETCH", "AB ALLOCATED STRETCH"}, {"ALLOCATED SLACK", "AB ALLOCATED SLACK"}, {"UNALLOCATED SLACK", "AB UNALLOCATED SLACK"}, {"LATENT", "AB LATENT"}, {"LATENT.ALLOCATED", "AB LATENT.ALLOCATED"}, {"LATENT.OVERALLOCATED", "AB  LATENT.OVERALLOCATED"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns1",{"Date", "L1.1", "L1.2",  "AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "Potential"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns1",{{"Potential", "AB Potential"}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns2",{"Facility", "Date", "L1.1", "L1.2", "AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential", "CapacityMaxHC", "DATESHIFT"})
in
    #"Reordered Columns2";

shared InefficienciesAG1_1DayShiftEF = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftEF"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Level", type text}, {"Potential", Int64.Type}, {"Planned", Int64.Type}, {"Degree", Int64.Type}, {"Spare", Int64.Type}, {"Spare Allocated", Int64.Type}, {"Block", Int64.Type}, {"EXCESS OVER-ALLOCATION", type number}, {"SPARE STRETCH", type number}, {"SPARE SLACK", type number}, {"SPARE STRETCH ALLOCATED", type number}, {"SPARE SLACK ALLOCATED", type number}, {"EXCESS STRETCH (ALLOCATED)", type number}, {"EXCESS ALLOCATED SLACK", type number}, {"EXCESS STRETCH", type number}, {"EXCESS SLACK", type number}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"POTENTIAL SHORTFALL", type number}, {"WASTED STRETCH", type number}, {"WASTED SLACK", type number}, {"ALLOCATED STRETCH", type number}, {"ALLOCATED SLACK", type number}, {"UNALLOCATED SLACK", Int64.Type}, {"LATENT", Int64.Type}, {"LATENT.ALLOCATED", Int64.Type}, {"LATENT.OVERALLOCATED", Int64.Type}, {"Potential2", Int64.Type}, {"Check", type number}, {"wAB EXCESS OVER-ALLOCATION", type number}, {"wAB SPARE STRETCH", type number}, {"wAB SPARE SLACK", type number}, {"wAB SPARE STRETCH ALLOCATED", type number}, {"wAB SPARE SLACK ALLOCATED", type number}, {"wAB EXCESS STRETCH (ALLOCATED)", type number}, {"wAB EXCESS ALLOCATED SLACK", type number}, {"wAB EXCESS STRETCH", type number}, {"wAB EXCESS SLACK", type number}, {"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"wAB POTENTIAL SHORTFALL", type number}, {"wAB WASTED STRETCH", type number}, {"wAB WASTED SLACK", type number}, {"wAB ALLOCATED STRETCH", type number}, {"wAB ALLOCATED SLACK", type number}, {"wAB UNALLOCATED SLACK", type number}, {"wAB LATENT", Int64.Type}, {"wAB LATENT.ALLOCATED", Int64.Type}, {"wAB LATENT.OVERALLOCATED", Int64.Type}, {"wPotential", Int64.Type}, {"Column1", type number}, {"Column2", type any}, {"Date", type datetime}, {"L1.1", type text}, {"L1.2", type text}, {"D", type number}, {"Cs", type number}, {"CX", type number}, {"A", type number}, {"Epn", type number}, {"EpX", type number}, {"Ein", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2", "D", "Cs", "CX", "A"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Epn", "EpX", "Ein", "L1.2", "L1.1", "Date", "Potential2", "LATENT.OVERALLOCATED", "LATENT.ALLOCATED", "LATENT", "UNALLOCATED SLACK", "ALLOCATED SLACK", "ALLOCATED STRETCH", "WASTED SLACK", "WASTED STRETCH", "POTENTIAL SHORTFALL", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EXCESS SLACK", "EXCESS STRETCH", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH (ALLOCATED)", "SPARE SLACK ALLOCATED", "SPARE STRETCH ALLOCATED", "SPARE SLACK", "SPARE STRETCH", "EXCESS OVER-ALLOCATION"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Potential2", "Potential"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns",{"Date", "L1.1", "L1.2", "Epn", "EpX", "Ein", "EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns1",{{"Date", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"EXCESS OVER-ALLOCATION", "EF EXCESS OVER-ALLOCATION"}, {"SPARE STRETCH", "EF SPARE STRETCH"}, {"SPARE SLACK", "EF SPARE SLACK"}, {"SPARE STRETCH ALLOCATED", "EF SPARE STRETCH ALLOCATED"}, {"SPARE SLACK ALLOCATED", "EF SPARE SLACK ALLOCATED"}, {"EXCESS STRETCH (ALLOCATED)", "EF EXCESS STRETCH (ALLOCATED)"}, {"EXCESS ALLOCATED SLACK", "EF EXCESS ALLOCATED SLACK"}, {"EXCESS STRETCH", "EF EXCESS STRETCH"}, {"EXCESS SLACK", "EF EXCESS SLACK"}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)"}, {"POTENTIAL SHORTFALL", "EF POTENTIAL SHORTFALL"}, {"WASTED STRETCH", "EF WASTED STRETCH"}, {"WASTED SLACK", "EF WASTED SLACK"}, {"ALLOCATED STRETCH", "EF ALLOCATED STRETCH"}, {"ALLOCATED SLACK", "EF ALLOCATED SLACK"}, {"UNALLOCATED SLACK", "EF UNALLOCATED SLACK"}, {"LATENT", "EF LATENT"}, {"LATENT.ALLOCATED", "EF LATENT.ALLOCATED"}, {"LATENT.OVERALLOCATED", "EF LATENT.OVERALLOCATED"}, {"Potential", "EF Potential"}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns1",{"Facility", "Date", "L1.1", "L1.2", "Epn", "EpX", "Ein", "EF EXCESS OVER-ALLOCATION", "EF SPARE STRETCH", "EF SPARE SLACK", "EF SPARE STRETCH ALLOCATED", "EF SPARE SLACK ALLOCATED", "EF EXCESS STRETCH (ALLOCATED)", "EF EXCESS ALLOCATED SLACK", "EF EXCESS STRETCH", "EF EXCESS SLACK", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL", "EF WASTED STRETCH", "EF WASTED SLACK", "EF ALLOCATED STRETCH", "EF ALLOCATED SLACK", "EF UNALLOCATED SLACK", "EF LATENT", "CapacityMaxHC", "DATESHIFT", "EF LATENT.ALLOCATED", "EF LATENT.OVERALLOCATED", "EF Potential"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Reordered Columns2",{{"EF LATENT", type number}, {"EF UNALLOCATED SLACK", type number}, {"CapacityMaxHC", type number}, {"EF LATENT.ALLOCATED", type number}, {"EF LATENT.OVERALLOCATED", type number}, {"EF Potential", type number}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type2", {{"EF LATENT", 0}})
in
    #"Replaced Errors";

shared InefficienciesAG1_1DayShiftIN = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftIN"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2", "D", "Cs", "CX", "A", "Apn", "Apx", "Ai", "Epn", "EpX", "Ein"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Potential2", "Potential"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Date", "L1.1", "L1.2", "Ipn", "Ipx", "Iin", "EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Date", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"EXCESS OVER-ALLOCATION", "IN EXCESS OVER-ALLOCATION"}, {"SPARE STRETCH", "IN SPARE STRETCH"}, {"SPARE SLACK", "IN SPARE SLACK"}, {"SPARE STRETCH ALLOCATED", "IN SPARE STRETCH ALLOCATED"}, {"SPARE SLACK ALLOCATED", "IN SPARE SLACK ALLOCATED"}, {"EXCESS STRETCH (ALLOCATED)", "IN EXCESS STRETCH (ALLOCATED)"}, {"EXCESS ALLOCATED SLACK", "IN EXCESS ALLOCATED SLACK"}, {"EXCESS STRETCH", "IN EXCESS STRETCH"}, {"EXCESS SLACK", "IN EXCESS SLACK"}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)"}, {"POTENTIAL SHORTFALL", "IN POTENTIAL SHORTFALL"}, {"WASTED STRETCH", "IN WASTED STRETCH"}, {"WASTED SLACK", "IN WASTED SLACK"}, {"ALLOCATED STRETCH", "IN ALLOCATED STRETCH"}, {"ALLOCATED SLACK", "IN ALLOCATED SLACK"}, {"UNALLOCATED SLACK", "IN UNALLOCATED SLACK"}, {"LATENT", "IN LATENT"}, {"LATENT.ALLOCATED", "IN LATENT.ALLOCATED"}, {"LATENT.OVERALLOCATED", "IN LATENT.OVERALLOCATED"}, {"Potential", "IN Potential"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns1",{"Facility", "Date", "L1.1", "L1.2", "Ipn", "Ipx", "Iin", "IN EXCESS OVER-ALLOCATION", "IN SPARE STRETCH", "IN SPARE SLACK", "IN SPARE STRETCH ALLOCATED", "IN SPARE SLACK ALLOCATED", "IN EXCESS STRETCH (ALLOCATED)", "IN EXCESS ALLOCATED SLACK", "IN EXCESS STRETCH", "IN EXCESS SLACK", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL", "IN WASTED STRETCH", "IN WASTED SLACK", "IN ALLOCATED STRETCH", "IN ALLOCATED SLACK", "IN UNALLOCATED SLACK", "IN LATENT", "CapacityMaxHC", "DATESHIFT", "IN LATENT.ALLOCATED", "IN LATENT.OVERALLOCATED", "IN Potential"})
in
    #"Reordered Columns1";

shared #"InefficienciesAG1_1DayShiftALL Prep" = let
    Source = EffortOutcomesAG1_1DayShiftAB,
    #"Sorted Rows" = Table.Sort(Source,{{"L1.1", Order.Ascending}, {"Date", Order.Ascending}, {"L1.2", Order.Ascending}, {"Version", Order.Ascending}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Sorted Rows",{"Version", "Facility", "L1.1", "Date", "L1.2", "D", "Cs", "CX", "CapacityMaxHC", "A", "DATESHIFT", "Apn", "Apx", "Ai"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Date", "L1.1", "L1.2", "Facility", "Version"}, InefficienciesAG1_1DayShiftAB, {"Date", "L1.1", "L1.2", "Facility", "Version"}, "InefficienciesAG1_1DayShiftAB", JoinKind.LeftOuter),
    #"Expanded InefficienciesAG1_1DayShiftAB1" = Table.ExpandTableColumn(#"Merged Queries2", "InefficienciesAG1_1DayShiftAB", {"AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential"}, {"AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded InefficienciesAG1_1DayShiftAB1", {"Date", "L1.1", "L1.2", "Facility", "Version"}, InefficienciesAG1_1DayShiftEF, {"Date", "L1.1", "L1.2", "Facility", "Version"}, "InefficienciesAG1_1DayShiftAB", JoinKind.LeftOuter),
    #"Expanded InefficienciesAG1_1DayShiftAB" = Table.ExpandTableColumn(#"Merged Queries", "InefficienciesAG1_1DayShiftAB", {"Epn", "EpX", "Ein", "EF EXCESS OVER-ALLOCATION", "EF SPARE STRETCH", "EF SPARE SLACK", "EF SPARE STRETCH ALLOCATED", "EF SPARE SLACK ALLOCATED", "EF EXCESS STRETCH (ALLOCATED)", "EF EXCESS ALLOCATED SLACK", "EF EXCESS STRETCH", "EF EXCESS SLACK", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL", "EF WASTED STRETCH", "EF WASTED SLACK", "EF ALLOCATED STRETCH", "EF ALLOCATED SLACK", "EF UNALLOCATED SLACK", "EF LATENT", "EF LATENT.ALLOCATED", "EF LATENT.OVERALLOCATED", "EF Potential"}, {"Epn", "EpX", "Ein", "EF EXCESS OVER-ALLOCATION", "EF SPARE STRETCH", "EF SPARE SLACK", "EF SPARE STRETCH ALLOCATED", "EF SPARE SLACK ALLOCATED", "EF EXCESS STRETCH (ALLOCATED)", "EF EXCESS ALLOCATED SLACK", "EF EXCESS STRETCH", "EF EXCESS SLACK", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL", "EF WASTED STRETCH", "EF WASTED SLACK", "EF ALLOCATED STRETCH", "EF ALLOCATED SLACK", "EF UNALLOCATED SLACK", "EF LATENT", "EF LATENT.ALLOCATED", "EF LATENT.OVERALLOCATED", "EF Potential"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded InefficienciesAG1_1DayShiftAB", {"Date", "L1.1", "L1.2", "Facility", "Version"}, InefficienciesAG1_1DayShiftIN, {"Date", "L1.1", "L1.2", "Facility", "Version"}, "InefficienciesAG1_1DayShiftIN", JoinKind.LeftOuter),
    #"Expanded InefficienciesAG1_1DayShiftIN" = Table.ExpandTableColumn(#"Merged Queries1", "InefficienciesAG1_1DayShiftIN", {"Ipn", "Ipx", "Iin", "IN EXCESS OVER-ALLOCATION", "IN SPARE STRETCH", "IN SPARE SLACK", "IN SPARE STRETCH ALLOCATED", "IN SPARE SLACK ALLOCATED", "IN EXCESS STRETCH (ALLOCATED)", "IN EXCESS ALLOCATED SLACK", "IN EXCESS STRETCH", "IN EXCESS SLACK", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL", "IN WASTED STRETCH", "IN WASTED SLACK", "IN ALLOCATED STRETCH", "IN ALLOCATED SLACK", "IN UNALLOCATED SLACK", "IN LATENT", "IN LATENT.ALLOCATED", "IN LATENT.OVERALLOCATED", "IN Potential"}, {"Ipn", "Ipx", "Iin", "IN EXCESS OVER-ALLOCATION", "IN SPARE STRETCH", "IN SPARE SLACK", "IN SPARE STRETCH ALLOCATED", "IN SPARE SLACK ALLOCATED", "IN EXCESS STRETCH (ALLOCATED)", "IN EXCESS ALLOCATED SLACK", "IN EXCESS STRETCH", "IN EXCESS SLACK", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL", "IN WASTED STRETCH", "IN WASTED SLACK", "IN ALLOCATED STRETCH", "IN ALLOCATED SLACK", "IN UNALLOCATED SLACK", "IN LATENT", "IN LATENT.ALLOCATED", "IN LATENT.OVERALLOCATED", "IN Potential"}),
    #"Added Custom" = Table.AddColumn(#"Expanded InefficienciesAG1_1DayShiftIN", "Level", each "3-Pool"),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Version", "Facility", "Level", "Date", "L1.1", "L1.2", "D", "Cs", "CapacityMaxHC", "CX", "DATESHIFT", "A", "Apn", "Apx", "Ai", "AB EXCESS OVER-ALLOCATION", "AB SPARE STRETCH", "AB SPARE SLACK", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential", "Epn", "EpX", "Ein", "EF EXCESS OVER-ALLOCATION", "EF SPARE STRETCH", "EF SPARE SLACK", "EF SPARE STRETCH ALLOCATED", "EF SPARE SLACK ALLOCATED", "EF EXCESS STRETCH (ALLOCATED)", "EF EXCESS ALLOCATED SLACK", "EF EXCESS STRETCH", "EF EXCESS SLACK", "EF POTENTIAL SHORTFALL (OVER-ALLOCATED)", "EF POTENTIAL SHORTFALL", "EF WASTED STRETCH", "EF WASTED SLACK", "EF ALLOCATED STRETCH", "EF ALLOCATED SLACK", "EF UNALLOCATED SLACK", "EF LATENT", "EF LATENT.ALLOCATED", "EF LATENT.OVERALLOCATED", "EF Potential", "Ipn", "Ipx", "Iin", "IN EXCESS OVER-ALLOCATION", "IN SPARE STRETCH", "IN SPARE SLACK", "IN SPARE STRETCH ALLOCATED", "IN SPARE SLACK ALLOCATED", "IN EXCESS STRETCH (ALLOCATED)", "IN EXCESS ALLOCATED SLACK", "IN EXCESS STRETCH", "IN EXCESS SLACK", "IN POTENTIAL SHORTFALL (OVER-ALLOCATED)", "IN POTENTIAL SHORTFALL", "IN WASTED STRETCH", "IN WASTED SLACK", "IN ALLOCATED STRETCH", "IN ALLOCATED SLACK", "IN UNALLOCATED SLACK", "IN LATENT", "IN LATENT.ALLOCATED", "IN LATENT.OVERALLOCATED", "IN Potential"})
in
    #"Reordered Columns";

shared InefficienciesAG1_1Level3210 = let
    Source = #"InefficienciesAG1_1DayShiftALL Prep",
    #"Appended Query" = Table.Combine({Source, InefficienciesAG1_1L2DayRole}),
    #"Filtered Rows" = Table.SelectRows(#"Appended Query", each true),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"L1.1", "Role"}, {"L1.2", "Shift"}}),
    #"Appended Query1" = Table.Combine({#"Renamed Columns", InefficienciesAG1_1L2Day, InefficienciesAG1_1L2DayL0Org, InefficienciesAG1_1L2DayL0Facility, InefficienciesAG1_1L2FacilityDay})
in
    #"Appended Query1";

shared InefficienciesAG1_1L2DayRole = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check"  }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"Column1", "Column2"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential"}),
    #"Grouped Rows" = Table.Group(#"Reordered Columns", {"Date", "L1.1", "Facility", "Version"}, {{"L2.1sw EXCESS OVER-ALLOCATION", each List.Sum([#"wAB EXCESS OVER-ALLOCATION"]), type number}, {"L2.1sw SPARE STRETCH", each List.Sum([wAB SPARE STRETCH]), type number}, {"L2.1sw SPARE SLACK", each List.Sum([wAB SPARE SLACK]), type number}, {"L2.1sw SPARE STRETCH ALLOCATED", each List.Sum([wAB SPARE STRETCH ALLOCATED]), type number}, {"L2.1sw SPARE SLACK ALLOCATED", each List.Sum([wAB SPARE SLACK ALLOCATED]), type number}, {"L2.1sw EXCESS STRETCH (ALLOCATED)", each List.Sum([#"wAB EXCESS STRETCH (ALLOCATED)"]), type number}, {"L2.1sw EXCESS ALLOCATED SLACK", each List.Sum([wAB EXCESS ALLOCATED SLACK]), type number}, {"L2.1sw EXCESS STRETCH", each List.Sum([wAB EXCESS STRETCH]), type number}, {"L2.1sw EXCESS SLACK", each List.Sum([wAB EXCESS SLACK]), type number}, {"L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", each List.Sum([#"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)"]), type number}, {"L2.1sw POTENTIAL SHORTFALL", each List.Sum([wAB POTENTIAL SHORTFALL]), type number}, {"L2.1sw WASTED STRETCH", each List.Sum([wAB WASTED STRETCH]), type number}, {"L2.1sw WASTED SLACK", each List.Sum([wAB WASTED SLACK]), type number}, {"L2.1sw ALLOCATED SLACK", each List.Sum([wAB ALLOCATED SLACK]), type number}, {"L2.1sw ALLOCATED STRETCH", each List.Sum([wAB ALLOCATED STRETCH]), type number}, {"L2.1sw UNALLOCATED SLACK", each List.Sum([wAB UNALLOCATED SLACK]), type number}, {"L2.1sw LATENT", each List.Sum([wAB LATENT]), type number}, {"L2.1sw LATENT.ALLOCATED", each List.Sum([wAB LATENT.ALLOCATED]), type number}, {"L2.1sw  LATENT.OVERALLOCATED", each List.Sum([wAB LATENT.OVERALLOCATED]), type number}, {"L2.1sw Potential", each List.Sum([wPotential]), type number}, {"D", each List.Sum([D]), type number}, {"Cs", each List.Sum([Cs]), type number}, {"CX", each List.Sum([CX]), type number}, {"A", each List.Sum([A]), type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{"Date", "L1.1", "D", "Cs", "CX", "A", "L2.1sw EXCESS OVER-ALLOCATION", "L2.1sw SPARE STRETCH",  "L2.1sw SPARE STRETCH ALLOCATED", "L2.1sw SPARE SLACK ALLOCATED", "L2.1sw EXCESS STRETCH (ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL", "L2.1sw WASTED STRETCH", "L2.1sw WASTED SLACK", "L2.1sw ALLOCATED STRETCH", "L2.1sw ALLOCATED SLACK", "L2.1sw UNALLOCATED SLACK", "L2.1sw LATENT", "L2.1sw LATENT.ALLOCATED", "L2.1sw  LATENT.OVERALLOCATED"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Reordered Columns1", {"Date", "L1.1",  "D", "Cs",  "CX", "A","Facility","Version"}, "Attribute", "Value"),
    NORMALISE = Table.AddColumn(#"Unpivoted Other Columns", "L2.1Value", each [Value]/[D]),
    #"Replaced Value" = Table.ReplaceValue(NORMALISE,"L2.1sw ","AB ",Replacer.ReplaceText,{"Attribute"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Replaced Value",{{"Attribute", "Inefficiencies"}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns1",{"Value"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns3", List.Distinct(#"Removed Columns3"[Inefficiencies]), "Inefficiencies", "L2.1Value", List.Sum),
    #"2POND" = Table.AddColumn(#"Pivoted Column", "Level", each "2-Pond"),
    #"Changed Type" = Table.TransformColumnTypes(#"2POND",{{"Date", type date}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Changed Type",{"Level", "Date", "Facility", "Version", "L1.1", "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential"})
in
    #"Reordered Columns2";

shared InefficienciesAG1_1L2DayL0Org = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check"  }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"Column1", "Column2"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential"}),
    #"added 0.1 ORG" = Table.AddColumn(#"Reordered Columns", "Level", each "0.1-Org"),
    #"Grouped Rows" = Table.Group(#"added 0.1 ORG", {"Level", "Version"}, {{"L2.1sw EXCESS OVER-ALLOCATION", each List.Sum([#"wAB EXCESS OVER-ALLOCATION"]), type number}, {"L2.1sw SPARE STRETCH", each List.Sum([wAB SPARE STRETCH]), type number}, {"L2.1sw SPARE SLACK", each List.Sum([wAB SPARE SLACK]), type number}, {"L2.1sw SPARE STRETCH ALLOCATED", each List.Sum([wAB SPARE STRETCH ALLOCATED]), type number}, {"L2.1sw SPARE SLACK ALLOCATED", each List.Sum([wAB SPARE SLACK ALLOCATED]), type number}, {"L2.1sw EXCESS STRETCH (ALLOCATED)", each List.Sum([#"wAB EXCESS STRETCH (ALLOCATED)"]), type number}, {"L2.1sw EXCESS ALLOCATED SLACK", each List.Sum([wAB EXCESS ALLOCATED SLACK]), type number}, {"L2.1sw EXCESS STRETCH", each List.Sum([wAB EXCESS STRETCH]), type number}, {"L2.1sw EXCESS SLACK", each List.Sum([wAB EXCESS SLACK]), type number}, {"L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", each List.Sum([#"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)"]), type number}, {"L2.1sw POTENTIAL SHORTFALL", each List.Sum([wAB POTENTIAL SHORTFALL]), type number}, {"L2.1sw WASTED STRETCH", each List.Sum([wAB WASTED STRETCH]), type number}, {"L2.1sw WASTED SLACK", each List.Sum([wAB WASTED SLACK]), type number}, {"L2.1sw ALLOCATED SLACK", each List.Sum([wAB ALLOCATED SLACK]), type number}, {"L2.1sw ALLOCATED STRETCH", each List.Sum([wAB ALLOCATED STRETCH]), type number}, {"L2.1sw UNALLOCATED SLACK", each List.Sum([wAB UNALLOCATED SLACK]), type number}, {"L2.1sw LATENT", each List.Sum([wAB LATENT]), type number}, {"L2.1sw LATENT.ALLOCATED", each List.Sum([wAB LATENT.ALLOCATED]), type number}, {"L2.1sw  LATENT.OVERALLOCATED", each List.Sum([wAB LATENT.OVERALLOCATED]), type number}, {"D", each List.Sum([D]), type number}, {"Cs", each List.Sum([Cs]), type number}, {"CX", each List.Sum([CX]), type number}, {"A", each List.Sum([A]), type number}, {"L2.1sw Potential", each List.Sum([wPotential]), type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{ "D", "Cs", "CX", "A", "L2.1sw EXCESS OVER-ALLOCATION", "L2.1sw SPARE STRETCH",  "L2.1sw SPARE STRETCH ALLOCATED", "L2.1sw SPARE SLACK ALLOCATED", "L2.1sw EXCESS STRETCH (ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL", "L2.1sw WASTED STRETCH", "L2.1sw WASTED SLACK", "L2.1sw ALLOCATED STRETCH", "L2.1sw ALLOCATED SLACK", "L2.1sw UNALLOCATED SLACK", "L2.1sw LATENT", "L2.1sw LATENT.ALLOCATED", "L2.1sw  LATENT.OVERALLOCATED"}),
    #"Unpivoted Other Columns1" = Table.UnpivotOtherColumns(#"Reordered Columns1", {"Level", "D", "Cs", "CX", "A","Version"}, "Attribute", "Value"),
    NORMALISE = Table.AddColumn(#"Unpivoted Other Columns1", "L2.1Value", each [Value]/[D]),
    #"Replaced Value" = Table.ReplaceValue(NORMALISE,"L2.1sw ","AB ",Replacer.ReplaceText,{"Attribute"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Replaced Value",{{"Attribute", "Inefficiencies"}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns1",{"Value"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns3", List.Distinct(#"Removed Columns3"[Inefficiencies]), "Inefficiencies", "L2.1Value", List.Sum),
    #"Reordered Columns2" = Table.ReorderColumns(#"Pivoted Column",{"Level",  "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED"})
in
    #"Reordered Columns2";

shared #"Start Date" = let
    Source = #"EffortOutcomesAG1_1DayShiftIMPORT !!",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Date", Order.Ascending}}),
    #"Kept First Rows" = Table.FirstN(#"Sorted Rows",1),
    #"Changed Type" = Table.TransformColumnTypes(#"Kept First Rows",{{"Date", Int64.Type}}),
    Date = #"Changed Type"{0}[Date]
in
    Date;

shared InefficienciesAG1_1L2DayL0Facility = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check"  }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"Column1", "Column2"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential"}),
    #"added 0.1 ORG" = Table.AddColumn(#"Reordered Columns", "Level", each "0.1-Org"),
    #"Grouped Rows" = Table.Group(#"added 0.1 ORG", {"Facility", "Version"}, {{"L2.1sw EXCESS OVER-ALLOCATION", each List.Sum([#"wAB EXCESS OVER-ALLOCATION"]), type number}, {"L2.1sw SPARE STRETCH", each List.Sum([wAB SPARE STRETCH]), type number}, {"L2.1sw SPARE SLACK", each List.Sum([wAB SPARE SLACK]), type number}, {"L2.1sw SPARE STRETCH ALLOCATED", each List.Sum([wAB SPARE STRETCH ALLOCATED]), type number}, {"L2.1sw SPARE SLACK ALLOCATED", each List.Sum([wAB SPARE SLACK ALLOCATED]), type number}, {"L2.1sw EXCESS STRETCH (ALLOCATED)", each List.Sum([#"wAB EXCESS STRETCH (ALLOCATED)"]), type number}, {"L2.1sw EXCESS ALLOCATED SLACK", each List.Sum([wAB EXCESS ALLOCATED SLACK]), type number}, {"L2.1sw EXCESS STRETCH", each List.Sum([wAB EXCESS STRETCH]), type number}, {"L2.1sw EXCESS SLACK", each List.Sum([wAB EXCESS SLACK]), type number}, {"L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", each List.Sum([#"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)"]), type number}, {"L2.1sw POTENTIAL SHORTFALL", each List.Sum([wAB POTENTIAL SHORTFALL]), type number}, {"L2.1sw WASTED STRETCH", each List.Sum([wAB WASTED STRETCH]), type number}, {"L2.1sw WASTED SLACK", each List.Sum([wAB WASTED SLACK]), type number}, {"L2.1sw ALLOCATED SLACK", each List.Sum([wAB ALLOCATED SLACK]), type number}, {"L2.1sw ALLOCATED STRETCH", each List.Sum([wAB ALLOCATED STRETCH]), type number}, {"L2.1sw UNALLOCATED SLACK", each List.Sum([wAB UNALLOCATED SLACK]), type number}, {"L2.1sw LATENT", each List.Sum([wAB LATENT]), type number}, {"L2.1sw LATENT.ALLOCATED", each List.Sum([wAB LATENT.ALLOCATED]), type number}, {"L2.1sw  LATENT.OVERALLOCATED", each List.Sum([wAB LATENT.OVERALLOCATED]), type number}, {"L2.1sw Potential", each List.Sum([wPotential]), type number}, {"D", each List.Sum([D]), type number}, {"Cs", each List.Sum([Cs]), type number}, {"CX", each List.Sum([CX]), type number}, {"A", each List.Sum([A]), type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{ "D", "Cs", "CX", "A", "L2.1sw EXCESS OVER-ALLOCATION", "L2.1sw SPARE STRETCH",  "L2.1sw SPARE STRETCH ALLOCATED", "L2.1sw SPARE SLACK ALLOCATED", "L2.1sw EXCESS STRETCH (ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL", "L2.1sw WASTED STRETCH", "L2.1sw WASTED SLACK", "L2.1sw ALLOCATED STRETCH", "L2.1sw ALLOCATED SLACK", "L2.1sw UNALLOCATED SLACK", "L2.1sw LATENT", "L2.1sw LATENT.ALLOCATED", "L2.1sw  LATENT.OVERALLOCATED"}),
    #"Unpivoted Other Columns1" = Table.UnpivotOtherColumns(#"Reordered Columns1", {"Facility", "D", "Cs", "CX", "A","Version"}, "Attribute", "Value"),
    NORMALISE = Table.AddColumn(#"Unpivoted Other Columns1", "L2.1Value", each [Value]/[D]),
    #"Replaced Value" = Table.ReplaceValue(NORMALISE,"L2.1sw ","AB ",Replacer.ReplaceText,{"Attribute"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Replaced Value",{{"Attribute", "Inefficiencies"}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns1",{"Value"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns3", List.Distinct(#"Removed Columns3"[Inefficiencies]), "Inefficiencies", "L2.1Value", List.Sum),
    #"Reordered Columns2" = Table.ReorderColumns(#"Pivoted Column",{"Facility",  "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED"}),
    #"Added LEVEL" = Table.AddColumn(#"Reordered Columns2", "Level", each "1.1-Pond(Facility)"),
    #"Reordered Columns3" = Table.ReorderColumns(#"Added LEVEL",{"Facility", "Version", "AB  LATENT.OVERALLOCATED", "Level", "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB Potential"})
in
    #"Reordered Columns3";

shared InefficienciesAG1_1L2FacilityDay = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check"  }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"Column1", "Column2"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential"}),
    #"Grouped Rows" = Table.Group(#"Reordered Columns", {"Date", "Facility", "Version"}, {{"L2.1sw EXCESS OVER-ALLOCATION", each List.Sum([#"wAB EXCESS OVER-ALLOCATION"]), type number}, {"L2.1sw SPARE STRETCH", each List.Sum([wAB SPARE STRETCH]), type number}, {"L2.1sw SPARE SLACK", each List.Sum([wAB SPARE SLACK]), type number}, {"L2.1sw SPARE STRETCH ALLOCATED", each List.Sum([wAB SPARE STRETCH ALLOCATED]), type number}, {"L2.1sw SPARE SLACK ALLOCATED", each List.Sum([wAB SPARE SLACK ALLOCATED]), type number}, {"L2.1sw EXCESS STRETCH (ALLOCATED)", each List.Sum([#"wAB EXCESS STRETCH (ALLOCATED)"]), type number}, {"L2.1sw EXCESS ALLOCATED SLACK", each List.Sum([wAB EXCESS ALLOCATED SLACK]), type number}, {"L2.1sw EXCESS STRETCH", each List.Sum([wAB EXCESS STRETCH]), type number}, {"L2.1sw EXCESS SLACK", each List.Sum([wAB EXCESS SLACK]), type number}, {"L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", each List.Sum([#"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)"]), type number}, {"L2.1sw POTENTIAL SHORTFALL", each List.Sum([wAB POTENTIAL SHORTFALL]), type number}, {"L2.1sw WASTED STRETCH", each List.Sum([wAB WASTED STRETCH]), type number}, {"L2.1sw WASTED SLACK", each List.Sum([wAB WASTED SLACK]), type number}, {"L2.1sw ALLOCATED SLACK", each List.Sum([wAB ALLOCATED SLACK]), type number}, {"L2.1sw ALLOCATED STRETCH", each List.Sum([wAB ALLOCATED STRETCH]), type number}, {"L2.1sw UNALLOCATED SLACK", each List.Sum([wAB UNALLOCATED SLACK]), type number}, {"L2.1sw LATENT", each List.Sum([wAB LATENT]), type number}, {"L2.1sw LATENT.ALLOCATED", each List.Sum([wAB LATENT.ALLOCATED]), type number}, {"L2.1sw  LATENT.OVERALLOCATED", each List.Sum([wAB LATENT.OVERALLOCATED]), type number}, {"D", each List.Sum([D]), type number}, {"Cs", each List.Sum([Cs]), type number}, {"CX", each List.Sum([CX]), type number}, {"A", each List.Sum([A]), type number}, {"L2.1sw Potential", each List.Sum([wPotential]), type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{"Date",  "D", "Cs", "CX", "A", "L2.1sw EXCESS OVER-ALLOCATION", "L2.1sw SPARE STRETCH",  "L2.1sw SPARE STRETCH ALLOCATED", "L2.1sw SPARE SLACK ALLOCATED", "L2.1sw EXCESS STRETCH (ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL", "L2.1sw WASTED STRETCH", "L2.1sw WASTED SLACK", "L2.1sw ALLOCATED STRETCH", "L2.1sw ALLOCATED SLACK", "L2.1sw UNALLOCATED SLACK", "L2.1sw LATENT", "L2.1sw LATENT.ALLOCATED", "L2.1sw  LATENT.OVERALLOCATED"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Reordered Columns1", {"Date",   "D", "Cs",  "CX", "A","Facility","Version"}, "Attribute", "Value"),
    NORMALISE = Table.AddColumn(#"Unpivoted Other Columns", "L2.1Value", each [Value]/[D]),
    #"Replaced Value" = Table.ReplaceValue(NORMALISE,"L2.1sw ","AB ",Replacer.ReplaceText,{"Attribute"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Replaced Value",{{"Attribute", "Inefficiencies"}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns1",{"Value"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns3", List.Distinct(#"Removed Columns3"[Inefficiencies]), "Inefficiencies", "L2.1Value", List.Sum),
    #"1LAKE" = Table.AddColumn(#"Pivoted Column", "Level", each "3.1-Pond FacilityDay"),
    #"Changed Type" = Table.TransformColumnTypes(#"1LAKE",{{"Date", type date}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Changed Type",{"Level", "Facility", "Version", "Date", "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential"})
in
    #"Reordered Columns2";

shared InefficienciesAG1_1L2Day = let
    Source = Excel.CurrentWorkbook(){[Name="EffortOutcomesAG1_1DayShiftAB"]}[Content],
    #"Removed Columns" = Table.RemoveColumns(Source,{"Level", "Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block", "Check"  }),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AB Potential", "Potential"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns",{"Column1", "Column2"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Removed Columns1",{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "Potential"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns2",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential"}),
    #"Grouped Rows" = Table.Group(#"Reordered Columns", {"Date", "Version"}, {{"L2.1sw EXCESS OVER-ALLOCATION", each List.Sum([#"wAB EXCESS OVER-ALLOCATION"]), type number}, {"L2.1sw SPARE STRETCH", each List.Sum([wAB SPARE STRETCH]), type number}, {"L2.1sw SPARE SLACK", each List.Sum([wAB SPARE SLACK]), type number}, {"L2.1sw SPARE STRETCH ALLOCATED", each List.Sum([wAB SPARE STRETCH ALLOCATED]), type number}, {"L2.1sw SPARE SLACK ALLOCATED", each List.Sum([wAB SPARE SLACK ALLOCATED]), type number}, {"L2.1sw EXCESS STRETCH (ALLOCATED)", each List.Sum([#"wAB EXCESS STRETCH (ALLOCATED)"]), type number}, {"L2.1sw EXCESS ALLOCATED SLACK", each List.Sum([wAB EXCESS ALLOCATED SLACK]), type number}, {"L2.1sw EXCESS STRETCH", each List.Sum([wAB EXCESS STRETCH]), type number}, {"L2.1sw EXCESS SLACK", each List.Sum([wAB EXCESS SLACK]), type number}, {"L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", each List.Sum([#"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)"]), type number}, {"L2.1sw POTENTIAL SHORTFALL", each List.Sum([wAB POTENTIAL SHORTFALL]), type number}, {"L2.1sw WASTED STRETCH", each List.Sum([wAB WASTED STRETCH]), type number}, {"L2.1sw WASTED SLACK", each List.Sum([wAB WASTED SLACK]), type number}, {"L2.1sw ALLOCATED SLACK", each List.Sum([wAB ALLOCATED SLACK]), type number}, {"L2.1sw ALLOCATED STRETCH", each List.Sum([wAB ALLOCATED STRETCH]), type number}, {"L2.1sw UNALLOCATED SLACK", each List.Sum([wAB UNALLOCATED SLACK]), type number}, {"L2.1sw LATENT", each List.Sum([wAB LATENT]), type number}, {"L2.1sw LATENT.ALLOCATED", each List.Sum([wAB LATENT.ALLOCATED]), type number}, {"L2.1sw  LATENT.OVERALLOCATED", each List.Sum([wAB LATENT.OVERALLOCATED]), type number}, {"D", each List.Sum([D]), type number}, {"Cs", each List.Sum([Cs]), type number}, {"CX", each List.Sum([CX]), type number}, {"A", each List.Sum([A]), type number}, {"L2.1sw Potential", each List.Sum([wPotential]), type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Grouped Rows",{"Date",  "D", "Cs", "CX", "A", "L2.1sw EXCESS OVER-ALLOCATION", "L2.1sw SPARE STRETCH",  "L2.1sw SPARE STRETCH ALLOCATED", "L2.1sw SPARE SLACK ALLOCATED", "L2.1sw EXCESS STRETCH (ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL (OVER-ALLOCATED)", "L2.1sw POTENTIAL SHORTFALL", "L2.1sw WASTED STRETCH", "L2.1sw WASTED SLACK", "L2.1sw ALLOCATED STRETCH", "L2.1sw ALLOCATED SLACK", "L2.1sw UNALLOCATED SLACK", "L2.1sw LATENT", "L2.1sw LATENT.ALLOCATED", "L2.1sw  LATENT.OVERALLOCATED"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Reordered Columns1", {"Date",   "D", "Cs",  "CX", "A","Version"}, "Attribute", "Value"),
    NORMALISE = Table.AddColumn(#"Unpivoted Other Columns", "L2.1Value", each [Value]/[D]),
    #"Replaced Value" = Table.ReplaceValue(NORMALISE,"L2.1sw ","AB ",Replacer.ReplaceText,{"Attribute"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Replaced Value",{{"Attribute", "Inefficiencies"}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Renamed Columns1",{"Value"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns3", List.Distinct(#"Removed Columns3"[Inefficiencies]), "Inefficiencies", "L2.1Value", List.Sum),
    #"1LAKE" = Table.AddColumn(#"Pivoted Column", "Level", each "1-Lake"),
    #"Changed Type" = Table.TransformColumnTypes(#"1LAKE",{{"Date", type date}}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Changed Type",{"Level", "Version", "Date", "D", "Cs", "CX", "A", "AB SPARE SLACK", "AB EXCESS OVER-ALLOCATION", "AB EXCESS ALLOCATED SLACK", "AB EXCESS STRETCH", "AB EXCESS SLACK", "AB SPARE STRETCH", "AB SPARE STRETCH ALLOCATED", "AB SPARE SLACK ALLOCATED", "AB EXCESS STRETCH (ALLOCATED)", "AB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "AB POTENTIAL SHORTFALL", "AB WASTED STRETCH", "AB WASTED SLACK", "AB ALLOCATED STRETCH", "AB ALLOCATED SLACK", "AB UNALLOCATED SLACK", "AB LATENT", "AB LATENT.ALLOCATED", "AB  LATENT.OVERALLOCATED", "AB Potential"})
in
    #"Reordered Columns2";

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