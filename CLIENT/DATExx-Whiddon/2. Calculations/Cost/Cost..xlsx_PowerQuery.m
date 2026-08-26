// Power Query from: Cost..xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\Cost\Cost..xlsx
// Extracted: 2026-05-21T00:48:03.036Z

section Section1;

shared TargetRatebyDate = let
    Source = PermutationDimensions,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Date", "DateX"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX] ,
28)),
    #"Grouped Rows" = Table.Group(#"Added Custom1", {"Date"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "TargetRate", each TargetDailyRate),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Date", type date}})
in
    #"Changed Type";

shared ShiftHrs = 7.6 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared RatesDelta = let
    Source = Excel.CurrentWorkbook(){[Name="RatesDelta"]}[Content],
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Role", "Period"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

shared EffortOutcomesAG1_1DayShiftAB = let
     Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "\2. Calculations\E-O-I\Inefficiencies.xlsx"), null, true),
    EffortOutcomesAG1_1DayShiftAB_Table = Source{[Item="EffortOutcomesAG1_1DayShiftAB",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortOutcomesAG1_1DayShiftAB_Table,{{"Level", type text}, {"Potential", Int64.Type}, {"Planned", Int64.Type}, {"Degree", Int64.Type}, {"Spare", Int64.Type}, {"Spare Allocated", Int64.Type}, {"Block", Int64.Type}, {"EXCESS OVER-ALLOCATION", type number}, {"SPARE STRETCH", type number}, {"SPARE SLACK", type number}, {"SPARE STRETCH ALLOCATED", type number}, {"SPARE SLACK ALLOCATED", type number}, {"EXCESS STRETCH (ALLOCATED)", type number}, {"EXCESS ALLOCATED SLACK", type number}, {"EXCESS STRETCH", type number}, {"EXCESS SLACK", type number}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"POTENTIAL SHORTFALL", type number}, {"WASTED STRETCH", type number}, {"WASTED SLACK", type number}, {"ALLOCATED STRETCH", type number}, {"ALLOCATED SLACK", type number}, {"UNALLOCATED SLACK", type number}, {"LATENT", Int64.Type}, {"LATENT.ALLOCATED", Int64.Type}, {"LATENT.OVERALLOCATED", Int64.Type}, {"AB Potential", Int64.Type}, {"Check", type number}, {"wAB EXCESS OVER-ALLOCATION", type number}, {"wAB SPARE STRETCH", type number}, {"wAB SPARE SLACK", type number}, {"wAB SPARE STRETCH ALLOCATED", type number}, {"wAB SPARE SLACK ALLOCATED", type number}, {"wAB EXCESS STRETCH (ALLOCATED)", type number}, {"wAB EXCESS ALLOCATED SLACK", type number}, {"wAB EXCESS STRETCH", type number}, {"wAB EXCESS SLACK", type number}, {"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"wAB POTENTIAL SHORTFALL", type number}, {"wAB WASTED STRETCH", type number}, {"wAB WASTED SLACK", type number}, {"wAB ALLOCATED STRETCH", type number}, {"wAB ALLOCATED SLACK", type number}, {"wAB UNALLOCATED SLACK", type number}, {"wAB LATENT", Int64.Type}, {"wAB LATENT.ALLOCATED", Int64.Type}, {"wAB LATENT.OVERALLOCATED", Int64.Type}, {"wPotential", Int64.Type}, {"Column1", type any}, {"Column2", type any}, {"Date", type date}, {"L1.1", type text}, {"L1.2", type text}, {"D", type number}, {"Cs", type number}, {"CX", type number}, {"A", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "Apn", "Apx", "Ai", "Level", "EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "AB Potential", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns",{"Level", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2"})
in
    #"Removed Columns1";

shared PermutationDimensions = let
         Source1 = Excel.CurrentWorkbook(){[Name="Folder"]}[Content],
    Folder = Source1{0}[Folder],
    
    
    Source = Excel.Workbook(File.Contents(Folder &  "FACILITIES\ASHB\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared Rates = let
    Source = Excel.CurrentWorkbook(){[Name="Rates"]}[Content],
    #"Added Custom" = Table.AddColumn(Source, "$OT", each [OTLoading]*[Rate]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "$Agency", each [AgencyPremium]*[Rate]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{ "OTLoading", "AgencyPremium"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Rate", "$Rate"}})
in
    #"Renamed Columns";

shared TargetDailyRate = 5 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared CostAG1_1DayShiftAB = let
    Source = EffortOutcomesAG1_1DayShiftAB,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] <> null)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Cs", "CX", "A", "Apn", "Apx", "Ai", "DATESHIFT", "CapacityMaxHC"}),
    EXCESSALLOCATION = Table.AddColumn(#"Removed Columns", "ExcessAllocation", each 1*(
 
[#"EXCESS OVER-ALLOCATION"]+
[SPARE STRETCH ALLOCATED]+
[SPARE SLACK ALLOCATED]+
[#"EXCESS STRETCH (ALLOCATED)"]+
[EXCESS ALLOCATED SLACK])),
    #"Added REMOVECARERforRN" = Table.AddColumn(EXCESSALLOCATION, "ExcessAllocation-Reaslistic", each if [Facility]= "Berry" and [L1.2] = "NIGHT" and [L1.1] = "Carer" then 0
else [ExcessAllocation]),
    SHORTAGE = Table.AddColumn(#"Added REMOVECARERforRN", "Shortage", each [D]*([POTENTIAL SHORTFALL]+
[WASTED STRETCH]+
[WASTED SLACK]+
[UNALLOCATED SLACK])),
    #"Added SHORTAGE-REALISTIC" = Table.AddColumn(SHORTAGE, "Shortage-Realistic", each if [Facility]= "Berry" and [L1.2] = "NIGHT" and [L1.1] = "Registered Nurse" then 0
else [Shortage]),
    PLANNEDOT = Table.AddColumn(#"Added SHORTAGE-REALISTIC", "PlannedOvertime", each [D]*[#"POTENTIAL SHORTFALL (OVER-ALLOCATED)"]),
    #"Removed Columns1" = Table.RemoveColumns(PLANNEDOT,{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "AB Potential"}),
    #"Inserted Day Name" = Table.AddColumn(#"Removed Columns1", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Merged Queries" = Table.NestedJoin(#"Inserted Day Name", {"L1.1", "L1.2", "Day Name"}, RatesDelta, {"Role", "Period", "Attribute"}, "Rates", JoinKind.LeftOuter),
    #"Expanded Rates1" = Table.ExpandTableColumn(#"Merged Queries", "Rates", {"Attribute", "Value"}, {"Rates.Attribute", "Rates.Value"}),
    UNNECESSARYALLOCATION = Table.AddColumn(#"Expanded Rates1", "UnnecessaryAllocation$", each 23
*[#"ExcessAllocation-Reaslistic"]*ShiftHrs),
    SHORTAGEREPLACE = Table.AddColumn(UNNECESSARYALLOCATION, "ShortageReplacement$", each [#"Shortage-Realistic"]*
[Rates.Value]
*ShiftHrs),
    #"OVERTIME$" = Table.AddColumn(SHORTAGEREPLACE, "Overtime$", each [PlannedOvertime]
*0
*ShiftHrs),
    #"Replaced Value" = Table.ReplaceValue(#"OVERTIME$",null,0,Replacer.ReplaceValue,{"UnnecessaryAllocation$", "ShortageReplacement$", "Overtime$"}),
    INEFFICIENCYCOST = Table.AddColumn(#"Replaced Value", "InefficiencyCost", each [#"UnnecessaryAllocation$"]
+[#"ShortageReplacement$"]
+[#"Overtime$"])
in
    INEFFICIENCYCOST;

shared CostAG1_1DayShiftABSummary = let
    Source = CostAG1_1DayShiftAB,
    #"Grouped Rows" = Table.Group(Source, {"L1.1", "Facility", "Version"}, {{"RoleCost", each List.Sum([InefficiencyCost]), type number}, {"UnnecessaryAllocation", each List.Sum([#"UnnecessaryAllocation$"]), type number}, {"Shortage Replacement", each List.Sum([#"ShortageReplacement$"]), type number}, {"OT", each List.Sum([#"Overtime$"]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Version", Order.Ascending}, {"Facility", Order.Ascending}, {"L1.1", Order.Ascending}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Sorted Rows",{"Version", "Facility", "L1.1", "RoleCost", "UnnecessaryAllocation", "Shortage Replacement", "OT"})
in
    #"Reordered Columns";

shared Saving = let
    Source = CostAG1_1DayShiftAB,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Version", "L1.2", "L1.1", "Date","InefficiencyCost","Facility"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Version] <> null)),
    #"Pivoted Column" = Table.Pivot(#"Filtered Rows", List.Distinct(#"Filtered Rows"[Version]), "Version", "InefficiencyCost", List.Sum),
    #"Inserted Subtraction" = Table.AddColumn(#"Pivoted Column", "Saving", each [Base] - [AllocationChange], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"AllocationChange", "Base"}),
    #"Grouped Rows" = Table.Group(#"Removed Columns", {"Date", "Facility"}, {{"Savings", each List.Sum([Saving]), type number}})
in
    #"Grouped Rows";

shared SavingFacilityDay = let
    Source = CostAG1_1DayShiftAB,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Version", "L1.2", "L1.1", "Date","InefficiencyCost","Facility"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[Version]), "Version", "InefficiencyCost", List.Sum),
    #"Inserted Subtraction" = Table.AddColumn(#"Pivoted Column", "Saving", each [Base] - [AllocationChange], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"AllocationChange", "Base"}),
    #"Grouped Rows" = Table.Group(#"Removed Columns", {"Date", "Facility"}, {{"Savings", each List.Sum([Saving]), type number}})
in
    #"Grouped Rows";

shared CostPerDay = let
    Source = CostAG1_1DayShiftAB,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Facility"}, {{"OpportunityCostDay", each List.Sum([InefficiencyCost]), type nullable number}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Facility"}, {{"Count", each List.Sum([OpportunityCostDay]), type nullable number}})
in
    #"Grouped Rows1";