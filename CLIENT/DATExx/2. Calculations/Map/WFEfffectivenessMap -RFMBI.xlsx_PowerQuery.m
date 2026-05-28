// Power Query from: WFEfffectivenessMap -RFMBI.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\2. Calculations\Map\WFEfffectivenessMap -RFMBI.xlsx
// Extracted: 2026-05-21T00:47:48.507Z

section Section1;

shared LocationsPNE = let
    Source = Excel.CurrentWorkbook(){[Name="Locations"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Town,Population,Northing,Easting", type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Changed Type", "Town,Population,Northing,Easting", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Town,Population,Northing,Easting.1", "Town,Population,Northing,Easting.2", "Town,Population,Northing,Easting.3", "Town,Population,Northing,Easting.4", "Town,Population,Northing,Easting.5"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Town,Population,Northing,Easting.1", type text}, {"Town,Population,Northing,Easting.2", type text}, {"Town,Population,Northing,Easting.3", type text}, {"Town,Population,Northing,Easting.4", Int64.Type}, {"Town,Population,Northing,Easting.5", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Town,Population,Northing,Easting.2", "Population"}, {"Town,Population,Northing,Easting.3", "Northing"}, {"Town,Population,Northing,Easting.4", "Easting"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Town,Population,Northing,Easting.5"}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Removed Columns", "New SE", Splitter.SplitTextByDelimiter(":", QuoteStyle.Csv), {"New SE.1", "New SE.2"}),
    #"Split Column by Delimiter2" = Table.SplitColumn(#"Split Column by Delimiter1", "New SE.2", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"New SE.2.1", "New SE.2.2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Split Column by Delimiter2",{"Northing", "Easting", "New SE.1"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns1",{{"New SE.2.1", "South"}, {"New SE.2.2", "East"}, {"Town,Population,Northing,Easting.1", "Facility"}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns1",{"Place", "Address"})
in
    #"Removed Columns2";

shared SumPopulation = let
    Source = LocationsPNE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Population"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Population", Int64.Type}}),
    Population = #"Changed Type"[Population],
    #"Calculated Sum" = List.Sum(Population)
in
    #"Calculated Sum";

shared FacilityEffectivenessMatrix = let
    Source = Excel.CurrentWorkbook(){[Name="Table3"]}[Content]
in
    Source;

shared FaciltiyEffectivenessList = let
    Source = FacilityEffectivenessMatrix,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Excess Potential", type number}, {"Excess Rostered", type number}, {"Shortfall", type number}, {"Unallocated Potential", type number}, {"Ability to deliver", type number}}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Column1"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Columns",{{"Column1", "Facility"}, {"Attribute", "MetricType"}, {"Value", "MetricValue"}})
in
    #"Renamed Columns";

shared FaciltiyEffectivenessSummary = let
    Source = Excel.CurrentWorkbook(){[Name="Table3"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Excess Potential", type number}, {"Excess Rostered", type number}, {"Shortfall", type number}, {"Unallocated Potential", type number}, {"Ability to deliver", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Column1"}, LocationsPNE, {"Facility"}, "LocationsPNE", JoinKind.LeftOuter),
    #"Expanded LocationsPNE" = Table.ExpandTableColumn(#"Merged Queries", "LocationsPNE", {"Population", "Beds"}, {"Population", "Beds"}),
    #"Added Custom" = Table.AddColumn(#"Expanded LocationsPNE", "Custom", each (1- [Ability to deliver])*30 + [Excess Rostered]*75),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"Custom", "$Ineff.Day.Resident"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "$Ineff.Day.Facility", each [Beds]*[#"$Ineff.Day.Resident"])
in
    #"Added Custom1";