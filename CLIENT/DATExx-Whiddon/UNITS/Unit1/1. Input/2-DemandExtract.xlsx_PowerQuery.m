// Power Query from: 2-DemandExtract.xlsx
// Pathname: CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\2-DemandExtract.xlsx
// Extracted: 2026-08-26T10:04:05.610Z

section Section1;

shared LocalUnitFolder = let
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
    LocalWorkbookFolder = if IsLocalPath then WorkbookFolder else error "FilePathUrl did not resolve to a local path. Open/save this workbook from the active local sync folder before refreshing.",
    UnitFolder = if Text.Contains(LocalWorkbookFolder, "\1. Input") then Text.BeforeDelimiter(LocalWorkbookFolder, "\1. Input", {0, RelativePosition.FromEnd}) else if Text.Contains(LocalWorkbookFolder, "\2. Calculations") then Text.BeforeDelimiter(LocalWorkbookFolder, "\2. Calculations", {0, RelativePosition.FromEnd}) else LocalWorkbookFolder
in
    UnitFolder;

shared StartAM = let
    Source = ShiftStart,
    StartAM1 = Source{0}[StartAM]
in
    StartAM1;

shared StartPM = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartPM]
in
    StartPM1;

shared StartNIGHT = let
    Source = ShiftStart,
    StartPM1 = Source{0}[StartNIGHT]
in
    StartPM1;

shared ShiftOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WcvRV0lEyVIrViVYKADGNwEw/T3ePECDPWCk2FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t, Column2 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "Shift"}, {"Column2", "Order"}})
in
    #"Renamed Columns";

shared #"IMPORT LocRoleDayShift%" = let
    Source = Excel.Workbook(File.Contents(LocalUnitFolder & "\1. Input\Demand-MasterRoster Manual Read.xlsx"), null, true),
    #"LocRoleDayShift%_Table" = Source{[Item="LocRoleDayShift%",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(#"LocRoleDayShift%_Table",{{"Location", type text}, {"Role", type text}, {"Attribute", type text}, {"Value", type number}, {"Hours", type number}, {"LocRoleDayShift%", Percentage.Type}})
in
    #"Changed Type";

shared ShiftStart = let
    Source = Excel.Workbook(File.Contents(LocalUnitFolder & "\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftStart_Table = Source{[Item="ShiftStart",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftStart_Table,{{"StartAM", type number}, {"StartPM", type number}, {"StartNIGHT", type number}})
in
    #"Changed Type";

shared ShiftDuration = let
    Source = Excel.Workbook(File.Contents(LocalUnitFolder & "\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared #"IMPORT ShiftPeriod" = let
    Source = Excel.Workbook(File.Contents(LocalUnitFolder & "\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}, {"DurationOfShifts", type number}})
in
    #"Changed Type";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents(LocalUnitFolder & "\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    #"Changed Type";

shared #"LocRoleDayShift Prepare" = let
    Source = #"IMPORT LocRoleDayShift%",
    #"Added Shift" = Table.AddColumn(Source, "Shift", each if Text.EndsWith([Attribute], "AM", Comparer.OrdinalIgnoreCase) then "AM" else if Text.EndsWith([Attribute], "PM", Comparer.OrdinalIgnoreCase) then "PM" else if Text.EndsWith([Attribute], "NS", Comparer.OrdinalIgnoreCase) then "NIGHT" else error "Unrecognized day/shift attribute: " & [Attribute], type text),
    #"Added Day Name" = Table.AddColumn(#"Added Shift", "DayName", each Text.Start([Attribute], Text.Length([Attribute]) - 2), type text),
    #"Added Day" = Table.AddColumn(#"Added Day Name", "Day", each let DayIndex = List.PositionOf({"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}, [DayName]) in if DayIndex < 0 then error "Unrecognized weekday in attribute: " & [Attribute] else DayIndex + 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Day",{{"Location", "Unit"}, {"Value", "DemandHRS"}}),
    #"Removed Source Columns" = Table.RemoveColumns(#"Renamed Columns",{"Attribute", "DayName", "Hours", "LocRoleDayShift%"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Source Columns",{{"Unit", type text}, {"Role", type text}, {"Day", Int64.Type}, {"Shift", type text}, {"DemandHRS", type number}}),
    #"Merged Shift Period" = Table.NestedJoin(#"Changed Type", {"Shift", "Role"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "IMPORT ShiftPeriod", JoinKind.LeftOuter),
    #"Expanded Shift Period" = Table.ExpandTableColumn(#"Merged Shift Period", "IMPORT ShiftPeriod", {"DurationOfShifts"}, {"DurationOfShifts"}),
    #"Added DemandFTE" = Table.AddColumn(#"Expanded Shift Period", "DemandFTE", each if [DurationOfShifts] = null or [DurationOfShifts] = 0 then error "Missing or zero shift duration for role " & [Role] & " and shift " & [Shift] else [DemandHRS] / [DurationOfShifts], type number),
    #"Reordered Columns" = Table.ReorderColumns(#"Added DemandFTE",{"Unit", "Role", "Day", "Shift", "DemandFTE", "DemandHRS", "DurationOfShifts"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Unit", Order.Ascending}, {"Role", Order.Ascending}, {"Day", Order.Ascending}, {"Shift", Order.Ascending}})
in
    #"Sorted Rows";

shared #"Permutation DateTimeRoleShift" = let
    Source = #"IMPORT PermutationDimensions",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"RolesList", "Roles"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns", {"Shifts", "Roles"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "IMPORT ShiftPeriod", JoinKind.LeftOuter),
    #"Expanded IMPORT ShiftPeriod1" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ShiftPeriod", {"StartDay", "EndDay"}, {"StartDay", "EndDay"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Expanded IMPORT ShiftPeriod1",{{"StartDay", "Start"}, {"EndDay", "End"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns2",{{"Start", type time}, {"End", type time}}),
    #"START TIMEDATE" = Table.AddColumn(#"Changed Type", "StartTime", each [Date] & [Start], type datetime),
    #"END TIMEDATE" = Table.AddColumn(#"START TIMEDATE", "EndTime", each if [Shifts] = "NIGHT" then Date.AddDays([Date],1) & [End]
else [Date]&[End]),
    #"Changed Type1" = Table.TransformColumnTypes(#"END TIMEDATE",{{"EndTime", type datetime}}),
    #"Added FACILITY" = Table.AddColumn(#"Changed Type1", "Facility", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added FACILITY",{"Start", "End"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Shifts", "Shift"}, {"Roles", "Role"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns1",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared ShiftUnitDemandHRS = let
    Source = #"Permutation DateTimeRoleShift",
    #"Merged Queries" = Table.NestedJoin(Source, {"Role", "Day", "Shift"}, #"LocRoleDayShift Prepare", {"Role", "Day", "Shift"}, "LocRoleDayShift Prepare", JoinKind.LeftOuter),
    #"Expanded LocRoleDayShift Prepare" = Table.ExpandTableColumn(#"Merged Queries", "LocRoleDayShift Prepare", {"Unit", "DemandFTE", "DemandHRS", "DurationOfShifts"}, {"Unit", "DemandFTE", "DemandHRS", "DurationOfShifts"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded LocRoleDayShift Prepare",{"Date", "Day", "Shift", "Period", "Role", "StartTime", "EndTime", "Unit", "Facility", "DemandFTE", "DemandHRS", "DurationOfShifts"}),
    #"Sorted Rows1" = Table.Sort(#"Reordered Columns",{{"Period", Order.Ascending}, {"Role", Order.Ascending}, {"StartTime", Order.Ascending}, {"EndTime", Order.Ascending}, {"Unit", Order.Ascending}})
in
    #"Sorted Rows1";
