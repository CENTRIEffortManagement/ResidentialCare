// Power Query from: OvertimeRosterChecks.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\TE\2. Calculations\OvertimeRosterChecks.xlsx
// Extracted: 2026-11-12T20:54:34.741Z

section Section1;

shared #"LINK  OvertimeRosterChecks Whiddon Lists" = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation = "2.0", ViewMode = "All"]);

shared #"LINK OvertimeRosterChecks Roles" = #"LINK  OvertimeRosterChecks Whiddon Lists"{[Id = "6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items];

shared #"LINK OvertimeRosterChecks Facilities" = #"LINK  OvertimeRosterChecks Whiddon Lists"{[Id = "1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items];

shared #"IMPORT CentriSyncPaths" = let
    SourcePath = "C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true)
in
    Navigation;

shared PathTABLE = let
    Source = Excel.CurrentWorkbook(),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TrimmedValue = if value = null then null else Text.Trim(value),
            SlashNormalized = if value = null then null else Text.Replace(TrimmedValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    WorkbookPathInputs = Table.SelectRows(Source, each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
    WorkbookPathTable = if Table.RowCount(WorkbookPathInputs) = 1 then WorkbookPathInputs{0}[Content]
        else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
    WorkbookPathColumn = if Table.HasColumns(WorkbookPathTable, {"FilePath"}) then Table.SelectColumns(WorkbookPathTable, {"FilePath"})
        else if Table.ColumnNames(WorkbookPathTable) = {"Column1"} then Table.RenameColumns(WorkbookPathTable, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or be a single named cell.",
    WorkbookPathRow = if Table.RowCount(WorkbookPathColumn) = 1 then WorkbookPathColumn{0}
        else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = WorkbookPathRow[FilePath],
    NormalizedWorkbookPath = if not Value.Is(RawFilePath, type text) then
            error "FilePathUrl[FilePath] must contain the current workbook path as text."
        else if Text.Trim(RawFilePath) = "" then
            error "FilePathUrl[FilePath] is blank. Save and recalculate OvertimeRosterChecks.xlsx."
        else NormalizePath(RawFilePath),
    InputWorkbookFolder = Text.BeforeDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputWorkbookSuffix = Text.AfterDelimiter(NormalizedWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    InputFileName = if Text.StartsWith(InputWorkbookSuffix, "[") then Text.BetweenDelimiters(InputWorkbookSuffix, "[", "]") else InputWorkbookSuffix,
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "OvertimeRosterChecks.xlsx") = 0 then InputWorkbookFolder & "\" & InputFileName
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
    CentriSyncPaths_Source = #"IMPORT CentriSyncPaths",
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_Selected = Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}),
    CentriSyncPaths_Typed = Table.TransformColumnTypes(
        CentriSyncPaths_Selected,
        {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    CentriSyncPaths_Normalized = Table.TransformColumns(CentriSyncPaths_Typed, {
        {"SharepointRootUrl", each NormalizePath(_), type nullable text},
        {"SyncedFolderRootPath", each NormalizePath(_), type nullable text}
    }),
    BufferedMappings = Table.Buffer(CentriSyncPaths_Normalized),
    SharePointCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    SharePointDocumentsCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each
        if [SharepointRootUrl] = null or Text.Trim([SharepointRootUrl]) = "" then null
        else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalCandidates = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    CandidatesWithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each
        if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(CandidatesWithLength, each
        [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    LongestMatch = if Table.IsEmpty(SortedMatches) then error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath else SortedMatches{0},
    LongestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    MatchedDestinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch = if List.Count(MatchedDestinations) = 1 then LongestMatch
        else error "CentriSyncPaths contains conflicting mappings for: " & FilePath,
    UntrimmedRelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath = Text.TrimStart(UntrimmedRelativePath, "\"),
    LocalFullPath = if RelativePath = "" then BestMatch[SyncedFolderRootPath]
        else BestMatch[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    FolderSegments = Text.Split(WorkbookFolder, "\"),
    Segments = List.Select(FolderSegments, each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase),
    CalculationIndex = List.PositionOf(Segments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidUnitFolder = UnitsIndex >= 0 and CalculationIndex = UnitsIndex + 2 and CalculationIndex = List.Count(Segments) - 1,
    UnitRoot = if ValidUnitFolder then Text.BeforeDelimiter(WorkbookFolder, "\", {0, RelativePosition.FromEnd})
        else error "OvertimeRosterChecks.xlsx must be under UNITS/<unit>/2. Calculations.",
    PortfolioSegments = List.FirstN(Segments, ResidentialCareIndex),
    PortfolioRoot = if ResidentialCareIndex > 0 then Text.Combine(PortfolioSegments, "\")
        else error "Resolved workbook path does not contain the ResidentialCare boundary.",
    UnitName = Segments{UnitsIndex + 1},
    Output = #table(
        {"Variable Name", "Value"},
        {
            {"Root Path", WorkbookFolder},
            {"Unit Root", UnitRoot},
            {"Portfolio Root", PortfolioRoot},
            {"Unit", UnitName},
            {"FileName", InputFileName}
        }),
    #"Buffered Path Output" = Table.Buffer(Output)
in
    #"Buffered Path Output";

shared Unit1Path = let
    Source = PathTABLE,
    SelectedPathRow = Source{[Variable Name = "Unit Root"]},
    UnitPath = SelectedPathRow[Value]
in
    UnitPath;

shared #"IMPORT CentriSyncPaths (2)" = let
        SourcePath = "C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx",
        SourceBinary = File.Contents(SourcePath),
        Navigation = Excel.Workbook(SourceBinary, null, true)
    in
        Navigation;

shared #"IMPORT Worker's Report" = let
    SourcePath = #"Unit1Path" & "\1. Input\Worker's Report.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    ReportSheet = Navigation{[Item="Sheet1",Kind="Sheet"]}[Data]
in
    ReportSheet;

shared #"IMPORT Allocation Extracted" = let
    SourcePath = #"Unit1Path" & "\1. Input\1-AllocationExtracted.xlsx",
    SourceBinary = File.Contents(SourcePath),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    AllocationExtracted_Sheet = Navigation{[Item="AllocationExtracted",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(AllocationExtracted_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Location", type text}, {"Pay Company", type text}, {"Department", type text}, {"Area", type text}, {"Name", type text}, {"Employment Type", type text}, {"RoleShift", type text}, {"Employee Code", Int64.Type}, {"Date", type date}, {"Day Of Week", type text}, {"Shift Type", type text}, {"Start", type datetime}, {"End", type datetime}, {"Break", Int64.Type}, {"Shift Length", type number}, {"Hours", type number}, {"Rate", type number}, {"Published", type logical}, {"Published At", type datetime}, {"Published By", type text}, {"Non Attended", type logical}, {"Status", type text}, {"Code", Int64.Type}, {"Facility-Abbrev", type text}, {"Role", type text}, {"Unit", type any}}),
    Custom1 = Table.Buffer(#"Changed Type")
in
    Custom1;

shared #"IMPORT Setting Data" = let

    Source = Excel.Workbook(File.Contents(Unit1Path & "\2. Calculations\Settings Data.xlsx"), null, true)
in
    Source;

shared #"EXPORT Setting Permutations" = let

    Source = Excel.Workbook(File.Contents(Unit1Path & "\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data]
in
    PermutationDimensions_Table;

shared #"EXPORT Shift Periods" = let

    Source = Excel.Workbook(File.Contents(Unit1Path & "\2. Calculations\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"Role", type text}, {"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type time}, {"DurationOfShifts", type number}, {"EndDay", type time}, {"EndTimeCheck", type time}})
in
    #"Changed Type";

shared #"LINK OvertimeRosterThreshold" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"]),
    #"ae98d433-faad-4ec0-9209-525ceaa71a29" = Source{[Id="ae98d433-faad-4ec0-9209-525ceaa71a29"]}[Items]
in
    #"ae98d433-faad-4ec0-9209-525ceaa71a29";

shared #"DailyHoursLimit-Most" = let
    Source = #"LINK OvertimeRosterThreshold",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = "DailyHoursLimit")),
    #"1" = #"Filtered Rows"{[ID=1]}[Value]
in
    #"1";

shared #"DailyHoursLimit-CS" = let
    Source = #"LINK OvertimeRosterThreshold",
    #"Filtered Rows" = Table.SelectRows(Source, each [Title] = "CSHospitalityHoursLimit"),
    #"2" = #"Filtered Rows"{[ID=2]}[Value]
in
    #"2";

shared MinimumBreakHours = let
    Source = #"LINK OvertimeRosterThreshold",
    #"Filtered Rows" = Table.SelectRows(Source, each [Title] = "MinimumBreakHours"),
    #"4" = #"Filtered Rows"{[ID=4]}[Value]
in
    #"4";

shared OTThresholdx2 = let
    Source = #"LINK OvertimeRosterThreshold",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = "OTThresholdx2")),
    #"10" = #"Filtered Rows"{[ID=10]}[Value]
in
    #"10";

shared ShiftPeriods = let
    Source = #"EXPORT Shift Periods",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Role", "ShiftPeriod", "StartDay", "EndDay"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Removed Other Columns", {"Role", "ShiftPeriod"}, "Point", "Time"),
    #"Inserted Merged Column" = Table.AddColumn(#"Unpivoted Columns", "ShiftPoint", each Text.Combine({[ShiftPeriod], [Point]}, "-"), type text),
    #"Replaced Value" = Table.ReplaceValue(#"Inserted Merged Column","Day","",Replacer.ReplaceText,{"ShiftPoint"}),
    #"Extracted Time" = Table.TransformColumns(#"Replaced Value",{}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Extracted Time",{"ShiftPoint", "Time", "Role"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns1",{{"Time", type number}}),
    #"Pivoted Column" = Table.Pivot(#"Changed Type", List.Distinct(#"Changed Type"[ShiftPoint]), "ShiftPoint", "Time")
in
    #"Pivoted Column";

shared #"Roster Prepare" = let
    Source = #"IMPORT Allocation Extracted",
    #"Filtered NOTPUBLISHED" = Table.SelectRows(Source, each ([Published] = true)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NOTPUBLISHED",{"Employee Code", "Date", "Start", "End", "Hours", "Facility-Abbrev", "Role"})
in
    #"Removed Other Columns";

shared #"Overtime+BreakSufficiency-Prepare" = let
    Source = #"Roster Prepare",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Start", type time}, {"End", type time}}),
    #"Sorted Rows2" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Start", Order.Ascending}}),
    #"Added DATESTART" = Table.AddColumn(#"Sorted Rows2", "DateStart", each [Date] & [Start]),
    #"Added DATEEND" = Table.AddColumn(#"Added DATESTART", "DateEnd", each [Date] & [End]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added DATEEND",{{"DateStart", type datetime}, {"DateEnd", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type1",{{"Employee Code", Order.Ascending}, {"DateStart", Order.Ascending}}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Sorted Rows1",{"Employee Code", "Date", "Start", "End", "Hours", "Role", "DateStart", "DateEnd"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Other Columns1", "Index", 1, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 0, 1, Int64.Type),
    #"Merged Queries" = Table.NestedJoin(#"Added Index1", {"Index.1"}, #"Added Index1", {"Index"}, "Added Index.1", JoinKind.LeftOuter),
    #"Expanded Added Index.2" = Table.ExpandTableColumn(#"Merged Queries", "Added Index.1", {"Hours", "DateEnd"}, {"Hours.1", "DateStart-Next"}),
    #"Inserted HRSB2B" = Table.AddColumn(#"Expanded Added Index.2", "HrsBackToBack", each [Hours.1] + [Hours], type number),
    #"Removed Other Columns2" = Table.SelectColumns(#"Inserted HRSB2B",{"Employee Code", "Date", "Start", "End", "Hours", "Role", "DateStart", "DateEnd", "Index", "Index.1", "Hours.1", "DateStart-Next", "HrsBackToBack"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Other Columns2",{{"DateEnd", type number}, {"DateStart-Next", type number}, {"DateStart", type number}}),
    #"Filtered ENDROSTER NULL" = Table.SelectRows(#"Changed Type2", each ([#"DateStart-Next"] <> null)),
    #"Added BACKTOBACK" = Table.AddColumn(#"Filtered ENDROSTER NULL", "BackToBack", each if [DateStart] = [#"DateStart-Next"] then true else false),
    #"Added OVERTIME1" = Table.AddColumn(#"Added BACKTOBACK", "Overtime", each if [BackToBack] = true 
then 
if
[Role] <> "CS" and [HrsBackToBack] > #"DailyHoursLimit-Most" then "OvertimeShift"
else 
        if [Role] = "CS" and [HrsBackToBack] > #"DailyHoursLimit-CS" 
       then
     "OvertimeShift"
        else
        null
        else null),
    #"Merged Queries1" = Table.NestedJoin(#"Added OVERTIME1", {"Employee Code", "Index"}, #"Added OVERTIME1", {"Employee Code", "Index.1"}, "Expanded Added Index.1", JoinKind.LeftOuter),
    #"Expanded Expanded Added Index.1" = Table.ExpandTableColumn(#"Merged Queries1", "Expanded Added Index.1", {"Overtime"}, {"Overtime.1"}),
    #"Added Suffix" = Table.TransformColumns(#"Expanded Expanded Added Index.1", {{"Overtime", each _ & "2", type text}}),
    #"Added Suffix1" = Table.TransformColumns(#"Added Suffix", {{"Overtime.1", each _ & "1", type text}}),
    #"Merged Columns" = Table.CombineColumns(#"Added Suffix1",{"Overtime.1", "Overtime"},Combiner.CombineTextByDelimiter("" , QuoteStyle.None),"Overtime"),
    #"Replaced Value" = Table.ReplaceValue(#"Merged Columns","",null,Replacer.ReplaceValue,{"Overtime"}),
    #"Inserted SHIFTBREAK" = Table.AddColumn(
    #"Replaced Value",
    "ShiftBreak",
    each
        if [BackToBack] = false 
then
            ([DateStart]-[#"DateStart-Next"] )*24
        else
            0,
    type number
),
    #"Added SHIFTBREAKSUFFFICIENCY" = Table.AddColumn(#"Inserted SHIFTBREAK", "ShiftBreakSufficiency", each if [HrsBackToBack] <= #"DailyHoursLimit-Most" 
then null
else if
([ShiftBreak] + 0.01) < MinimumBreakHours and [ShiftBreak] > 0 then "InsufficientBreak" else null),
    #"Merged Queries2" = Table.NestedJoin(#"Added SHIFTBREAKSUFFFICIENCY", {"Employee Code", "Index"}, #"Added SHIFTBREAKSUFFFICIENCY", {"Employee Code", "Index.1"}, "Filtered Rows2", JoinKind.LeftOuter),
    #"Expanded Filtered Rows2" = Table.ExpandTableColumn(#"Merged Queries2", "Filtered Rows2", {"ShiftBreakSufficiency"}, {"ShiftBreakSufficiency.1"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Filtered Rows2",{{"Index", Order.Ascending}}),
    #"Added Suffix2" = Table.TransformColumns(#"Sorted Rows", {{"ShiftBreakSufficiency", each _ & "2", type text}}),
    #"Added Suffix3" = Table.TransformColumns(#"Added Suffix2", {{"ShiftBreakSufficiency.1", each _ & "1", type text}}),
    INSUFFICIENTBREAK = Table.CombineColumns(#"Added Suffix3",{"ShiftBreakSufficiency", "ShiftBreakSufficiency.1"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"ShiftBreakSufficiency"),
    #"Replaced Value1" = Table.ReplaceValue(INSUFFICIENTBREAK,"",null,Replacer.ReplaceValue,{"ShiftBreakSufficiency"}),
    #"Added SHIFTSCOMBINED" = Table.AddColumn(#"Replaced Value1", "ShiftType", each [Overtime] ?? [ShiftBreakSufficiency]),
    #"Filtered PENALTIES" = Table.SelectRows(#"Added SHIFTSCOMBINED", each ([ShiftType] <> null)),
    #"Sorted Rows4" = Table.Sort(#"Filtered PENALTIES",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}, {"Start", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows4",{"Employee Code", "Date", "Start", "End", "Role", "DateStart-Next", "HrsBackToBack", "ShiftBreak", "ShiftType"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Removed Other Columns",{{"DateStart-Next", type datetime}}),
    #"Appended Query" = Table.Combine({#"Changed Type3", TABLEShiftType}),
    #"Inserted Last Characters" = Table.AddColumn(#"Appended Query", "ShiftTypeOrder", each Text.End([ShiftType], 1), type text),
    #"Extracted First Characters1" = Table.TransformColumns(
    #"Inserted Last Characters",
    {{"ShiftType", each Text.Start(_, Text.Length(_) - 1), type text}}
),
    #"Pivoted Column" = Table.Pivot(#"Extracted First Characters1", List.Distinct(#"Extracted First Characters1"[ShiftType]), "ShiftType", "HrsBackToBack", List.Sum),
    #"Filtered Rows" = Table.SelectRows(#"Pivoted Column", each ([Employee Code] <> null)),
    #"Added Conditional Column" = Table.AddColumn(#"Filtered Rows", "ShiftType", each 
if [OvertimeShift] = null 
then "InsuffcientBreak" 
       else if [InsufficientBreak] = null 
       then "OvertimeShift" 
       else "OT+InsufficientBreak"),
    #"Added HRSBACKTOBACK" = Table.AddColumn(#"Added Conditional Column", "HrsBackToBack", each [OvertimeShift] ?? [InsufficientBreak]),
    #"Removed Columns" = Table.RemoveColumns(#"Added HRSBACKTOBACK",{"DateStart-Next"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER;

shared #"Overtime+BreakSufficiency-Segmented" = let
    Source = #"Overtime+BreakSufficiency-Prepare",
    #"Sorted Rows" = Table.Sort(Source,{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows",  "DateLate", each if [Start] > [End] then Date.AddDays([Date],1) else null ),
    #"Sorted Rows1" = Table.Sort(#"Added Custom",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows1",{"HrsBackToBack", "OvertimeShift", "InsufficientBreak"}),
    #"Unpivoted Columns1" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Employee Code", "Role",  "ShiftType",  "ShiftBreak", "Start", "End","ShiftTypeOrder"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Columns1",{{"Value", "Date"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Date", type date}})
in
    #"Changed Type";

shared #"Overtime+BreakSufficiency" = let
    Source = #"Overtime+BreakSufficiency-Segmented",
    #"Duplicated Column" = Table.DuplicateColumn(Source, "Start", "Start - Copy"),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Duplicated Column", "End", "End - Copy"),
    #"Inserted Merged Column" = Table.AddColumn(#"Duplicated Column1", "Merged", each Text.Combine({Text.From([#"Start - Copy"], "en-AU"), Text.From([#"End - Copy"], "en-AU")}, "-"), type text),
    #"Sorted Rows1" = Table.Sort(#"Inserted Merged Column",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}, {"Start", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows1", "Index", 1, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 0, 1, Int64.Type),
    #"Merged Queries1" = Table.NestedJoin(#"Added Index1", {"Index"}, #"Added Index1", {"Index.1"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded Added Index1" = Table.ExpandTableColumn(#"Merged Queries1", "Added Index1", {"Date", "Start", "End"}, {"Date.1", "Start.Next", "End.Previous"}),
    #"Added BACKTOBACK" = Table.AddColumn(#"Expanded Added Index1", "BackToBack", each if [Date] = [Date.1] then
            if [Start] = [End.Previous] then "B2BStart"
            else if [End] = [Start.Next] then "B2BEnd"
            else null
        else null),
    #"Removed Columns" = Table.RemoveColumns(#"Added BACKTOBACK",{"Index", "Index.1", "Date.1", "Start.Next", "End.Previous"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Start", type number}, {"End", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Role"}, ShiftPeriods, {"Role"}, "EXPORT Shift Periods (2)", JoinKind.LeftOuter),
    #"Expanded EXPORT Shift Periods (2)" = Table.ExpandTableColumn(#"Merged Queries", "EXPORT Shift Periods (2)", {"AM-Start", "AM-End", "PM-Start", "PM-End", "NIGHT-Start", "NIGHT-End"}, {"AM-Start", "AM-End", "PM-Start", "PM-End", "NIGHT-Start", "NIGHT-End"}),
    #"Added SHIFTCORE" = Table.AddColumn(#"Expanded EXPORT Shift Periods (2)", "Shift", each if 
([Start] >= [#"AM-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start]  < [#"PM-Start"] - #"EXTRACT ShiftBoundardyTolerance" )
or 
([End] > [#"AM-Start"] +  #"EXTRACT ShiftBoundardyTolerance"
and [End]  < [#"PM-Start"] + #"EXTRACT ShiftBoundardyTolerance") 
then "AM"

else if
([Start] >= [#"PM-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start]  <  [#"NIGHT-Start"]  - #"EXTRACT ShiftBoundardyTolerance" )
or 
([End] > [#"PM-Start"] +  #"EXTRACT ShiftBoundardyTolerance"
and ([End] <  [#"NIGHT-End"] + #"EXTRACT ShiftBoundardyTolerance" 
or [End]-24 < [#"NIGHT-End"] + #"EXTRACT ShiftBoundardyTolerance" -0.5 )) 
then "PM"

else if 
[Attribute] ="DateLate" 
or 
([Start]  >= [#"NIGHT-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start] -0.5 <  [#"AM-Start"] +0.5 -
#"EXTRACT ShiftBoundardyTolerance" )
or 
([End] - 0.5 > [#"NIGHT-Start"] - 0.5 +  #"EXTRACT ShiftBoundardyTolerance"
and [End] -0.5 < [#"AM-Start"] + #"EXTRACT ShiftBoundardyTolerance") 
then "NIGHT" 
else "ERROR"),
    #"Removed Other Columns" = Table.SelectColumns(#"Added SHIFTCORE",{"Shift", "Employee Code", "ShiftType", "Date", "Role"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows", "Facility", each Unit)
in
    #"Added Custom";

shared #"Overtime+BreakSufficiency-Penalties" = let
    Source = #"Overtime+BreakSufficiency-Prepare",
    #"Inserted OVERTIME" = Table.AddColumn(Source, "Overtime", each [HrsBackToBack] - 

(if [Role] = "CS" then #"DailyHoursLimit-CS" else  #"DailyHoursLimit-Most"), type number),
    #"Added OTx1.5" = Table.AddColumn(#"Inserted OVERTIME", "OTx1.5", each if [Overtime] < OTThresholdx2 then [Overtime]
else  OTThresholdx2),
    #"Changed Type" = Table.TransformColumnTypes(#"Added OTx1.5",{{"OTx1.5", type number}}),
    #"Added OTx2" = Table.AddColumn(#"Changed Type", "OTx2", each if [OTx1.5] > OTThresholdx2 then 0
else [Overtime] - [OTx1.5]),
    #"Added Custom" = Table.AddColumn(#"Added OTx2", "Unit", each Unit),
    #"Sorted Rows3" = Table.Sort(#"Added Custom",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}})
in
    #"Sorted Rows3";

shared YearWeekStart = let
    Source = #"Roster Shifts",
    #"Calculated Week of Year" = Table.TransformColumns(Source,{{"Date", Date.WeekOfYear, Int64.Type}}),
    Date = #"Calculated Week of Year"[Date],
    #"Calculated Minimum" = List.Min(Date)
in
    #"Calculated Minimum";

shared #"DailyHoursLimit-Most (2)" = let
    Source = #"LINK OvertimeRosterThreshold",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = "DailyHoursLimit"))
in
    #"Filtered Rows";

shared FortnightHours = let
    Source = #"Roster Prepare",
    #"Inserted Week of Year" = Table.AddColumn(Source, "Week of Year", each Date.WeekOfYear([Date]), Int64.Type),
    #"Subtracted from Column" = Table.TransformColumns(#"Inserted Week of Year", {{"Week of Year", each _ - YearWeekStart+1, type number}}),
    #"Added Conditional Column" = Table.AddColumn(#"Subtracted from Column", "Fortnight", each if [Week of Year] = 1 then 1 else if [Week of Year] = 2 then 1 else 2),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"Employee Code", "Fortnight"}, {{"Hours", each List.Sum([Hours]), type nullable number}})
in
    #"Grouped Rows";

shared #"Roster Shifts" = let
    Source = #"Roster Prepare",
    #"Duplicated Column" = Table.DuplicateColumn(Source, "Start", "Start - Copy"),
    #"Filtered Rows" = Table.SelectRows(#"Duplicated Column", each ([Employee Code] = 22818)),
    #"Duplicated Column1" = Table.DuplicateColumn(#"Filtered Rows", "End", "End - Copy"),
    #"Changed Type1" = Table.TransformColumnTypes(#"Duplicated Column1",{{"Start", type time}, {"End", type time}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Start", type number}, {"End", type number}, {"Start - Copy", type time}, {"End - Copy", type time}}),
    #"Inserted Merged Column" = Table.AddColumn(#"Changed Type", "Merged", each Text.Combine({Text.From([#"Start - Copy"], "en-AU"), Text.From([#"End - Copy"], "en-AU")}, "-"), type text),
    #"Sorted Rows1" = Table.Sort(#"Inserted Merged Column",{{"Employee Code", Order.Ascending}, {"Date", Order.Ascending}, {"Start", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows1", "Index", 1, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 0, 1, Int64.Type),
    #"Merged Queries1" = Table.NestedJoin(#"Added Index1", {"Index"}, #"Added Index1", {"Index.1"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded Added Index1" = Table.ExpandTableColumn(#"Merged Queries1", "Added Index1", {"Date", "Start", "End"}, {"Date.1", "Start.Next", "End.Previous"}),
    #"Added BACKTOBACK" = Table.AddColumn(#"Expanded Added Index1", "BackToBack", each if [Date] = [Date.1] then
            if [Start] = [End.Previous] then "B2BStart"
            else if [End] = [Start.Next] then "B2BEnd"
            else null
        else null),
    #"Removed Columns" = Table.RemoveColumns(#"Added BACKTOBACK",{"Index", "Index.1", "Date.1", "Start.Next", "End.Previous"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Role"}, ShiftPeriods, {"Role"}, "EXPORT Shift Periods (2)", JoinKind.LeftOuter),
    #"Expanded EXPORT Shift Periods (2)" = Table.ExpandTableColumn(#"Merged Queries", "EXPORT Shift Periods (2)", {"AM-Start", "AM-End", "PM-Start", "PM-End", "NIGHT-Start", "NIGHT-End"}, {"AM-Start", "AM-End", "PM-Start", "PM-End", "NIGHT-Start", "NIGHT-End"}),
    #"Added SHIFTCORE" = Table.AddColumn(#"Expanded EXPORT Shift Periods (2)", "ShiftCore", each if
([Start] >= [#"AM-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start]  < [#"PM-Start"] - #"EXTRACT ShiftBoundardyTolerance" )
or 
([End] > [#"AM-Start"] +  #"EXTRACT ShiftBoundardyTolerance"
and [End]  < [#"PM-Start"] - #"EXTRACT ShiftBoundardyTolerance") 
then "AM"

else if
([Start] >= [#"PM-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start]  <  [#"NIGHT-Start"]  - #"EXTRACT ShiftBoundardyTolerance" )
or 
([End] > [#"PM-Start"] +  #"EXTRACT ShiftBoundardyTolerance"
and ([End] <  [#"NIGHT-End"] + #"EXTRACT ShiftBoundardyTolerance" 
or [End]-24 < [#"NIGHT-End"] + #"EXTRACT ShiftBoundardyTolerance" -0.5 )) 
then "PM"

else if 
([Start]  >= [#"NIGHT-Start"] - #"EXTRACT ShiftBoundardyTolerance"
and [Start] -0.5 <  [#"AM-Start"] +0.5 -
#"EXTRACT ShiftBoundardyTolerance" )
or 
([End] - 0.5 > [#"NIGHT-Start"] - 0.5 +  #"EXTRACT ShiftBoundardyTolerance"
and [End] -0.5 < [#"AM-Start"] + #"EXTRACT ShiftBoundardyTolerance") 
then "NIGHT" 
else "ERROR"),
    #"Added SHIFTEARLY" = Table.AddColumn(#"Added SHIFTCORE", "ShiftEarly", each if ([Role] <> "CS" and [Hours] > #"DailyHoursLimit-CS" )
 or 
([Role] = "CS" and [Hours] > #"DailyHoursLimit-CS" )
then
( if
([Start] < [#"AM-Start"] 
and [Start]  > 0 + #"EXTRACT ShiftBoundardyTolerance" )
then "NIGHT"

else if
([Start] < [#"PM-Start"] 
and [Start]  >  [#"AM-Start"]  + #"EXTRACT ShiftBoundardyTolerance" )
then "AM"

else if 
([Start]  < [#"NIGHT-Start"] 
and [Start] > [#"PM-Start"]   +   #"EXTRACT ShiftBoundardyTolerance" )
then "PM"

else null
)
else null),
    #"Added SHIFTLATE" = Table.AddColumn(#"Added SHIFTEARLY", "ShiftLate", each if

([End] > [#"PM-End"] and [End] < 1)
or
[End]  > 0 and [End]  < #"EXTRACT ShiftBoundardyTolerance" 
then "NIGHT"

else if
([End] > [#"AM-Start"] 
and [End]  <  [#"PM-End"]  - #"EXTRACT ShiftBoundardyTolerance" )
then "AM"

else if 
([End]  > [#"PM-End"] 
and [End] < [#"NIGHT-End"] -  #"EXTRACT ShiftBoundardyTolerance" )
then "PM"

else null),
    #"Renamed Columns" = Table.RenameColumns(#"Added SHIFTLATE",{{"Date", "DateX"}}),
    DATEE = Table.AddColumn(#"Renamed Columns", "DateE", each if ([ShiftCore] ="AM" and [ShiftEarly] = "NIGHT" ) 
or 
([ShiftCore] ="NIGHT" and [ShiftEarly] = "PM" )
then 
Date.AddDays([DateX],-1)
else 
[DateX]),
    DATEY = Table.AddColumn(DATEE, "DateY", each if ([ShiftCore] ="NIGHT" and [ShiftLate] = "AM" ) 
or 
([ShiftCore] ="PM" and [ShiftLate] = "NIGHT" )
then 
Date.AddDays([DateX],1)
else 
[DateX]),
    #"Inserted Merged Column2" = Table.AddColumn(DATEY, "Core", each Text.Combine({Text.From([DateX], "en-AU"), [ShiftCore]}, "-"), type text),
    #"Inserted Merged Column3" = Table.AddColumn(#"Inserted Merged Column2", "Early", each Text.Combine({Text.From([DateE], "en-AU"), Text.From([ShiftEarly], "en-AU")}, "-"), type text),
    #"Inserted Merged Column4" = Table.AddColumn(#"Inserted Merged Column3", "Late", each Text.Combine({Text.From([DateY], "en-AU"), [ShiftLate]}, "-"), type text),
    #"Removed Other Columns" = Table.SelectColumns(#"Inserted Merged Column4",{"Late", "Early", "Core", "Employee Code", "Facility-Abbrev", "Role"}),
    #"Unpivoted Columns1" = Table.UnpivotOtherColumns(#"Removed Other Columns", {"Employee Code", "Facility-Abbrev", "Role"}, "Attribute", "Value"),
    #"Removed Columns2" = Table.RemoveColumns(#"Unpivoted Columns1",{"Attribute"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Removed Columns2", "Value", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Date", "Shift"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Date", type date}}),
    #"Filtered Rows1" = Table.SelectRows(#"Changed Type2", each ([Shift] <> null)),
    #"Removed Duplicates" = Table.Distinct(#"Filtered Rows1"),
    #"Sorted Rows2" = Table.Sort(#"Removed Duplicates",{{"Date", Order.Ascending}}),
    BUFFER = Table.Buffer(#"Sorted Rows2"),
    #"Sorted Rows" = Table.Sort(BUFFER,{{"Date", Order.Ascending}})
in
    #"Sorted Rows";

shared #"LINK EMParameters" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"])
in
    Source;

shared #"EXTRACT ShiftBoundardyTolerance" = let
    Source = #"LINK EMParameters",
    #"750bc5b4-f054-4517-af37-24175cb90b1e" = Source{[Id="750bc5b4-f054-4517-af37-24175cb90b1e"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"750bc5b4-f054-4517-af37-24175cb90b1e",{"Parameter", "Value"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [Parameter] = "ShiftBoundaryTolerance"),
    Value = #"Filtered Rows"{0}[Value],
    #"Divided Value" = Value / 24
in
    #"Divided Value";

shared TABLEShiftType = let
    Source = {"OvertimeShift1","OvertimeShift2","InsufficientBreak1","InsufficientBreak2"},
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "ShiftType"}})
in
    #"Renamed Columns";

shared Unit = let
    Source = PathTABLE,
    Value = Source{3}[Value]
in
    Value;