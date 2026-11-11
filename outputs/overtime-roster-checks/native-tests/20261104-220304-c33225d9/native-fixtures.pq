section OvertimeRosterChecksNative;
shared OvertimeRosterChecksNativeTests = () as table => let
Evaluate = (Inputs as record) as record => let
#"IMPORT OvertimeRosterChecks CentriSyncPaths" = Inputs[PathMapping],
#"INPUT OvertimeRosterChecks FilePathUrl" = Inputs[PathInput],
OvertimeRosterChecks_Path = let
    NormalizePath = (value as any) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(Text.From(value)), "/", "\"), "\"),
    PathInput = #"INPUT OvertimeRosterChecks FilePathUrl",
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then PathInput[FilePath]
        else if Table.ColumnNames(PathInput) = {"Column1"} then PathInput[Column1]
        else error "FilePathUrl must have a FilePath column or be one named cell.",
    RawPath = if List.Count(PathColumn) = 1 then NormalizePath(PathColumn{0}) else error "FilePathUrl requires exactly one value.",
    NonBlankPath = if RawPath = null or RawPath = "" then error "Save and recalculate OvertimeRosterChecks.xlsx before resolving FilePathUrl." else RawPath,
    WorkbookPath = if Text.Contains(NonBlankPath, "[") then Text.BeforeDelimiter(NonBlankPath, "[") & Text.BetweenDelimiters(NonBlankPath, "[", "]") else NonBlankPath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NamedWorkbookPath = if Comparer.OrdinalIgnoreCase(FileName, "OvertimeRosterChecks.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
    MappingColumns = Table.SelectColumns(#"IMPORT OvertimeRosterChecks CentriSyncPaths", {"SharepointRootUrl", "SyncedFolderRootPath"}),
    NormalizedMappings = Table.TransformColumns(MappingColumns, {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    BufferedMappings = Table.Buffer(NormalizedMappings),
    SharePointRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    CandidateRoots = Table.Combine({SharePointRoots, DocumentRoots, LocalRoots}),
    RootLengths = Table.AddColumn(CandidateRoots, "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRoots = Table.SelectRows(RootLengths, each [RootLength] > 0 and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(NamedWorkbookPath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(NamedWorkbookPath) = [RootLength] or Text.Range(NamedWorkbookPath, [RootLength], 1) = "\")),
    OrderedRoots = Table.Sort(MatchingRoots, {{"RootLength", Order.Descending}}),
    LongestRoot = if Table.IsEmpty(OrderedRoots) then error "FilePathUrl has no CentriSyncPaths mapping." else OrderedRoots{0},
    EqualRoots = Table.SelectRows(OrderedRoots, each [RootLength] = LongestRoot[RootLength]),
    Destinations = List.Distinct(EqualRoots[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    SelectedRoot = if List.Count(Destinations) = 1 then LongestRoot else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    RelativePath = Text.TrimStart(Text.Range(NamedWorkbookPath, SelectedRoot[RootLength]), "\"),
    LocalWorkbookPath = SelectedRoot[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NestedSuffix = "\2. Calculations\OvertimeRosterChecks",
    ClientPath = if Text.EndsWith(WorkbookFolder, NestedSuffix, Comparer.OrdinalIgnoreCase) then Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(NestedSuffix))
        else error "OvertimeRosterChecks.xlsx must be in the organisation's 2. Calculations/OvertimeRosterChecks folder.",
    ClientSegments = Text.Split(ClientPath, "\"),
    OrganisationPath = if List.PositionOf(ClientSegments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase) < 0 then ClientPath
        else error "OvertimeRosterChecks.xlsx is organisation-level and cannot be located beneath UNITS.",
    Result = [ClientPath = OrganisationPath, WorkbookPath = LocalWorkbookPath]
in
    Result,
#"IMPORT Published Roster" = (unit as text) as table => if unit = "BD" then Inputs[RosterBD] else if unit = "TE" then Inputs[RosterTE] else if unit = "JH-RY" then Inputs[RosterJHRY] else error "Unapproved fixture unit.",
#"IMPORT Workers" = (unit as text) as table => if unit = "BD" then Inputs[WorkersBD] else if unit = "TE" then Inputs[WorkersTE] else if unit = "JH-RY" then Inputs[WorkersJHRY] else error "Unapproved fixture unit.",
#"IMPORT Settings Data" = (unit as text) as table => if unit = "BD" then Inputs[SettingsBD] else if unit = "TE" then Inputs[SettingsTE] else if unit = "JH-RY" then Inputs[SettingsJHRY] else error "Unapproved fixture unit.",
#"IMPORT OvertimeRosterChecks Whiddon Lists" = error "SharePoint navigation is disabled in native fixtures.",
#"LINK OvertimeRosterChecks Roles" = Inputs[Roles],
#"LINK OvertimeRosterChecks Facilities" = Inputs[Facilities],
#"LINK OvertimeRosterThreshold" = Inputs[Thresholds],
fnOvertimeRosterChecksText = (value as any) as nullable text =>
let
    TextValue = if value = null then null else Text.Trim(Text.From(value, "en-AU")),
    Result = if TextValue = "" then null else TextValue
in
    Result,
fnOvertimeRosterChecksNumber = (value as any) as nullable number =>
let
    Parsed = try Number.From(value) otherwise null,
    Result = if Parsed = null then null else if Number.IsNaN(Parsed) or Number.Abs(Parsed) = #infinity then null else Parsed
in
    Result,
fnOvertimeRosterChecksLogical = (value as any) as nullable logical =>
let
    TextValue = fnOvertimeRosterChecksText(value),
    Key = if TextValue = null then null else Text.Upper(TextValue),
    Result = if Value.Is(value, type logical) then value else if List.Contains({"TRUE", "YES", "1"}, Key) then true
        else if List.Contains({"FALSE", "NO", "0"}, Key) then false else null
in
    Result,
fnOvertimeRosterChecksKey = (values as list) as text => Binary.ToText(Json.FromValue(values), BinaryEncoding.Base64),
fnOvertimeRosterChecksClockSafe = (start as nullable datetime, finish as nullable datetime) as logical =>
let
    ValidBounds = start <> null and finish <> null and finish >= start,
    Years = if not ValidBounds then {} else {Date.Year(Date.From(start))..Date.Year(Date.From(finish))},
    WindowsByYear = List.Transform(Years, (year) => List.Transform({4, 10}, (month) =>
        let FirstDate = #date(year, month, 1), Sunday = Date.AddDays(FirstDate, Number.Mod(7 - Date.DayOfWeek(FirstDate, Day.Sunday), 7))
        in [WindowStart = Sunday & #time(2, 0, 0), WindowEnd = Sunday & #time(3, 0, 0)])),
    Windows = List.Combine(WindowsByYear),
    CrossesClockChange = if not ValidBounds then true else List.AnyTrue(List.Transform(Windows, each start < [WindowEnd] and finish > [WindowStart]))
in
    ValidBounds and not CrossesClockChange,
fnOvertimeRosterChecksSheetPrepare = (source as table, unit as text, required as list) as table =>
let
    Headers = if Table.HasColumns(source, required) then source else Table.PromoteHeaders(source, [PromoteAllScalars = true]),
    MissingColumns = List.Difference(required, Table.ColumnNames(Headers)),
    RequiredSchema = if List.IsEmpty(MissingColumns) then Headers else error Error.Record("OvertimeRosterChecks source schema", "Required source columns are missing.", [SourceUnit = unit, MissingColumns = MissingColumns]),
    SourceRows = Table.AddIndexColumn(RequiredSchema, "SourceRow", 1, 1, Int64.Type),
    SourceUnit = Table.AddColumn(SourceRows, "SourceUnit", each unit, type text)
in
    SourceUnit,
OvertimeRosterChecks_DailyHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "DailyHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one DailyHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "DailyHoursLimit must be a positive number." else if ParsedValue <= 0 then error "DailyHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_CSHospitalityHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSHospitalityHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSHospitalityHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSHospitalityHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSHospitalityHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_CSPersonalCareHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSPersonalCareHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSPersonalCareHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSPersonalCareHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSPersonalCareHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_MinimumBreakHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumBreakHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumBreakHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumBreakHours must be a positive number." else if ParsedValue <= 0 then error "MinimumBreakHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_ShortGapReviewHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "ShortGapReviewHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one ShortGapReviewHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "ShortGapReviewHours must be a positive number." else if ParsedValue <= 0 then error "ShortGapReviewHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_MinimumFortnightDutyFreeDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumFortnightDutyFreeDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumFortnightDutyFreeDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumFortnightDutyFreeDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "MinimumFortnightDutyFreeDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_FortnightHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightHoursLimit must be a positive number." else if ParsedValue <= 0 then error "FortnightHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_FortnightLengthDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightLengthDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightLengthDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightLengthDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "FortnightLengthDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_FortnightStartDate = let
    BDLandedSettings = #"IMPORT Settings Data"("BD"),
    TELandedSettings = #"IMPORT Settings Data"("TE"),
    JHRYLandedSettings = #"IMPORT Settings Data"("JH-RY"),
    PreparePermutationDates = (settings as table, unit as text) as table =>
        let MatchingTables = Table.SelectRows(settings, each [Item] = "PermutationDimensions" and [Kind] = "Table"),
            PermutationTable = if Table.RowCount(MatchingTables) = 1 then MatchingTables{0}[Data]
                else error Error.Record("OvertimeRosterChecks settings dates", "Expected exactly one PermutationDimensions table.", [SourceUnit = unit]),
            DateColumn = if Table.HasColumns(PermutationTable, {"Date"}) then Table.SelectColumns(PermutationTable, {"Date"})
                else error Error.Record("OvertimeRosterChecks settings dates", "PermutationDimensions must contain Date.", [SourceUnit = unit]),
            SuppliedDates = Table.SelectRows(DateColumn, each [Date] <> null),
            ValidatedDates = Table.TransformColumns(SuppliedDates, {{"Date", each try Date.From(_) otherwise
                error Error.Record("OvertimeRosterChecks settings dates", "PermutationDimensions Date contains an invalid date.", [SourceUnit = unit]), type date}})
        in ValidatedDates,
    BDDates = PreparePermutationDates(BDLandedSettings, "BD"),
    TEDates = PreparePermutationDates(TELandedSettings, "TE"),
    JHRYDates = PreparePermutationDates(JHRYLandedSettings, "JH-RY"),
    CombinedDates = Table.Combine({BDDates, TEDates, JHRYDates}),
    DateValues = CombinedDates[Date],
    FirstDate = if List.IsEmpty(DateValues) then null else List.Min(DateValues)
in
    FirstDate,
OvertimeRosterChecks_RosterPrepare = let
    RequiredColumns = {"Employee Code", "Employee_Code", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length"},
    BDLanded = #"IMPORT Published Roster"("BD"),
    TELanded = #"IMPORT Published Roster"("TE"),
    JHRYLanded = #"IMPORT Published Roster"("JH-RY"),
    BDRows = fnOvertimeRosterChecksSheetPrepare(BDLanded, "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(TELanded, "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(JHRYLanded, "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    NeededColumns = Table.SelectColumns(CombinedRows, {"SourceUnit", "SourceRow", "Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "Non Attended", "Status", "Published", "MealAvailabilityRequired", "MealStart", "MealEnd", "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    SourceFile = Table.AddColumn(NeededColumns, "SourceFile", each "UNITS/" & [SourceUnit] & "/1. Input/Allocation/Published Roster.xlsx", type text),
    RenamedSourceColumns = Table.RenameColumns(SourceFile, {{"Employee_Code", "AllocationCode"}, {"Employee Code", "WorkerEmployeeID"}, {"Pay Company", "EmployerKey"}, {"Employee Roster Name", "SourceName"}, {"Role", "SourceRole"}, {"Date", "SourceDate"}, {"Shift Type", "SourceShiftType"}, {"Employment Type", "SourceEmploymentType"}, {"Break Length Minutes", "BreakMinutes"}, {"Shift Net Length", "NetRosterHours"}}),
    NormalizedText = Table.TransformColumns(RenamedSourceColumns, {{"AllocationCode", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"EmployerKey", fnOvertimeRosterChecksText, type nullable text}, {"SourceName", fnOvertimeRosterChecksText, type nullable text}, {"SourceRole", fnOvertimeRosterChecksText, type nullable text}, {"Location", fnOvertimeRosterChecksText, type nullable text}, {"SourceEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"AgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"Grade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedQuantities = Table.TransformColumns(NormalizedText, {{"BreakMinutes", fnOvertimeRosterChecksNumber, type nullable number}, {"NetRosterHours", fnOvertimeRosterChecksNumber, type nullable number}, {"SourceDate", each try Date.From(_) otherwise null, type nullable date}, {"Start Time", each try Time.From(_) otherwise null, type nullable time}, {"End Time", each try Time.From(_) otherwise null, type nullable time}}),
    ParsedFacts = Table.TransformColumns(ParsedQuantities, {{"MealAvailabilityRequired", fnOvertimeRosterChecksLogical, type nullable logical}, {"OvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"BrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}, {"Published", fnOvertimeRosterChecksLogical, type nullable logical}, {"Non Attended", fnOvertimeRosterChecksLogical, type nullable logical}}),
    PublishedDutyCandidates = Table.SelectRows(ParsedFacts, each [Published] <> false and [Non Attended] <> true),
    ShiftStart = Table.AddColumn(PublishedDutyCandidates, "ShiftStart", each if [SourceDate] = null or [Start Time] = null then null else [SourceDate] & [Start Time], type nullable datetime),
    ShiftFinish = Table.AddColumn(ShiftStart, "ShiftFinish", each if [SourceDate] = null or [End Time] = null or [Start Time] = null then null
        else (if [End Time] < [Start Time] then Date.AddDays([SourceDate], 1) else [SourceDate]) & [End Time], type nullable datetime),
    GrossHours = Table.AddColumn(ShiftFinish, "GrossHours", each if [ShiftStart] = null or [ShiftFinish] = null then null else Duration.TotalHours([ShiftFinish] - [ShiftStart]), type nullable number),
    MealStartDateTime = Table.AddColumn(GrossHours, "MealStartDateTime", each
        let Raw = [MealStart], Parsed = if Value.Is(Raw, type time) and [SourceDate] <> null then [SourceDate] & Raw else try DateTime.From(Raw) otherwise null,
            Aligned = if Parsed <> null and [ShiftStart] <> null and Parsed < [ShiftStart] and Value.Is(Raw, type time) then Parsed + #duration(1, 0, 0, 0) else Parsed
        in Aligned, type nullable datetime),
    MealFinishDateTime = Table.AddColumn(MealStartDateTime, "MealFinishDateTime", each
        let Raw = [MealEnd], Parsed = if Value.Is(Raw, type time) and [SourceDate] <> null then [SourceDate] & Raw else try DateTime.From(Raw) otherwise null,
            Aligned = if Parsed <> null and [MealStartDateTime] <> null and Parsed <= [MealStartDateTime] and Value.Is(Raw, type time) then Parsed + #duration(1, 0, 0, 0) else Parsed
        in Aligned, type nullable datetime),
    DeclaredCoverage = Table.TransformColumns(MealFinishDateTime, {{"CoverageStart", each try DateTime.From(_) otherwise null, type nullable datetime},
        {"CoverageEnd", each if Value.Is(_, type date) then DateTime.From(Date.AddDays(_, 1)) else try DateTime.From(_) otherwise null, type nullable datetime}}),
    SourceRecordID = Table.AddColumn(DeclaredCoverage, "SourceRecordID", each fnOvertimeRosterChecksKey({[SourceUnit], [SourceFile], [SourceRow]}), type text),
    NarrowedSource = Table.RemoveColumns(SourceRecordID, {"Start Time", "End Time", "MealStart", "MealEnd"})
in
    NarrowedSource,
OvertimeRosterChecks_WorkerPrepare = let
    RequiredColumns = {"Employee Code"},
    BDLanded = #"IMPORT Workers"("BD"),
    TELanded = #"IMPORT Workers"("TE"),
    JHRYLanded = #"IMPORT Workers"("JH-RY"),
    BDRows = fnOvertimeRosterChecksSheetPrepare(BDLanded, "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(TELanded, "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(JHRYLanded, "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    SelectedFacts = Table.SelectColumns(CombinedRows, {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    NamedFacts = Table.RenameColumns(SelectedFacts, {{"Employee Code", "WorkerEmployeeID"}, {"Position Description", "WorkerPosition"}, {"Employee Status Description", "WorkerEmploymentType"}, {"AgreementStream", "WorkerAgreementStream"}, {"Grade", "WorkerGrade"}, {"OvertimeExempt", "WorkerOvertimeExempt"}, {"BrokenShiftException", "WorkerBrokenShiftException"}}),
    NormalizedText = Table.TransformColumns(NamedFacts, {{"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"WorkerPosition", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"WorkerAgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"WorkerGrade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedFacts = Table.TransformColumns(NormalizedText, {{"Start Date", each try Date.From(_) otherwise null, type nullable date}, {"Termination Date", each try Date.From(_) otherwise null, type nullable date}, {"Contracted FN Hours", fnOvertimeRosterChecksNumber, type nullable number}, {"WorkerOvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"WorkerBrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}}),
    IdentifiedWorkers = Table.SelectRows(ParsedFacts, each [WorkerEmployeeID] <> null),
    DistinctFacts = Table.Distinct(IdentifiedWorkers),
    FactFields = {"WorkerPosition", "WorkerEmploymentType", "Start Date", "Termination Date", "Contracted FN Hours", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException"},
    EmptyFacts = Record.FromList(List.Repeat({null}, List.Count(FactFields)), FactFields),
    GroupedWorkers = Table.Group(DistinctFacts, {"WorkerEmployeeID"}, {{"WorkerMatchCount", each Table.RowCount(_), Int64.Type}, {"WorkerFacts", each if Table.RowCount(_) = 1 then _{0} else EmptyFacts, type record}}),
    ExpandedFacts = Table.ExpandRecordColumn(GroupedWorkers, "WorkerFacts", {"WorkerPosition", "WorkerEmploymentType", "Start Date", "Termination Date", "Contracted FN Hours", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException"}),
    BufferedWorkers = Table.Buffer(ExpandedFacts)
in
    BufferedWorkers,
OvertimeRosterChecks_ShiftPrepare = let
    NormalizeKey = (value as any) as nullable text =>
        let TextValue = fnOvertimeRosterChecksText(value)
        in if TextValue = null then null else Text.Remove(Text.Upper(TextValue), {" ", "-", "_"}),
    PreparePeriods = (source as table, unit as text) as table =>
        let Selected = Table.SelectColumns(source, {"ShiftPeriod", "StartDay"}),
            NamedUnit = Table.AddColumn(Selected, "SourceUnit", each unit, type text),
            PeriodKeys = Table.TransformColumns(NamedUnit, {{"ShiftPeriod", each if NormalizeKey(_) = "NIGHT" then "NS" else NormalizeKey(_), type nullable text}, {"StartDay", fnOvertimeRosterChecksNumber, type nullable number}})
        in PeriodKeys,
    BDLandedSettings = #"IMPORT Settings Data"("BD"),
    TELandedSettings = #"IMPORT Settings Data"("TE"),
    JHRYLandedSettings = #"IMPORT Settings Data"("JH-RY"),
    BDLandedPeriods = BDLandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    TELandedPeriods = TELandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    JHRYLandedPeriods = JHRYLandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    BDPeriods = PreparePeriods(BDLandedPeriods, "BD"),
    TEPeriods = PreparePeriods(TELandedPeriods, "TE"),
    JHRYPeriods = PreparePeriods(JHRYLandedPeriods, "JH-RY"),
    CombinedPeriods = Table.Combine({BDPeriods, TEPeriods, JHRYPeriods}),
    Periods = Table.Buffer(CombinedPeriods),
    // Validate each unit's three period boundaries once, rather than for every shift.
    PeriodGroups = Table.Group(Periods, {"SourceUnit"}, {{"Bounds", each
        let AM = Table.SelectRows(_, each [ShiftPeriod] = "AM"), PM = Table.SelectRows(_, each [ShiftPeriod] = "PM"), NS = Table.SelectRows(_, each [ShiftPeriod] = "NS"),
            Valid = Table.RowCount(AM) = 1 and Table.RowCount(PM) = 1 and Table.RowCount(NS) = 1
        in [Valid = Valid, AMStart = if Valid then AM{0}[StartDay] else null, PMStart = if Valid then PM{0}[StartDay] else null, NSStart = if Valid then NS{0}[StartDay] else null], type record}}),
    PeriodBounds = Record.FromList(PeriodGroups[Bounds], PeriodGroups[SourceUnit]),
    RoleColumns = Table.SelectColumns(#"LINK OvertimeRosterChecks Roles", {"Roster Roles", "Role Group"}),
    RoleKeyRows = Table.AddColumn(RoleColumns, "RosterRoleKey", each NormalizeKey([Roster Roles]), type nullable text),
    // These small lookup snapshots contain only the mapping fields used by the joins.
    RoleKeys = Table.Buffer(RoleKeyRows),
    FacilitySource = #"LINK OvertimeRosterChecks Facilities",
    // SharePoint exposes the facility-name column by its current display name, Location.
    FacilityColumns = Table.SelectColumns(FacilitySource, {"Location", if Table.HasColumns(FacilitySource, {"Facility-Abbrev"}) then "Facility-Abbrev" else "field_1"}),
    FacilityNames = if Table.HasColumns(FacilityColumns, {"Facility-Abbrev"}) then FacilityColumns else Table.RenameColumns(FacilityColumns, {{"field_1", "Facility-Abbrev"}}),
    FacilityKeyRows = Table.AddColumn(FacilityNames, "LocationKey", each NormalizeKey([Location]), type nullable text),
    FacilityKeys = Table.Buffer(FacilityKeyRows),
    // Reuse the already narrowed roster across coverage, aliases and the main shift pipeline.
    SourceRows = Table.Buffer(OvertimeRosterChecks_RosterPrepare),
    UnitCoverage = Table.Group(SourceRows, {"SourceUnit"}, {
        {"Starts", each List.Distinct(List.RemoveNulls([CoverageStart])), type list},
        {"Ends", each List.Distinct(List.RemoveNulls([CoverageEnd])), type list}}),
    DeclaredUnits = Table.SelectRows(UnitCoverage, each List.Count([Starts]) = 1 and List.Count([Ends]) = 1 and [Starts]{0} < [Ends]{0}),
    AllUnitsDeclared = List.Sort(DeclaredUnits[SourceUnit]) = List.Sort({"BD", "TE", "JH-RY"}),
    CoverageStartAll = if AllUnitsDeclared then List.Max(List.Transform(DeclaredUnits[Starts], each _{0})) else null,
    CoverageEndAll = if AllUnitsDeclared then List.Min(List.Transform(DeclaredUnits[Ends], each _{0})) else null,
    // Enrich exact exported copies once, retaining all raw rows for coverage and alias checks.
    RawDuplicateColumns = List.RemoveItems(Table.ColumnNames(SourceRows), {"SourceUnit", "SourceRow", "SourceFile", "SourceRecordID"}),
    RawCopyMetadata = Table.Group(SourceRows, RawDuplicateColumns, {
        {"SourceCopyCount", each Table.RowCount(_), Int64.Type},
        {"SourceCopyUnits", each Text.Combine(List.Sort(List.Distinct([SourceUnit])), "|"), type text}}),
    // Buffer the provenance order so duplicate removal retains the same canonical source row.
    RawSortedRows = Table.Buffer(Table.Sort(SourceRows, {{"SourceUnit", Order.Ascending}, {"SourceRow", Order.Ascending}})),
    CanonicalRawRows = Table.Distinct(RawSortedRows, RawDuplicateColumns),
    RawCopyJoin = Table.NestedJoin(CanonicalRawRows, RawDuplicateColumns, RawCopyMetadata, RawDuplicateColumns, "RawCopies", JoinKind.LeftOuter),
    CanonicalRawWithCopies = Table.ExpandTableColumn(RawCopyJoin, "RawCopies", {"SourceCopyCount", "SourceCopyUnits"}),
    SourceRoleKeys = Table.AddColumn(CanonicalRawWithCopies, "RosterRoleKey", each NormalizeKey([SourceRole]), type nullable text),
    RoleJoin = Table.NestedJoin(SourceRoleKeys, {"RosterRoleKey"}, RoleKeys, {"RosterRoleKey"}, "RoleMatches", JoinKind.LeftOuter),
    MappedRole = Table.AddColumn(RoleJoin, "Role", each if Table.RowCount([RoleMatches]) = 1 then NormalizeKey([RoleMatches]{0}[Role Group]) else null, type nullable text),
    LocationKeys = Table.AddColumn(MappedRole, "LocationKey", each NormalizeKey([Location]), type nullable text),
    FacilityJoin = Table.NestedJoin(LocationKeys, {"LocationKey"}, FacilityKeys, {"LocationKey"}, "FacilityMatches", JoinKind.LeftOuter),
    Facility = Table.AddColumn(FacilityJoin, "Facility", each if Table.RowCount([FacilityMatches]) = 1 then fnOvertimeRosterChecksText([FacilityMatches]{0}[#"Facility-Abbrev"]) else null, type nullable text),
    ExportUnit = Table.RenameColumns(Facility, {{"SourceUnit", "SourceExportUnit"}}),
    PhysicalUnit = Table.AddColumn(ExportUnit, "SourceUnit", each [Facility], type nullable text),
    WorkerJoin = Table.NestedJoin(PhysicalUnit, {"WorkerEmployeeID"}, OvertimeRosterChecks_WorkerPrepare, {"WorkerEmployeeID"}, "WorkerMatches", JoinKind.LeftOuter),
    ExpandedWorker = Table.ExpandTableColumn(WorkerJoin, "WorkerMatches", {"WorkerMatchCount", "WorkerEmploymentType", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException", "Start Date", "Termination Date"}),
    EmployeeID = Table.AddColumn(ExpandedWorker, "EmployeeID", each if [WorkerMatchCount] = 1 then [WorkerEmployeeID] else null, type nullable text),
    AliasPairs = Table.SelectColumns(SourceRows, {"EmployerKey", "AllocationCode", "WorkerEmployeeID"}),
    DistinctAliasPairs = Table.Distinct(AliasPairs),
    AliasCounts = Table.Group(DistinctAliasPairs, {"EmployerKey", "AllocationCode"}, {{"AliasWorkerCount", each List.Count(List.Distinct(List.RemoveNulls([WorkerEmployeeID]))), Int64.Type}}),
    AliasJoin = Table.NestedJoin(EmployeeID, {"EmployerKey", "AllocationCode"}, AliasCounts, {"EmployerKey", "AllocationCode"}, "AliasCounts", JoinKind.LeftOuter),
    AliasCount = Table.ExpandTableColumn(AliasJoin, "AliasCounts", {"AliasWorkerCount"}),
    IdentityStatus = Table.AddColumn(AliasCount, "IdentityStatus", each if [EmployeeID] <> null and [AllocationCode] <> null and [EmployerKey] <> null and [Facility] <> null and [AliasWorkerCount] = 1 then "Verified" else "Unknown", type text),
    SourceShiftID = Table.AddColumn(IdentityStatus, "SourceShiftID", each fnOvertimeRosterChecksKey({if [EmployerKey] = null then null else Text.Upper([EmployerKey]), if [EmployeeID] <> null then [EmployeeID] else [WorkerEmployeeID], [Facility], [ShiftStart], [ShiftFinish], NormalizeKey([SourceRole])}), type text),
    EmployeeKey = Table.AddColumn(SourceShiftID, "EmployeeKey", each if [IdentityStatus] = "Verified" then fnOvertimeRosterChecksKey({Text.Upper([EmployerKey]), [EmployeeID]}) else "Unknown:" & [SourceRecordID], type text),
    EmploymentType = Table.AddColumn(EmployeeKey, "EmploymentType", each if [WorkerEmploymentType] <> null then [WorkerEmploymentType] else [SourceEmploymentType], type nullable text),
    ResolvedStream = Table.AddColumn(EmploymentType, "ResolvedAgreementStream", each if [AgreementStream] <> null then [AgreementStream] else [WorkerAgreementStream], type nullable text),
    ResolvedGrade = Table.AddColumn(ResolvedStream, "ResolvedGrade", each if [Grade] <> null then [Grade] else [WorkerGrade], type nullable text),
    ResolvedExempt = Table.AddColumn(ResolvedGrade, "ResolvedOvertimeExempt", each if [OvertimeExempt] <> null then [OvertimeExempt] else [WorkerOvertimeExempt], type nullable logical),
    ResolvedBroken = Table.AddColumn(ResolvedExempt, "ResolvedBrokenShiftException", each if [BrokenShiftException] <> null then [BrokenShiftException] else [WorkerBrokenShiftException], type nullable logical),
    RemovedSourcePolicies = Table.RemoveColumns(ResolvedBroken, {"AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}),
    NamedPolicies = Table.RenameColumns(RemovedSourcePolicies, {{"ResolvedAgreementStream", "AgreementStream"}, {"ResolvedGrade", "Grade"}, {"ResolvedOvertimeExempt", "OvertimeExempt"}, {"ResolvedBrokenShiftException", "BrokenShiftException"}}),
    PolicyStatus = Table.AddColumn(NamedPolicies, "PolicyStatus", each
        if ([AgreementStream] <> null and [WorkerAgreementStream] <> null and NormalizeKey([AgreementStream]) <> NormalizeKey([WorkerAgreementStream]))
            or ([Grade] <> null and [WorkerGrade] <> null and NormalizeKey([Grade]) <> NormalizeKey([WorkerGrade]))
            or ([OvertimeExempt] <> null and [WorkerOvertimeExempt] <> null and [OvertimeExempt] <> [WorkerOvertimeExempt])
            or ([BrokenShiftException] <> null and [WorkerBrokenShiftException] <> null and [BrokenShiftException] <> [WorkerBrokenShiftException]) then "Unknown" else "Known", type text),
    PublicationStatus = Table.AddColumn(PolicyStatus, "PublicationStatus", each if [Published] = true then "Known" else "Unknown", type text),
    DutyStatus = Table.AddColumn(PublicationStatus, "DutyStatus", each
        let StatusValue = NormalizeKey([Status]), Cancelled = StatusValue <> null and (Text.Contains(StatusValue, "CANCEL") or Text.Contains(StatusValue, "DRAFT")),
            ValidSpan = [ShiftStart] <> null and [ShiftFinish] <> null and [ShiftFinish] > [ShiftStart]
        in if [PublicationStatus] = "Known" and [Non Attended] = false and not Cancelled and ValidSpan then "Known" else "Unknown", type text),
    RosterScopeStatus = Table.AddColumn(DutyStatus, "RosterScopeStatus", each if [Role] <> null and not List.Contains({"AIN", "EN", "RN", "CS"}, [Role]) then "Outside scope"
        else if [IdentityStatus] = "Verified" and [DutyStatus] = "Known" and List.Contains({"AIN", "EN", "RN", "CS"}, [Role])
            and List.Contains({"PERMANENTFULLTIME", "PERMANENTPARTTIME", "FULLTIME", "PARTTIME", "FIXEDTERM", "CASUAL"}, NormalizeKey([EmploymentType])) then "Known" else "Unknown", type text),
    EligibilityStatus = Table.AddColumn(RosterScopeStatus, "EligibilityStatus", each
        if [Role] <> null and not List.Contains({"AIN", "EN", "RN", "CS"}, [Role]) then "Outside scope"
        else if [PolicyStatus] <> "Known" then "Unknown"
        else if [OvertimeExempt] = true then "Exempt"
        else if [IdentityStatus] <> "Verified" or [SourceDate] = null or [Role] = null or [Grade] = null or [OvertimeExempt] = null
            or not List.Contains({"PERSONALCARE", "CLINICAL", "HOSPITALITY"}, NormalizeKey([AgreementStream]))
            or (List.Contains({"AIN", "EN", "RN"}, [Role]) and not List.Contains({"PERSONALCARE", "CLINICAL"}, NormalizeKey([AgreementStream])))
            or ([Role] = "CS" and not List.Contains({"PERSONALCARE", "HOSPITALITY"}, NormalizeKey([AgreementStream])))
            or not List.Contains({"PERMANENTFULLTIME", "PERMANENTPARTTIME", "FULLTIME", "PARTTIME", "FIXEDTERM", "CASUAL"}, NormalizeKey([EmploymentType])) then "Unknown"
        else if ([Start Date] <> null and [SourceDate] < [Start Date]) or ([Termination Date] <> null and [SourceDate] > [Termination Date]) then "Unknown"
        else "Eligible", type text),
    DailyHoursLimit = Table.AddColumn(EligibilityStatus, "DailyHoursLimit", each if List.Contains({"AIN", "EN", "RN"}, [Role]) then OvertimeRosterChecks_DailyHoursLimit
        else if [Role] = "CS" and NormalizeKey([AgreementStream]) = "PERSONALCARE" then OvertimeRosterChecks_CSPersonalCareHoursLimit
        else if [Role] = "CS" and NormalizeKey([AgreementStream]) = "HOSPITALITY" then OvertimeRosterChecks_CSHospitalityHoursLimit else null, type nullable number),
    // Discard nested lookup tables and resolved policy helpers before grouping physical shifts.
    RemovedNestedInputs = Table.RemoveColumns(DailyHoursLimit, {"RoleMatches", "FacilityMatches", "RosterRoleKey", "LocationKey", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException", "WorkerEmploymentType"}),
    // Source-row and filename changes do not change duty identity. Identical exported copies count once.
    PhysicalFactColumns = {"WorkerEmployeeID", "AllocationCode", "EmployerKey", "SourceName", "SourceRole", "Location", "ShiftStart", "ShiftFinish", "BreakMinutes", "NetRosterHours", "SourceShiftType", "SourceEmploymentType", "Non Attended", "Status", "Published", "MealAvailabilityRequired", "MealStartDateTime", "MealFinishDateTime", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    PhysicalGroups = Table.Group(RemovedNestedInputs, {"SourceShiftID"}, {
        {"SourceRecordCount", each List.Sum([SourceCopyCount]), Int64.Type},
        {"SourceFactCount", each let FactColumns = Table.SelectColumns(_, PhysicalFactColumns), DistinctSourceFacts = Table.Distinct(FactColumns) in Table.RowCount(DistinctSourceFacts), Int64.Type},
        {"SourceUnits", each Text.Combine(List.Sort(List.Distinct(List.Combine(List.Transform([SourceCopyUnits], each Text.Split(_, "|"))))), "|"), type text},
        {"SelectedSource", each let OrderedSources = Table.Sort(_, {{"SourceExportUnit", Order.Ascending}, {"SourceRow", Order.Ascending}}) in Table.First(OrderedSources), type record}}),
    PhysicalColumns = List.RemoveItems(Table.ColumnNames(RemovedNestedInputs), {"SourceShiftID", "SourceCopyCount", "SourceCopyUnits"}),
    PhysicalShifts = Table.ExpandRecordColumn(PhysicalGroups, "SelectedSource", PhysicalColumns),
    ShiftMatchKey = Table.AddColumn(PhysicalShifts, "ShiftMatchKey", each fnOvertimeRosterChecksKey({[SourceUnit], [AllocationCode], [Role], [ShiftStart], [ShiftFinish]}), type text),
    // Materialise scalar physical-shift facts once before the count-and-join branch.
    BufferedMatchShifts = Table.Buffer(ShiftMatchKey),
    MatchCounts = Table.Group(BufferedMatchShifts, {"ShiftMatchKey"}, {{"MatchKeyCount", each Table.RowCount(_), Int64.Type}}),
    MatchJoin = Table.NestedJoin(BufferedMatchShifts, {"ShiftMatchKey"}, MatchCounts, {"ShiftMatchKey"}, "MatchCount", JoinKind.LeftOuter),
    MatchCount = Table.ExpandTableColumn(MatchJoin, "MatchCount", {"MatchKeyCount"}),
    MatchStatus = Table.AddColumn(MatchCount, "MatchStatus", each if [MatchKeyCount] = 1 and [SourceFactCount] = 1 and [IdentityStatus] = "Verified" and [AllocationCode] <> null and [Role] <> null and [ShiftStart] <> null and [ShiftFinish] <> null then "Unique" else "Unknown", type text),
    NetHoursStatus = Table.AddColumn(MatchStatus, "NetHoursStatus", each if [SourceFactCount] <> 1 or [NetRosterHours] = null or [NetRosterHours] < 0 or [GrossHours] = null or [GrossHours] <= 0
        or [BreakMinutes] = null or [BreakMinutes] < 0 or [NetRosterHours] > [GrossHours] + 0.0000001 then "Unknown" else "Known", type text),
    MealTreatmentStatus = Table.AddColumn(NetHoursStatus, "MealTreatmentStatus", each if [NetHoursStatus] <> "Known" then "Unknown"
        else if Number.Abs([GrossHours] - [NetRosterHours]) < 0.0000001 then "Known"
        else if [MealAvailabilityRequired] <> null and Number.Abs([GrossHours] - [NetRosterHours] - [BreakMinutes] / 60) < 0.0000001 then "Known" else "Unknown", type text),
    QualifyingNetHours = Table.AddColumn(MealTreatmentStatus, "QualifyingNetHours", each if [NetHoursStatus] <> "Known" then null
        else if [MealAvailabilityRequired] = true and Number.Abs([GrossHours] - [NetRosterHours] - [BreakMinutes] / 60) < 0.0000001 then [GrossHours] else [NetRosterHours], type nullable number),
    RosterShiftPeriod = Table.AddColumn(QualifyingNetHours, "RosterShiftPeriod", each
        let Bounds = if [SourceUnit] = null then null else Record.FieldOrDefault(PeriodBounds, [SourceUnit], null),
            Valid = if Bounds = null then false else Bounds[Valid] and [ShiftStart] <> null,
            StartFraction = if [ShiftStart] = null then null else Number.From(Time.From([ShiftStart]))
        in if not Valid then null else if StartFraction >= Bounds[AMStart] and StartFraction < Bounds[PMStart] then "AM"
            else if StartFraction >= Bounds[PMStart] and StartFraction < Bounds[NSStart] then "PM" else "NS", type nullable text),
    CoverageBounds = Table.AddColumn(RosterShiftPeriod, "ExportCoverageStart", each CoverageStartAll, type nullable datetime),
    CoverageFinish = Table.AddColumn(CoverageBounds, "ExportCoverageEnd", each CoverageEndAll, type nullable datetime),
    CoverageStatus = Table.AddColumn(CoverageFinish, "CoverageStatus", each if CoverageStartAll <> null and CoverageEndAll <> null and [ShiftStart] <> null and [ShiftFinish] <> null
        and [ShiftStart] >= CoverageStartAll and [ShiftFinish] <= CoverageEndAll then "Declared" else "Unknown", type text),
    // Define the complete output schema once, rather than discover it for every employee table.
    ChronologyColumns = {{"EmployeeShiftIndex", Int64.Type}, {"PreviousSourceShiftID", type nullable text}, {"PreviousShiftFinish", type nullable datetime},
        {"PreviousDutyStatus", type text}, {"NextShiftStart", type nullable datetime}, {"GapHours", type nullable number}, {"OverlapStatus", type text}},
    EmptyChronologySource = Table.FirstN(CoverageStatus, 0),
    ChronologySchema = List.Accumulate(ChronologyColumns, EmptyChronologySource, (schema, column) => Table.AddColumn(schema, column{0}, each null, column{1})),
    ChronologyTableType = Value.Type(ChronologySchema),
    PrepareChronology = (rows as table) as table =>
        let Ordered = if Table.RowCount(rows) < 2 then rows else Table.Sort(rows, {{"ShiftStart", Order.Ascending}, {"ShiftFinish", Order.Ascending}, {"SourceShiftID", Order.Ascending}}),
            OrderedRows = List.Buffer(Table.ToRecords(Ordered)),
            Count = List.Count(OrderedRows),
            PriorMaxFinishes = if Count < 2 then {null} else List.Buffer(List.Generate(() => [Index = 0, MaxFinish = null], each [Index] < Count,
                each let CurrentFinish = OrderedRows{[Index]}[ShiftFinish], NextMax = if [MaxFinish] = null then CurrentFinish else if CurrentFinish = null then [MaxFinish] else List.Max({[MaxFinish], CurrentFinish})
                in [Index = [Index] + 1, MaxFinish = NextMax], each [MaxFinish])),
            ChronologyRows = List.Transform(List.Positions(OrderedRows), (index) =>
                let Current = OrderedRows{index}, Previous = if index = 0 then null else OrderedRows{index - 1},
                    PreviousFinish = if Previous = null then null else Previous[ShiftFinish],
                    NextStart = if index + 1 >= Count then null else OrderedRows{index + 1}[ShiftStart],
                    GapHours = if PreviousFinish = null or Current[ShiftStart] = null then null else Duration.TotalHours(Current[ShiftStart] - PreviousFinish),
                    OverlapStatus = if Current[ShiftStart] = null or Current[ShiftFinish] = null then "Unknown"
                        else if (PriorMaxFinishes{index} <> null and PriorMaxFinishes{index} > Current[ShiftStart]) or (NextStart <> null and NextStart < Current[ShiftFinish]) then "Overlap" else "Clear",
                    Derived = [EmployeeShiftIndex = index, PreviousSourceShiftID = if Previous = null then null else Previous[SourceShiftID], PreviousShiftFinish = PreviousFinish,
                        PreviousDutyStatus = if Previous = null then "Unknown" else Previous[DutyStatus], NextShiftStart = NextStart, GapHours = GapHours, OverlapStatus = OverlapStatus]
                in Record.Combine({Current, Derived})),
            ChronologyTable = Table.FromRecords(ChronologyRows, ChronologyTableType)
        in ChronologyTable,
    EmployeeGroups = Table.Group(CoverageStatus, {"EmployeeKey"}, {{"EmployeeShifts", PrepareChronology, ChronologyTableType}}),
    ChronologicalShifts = if Table.IsEmpty(CoverageStatus) then PrepareChronology(CoverageStatus) else Table.Combine(EmployeeGroups[EmployeeShifts], ChronologyTableType),
    // Local datetime alone cannot establish elapsed hours across Sydney clock-change windows.
    GapWallHours = Table.RenameColumns(ChronologicalShifts, {{"GapHours", "GapWallHours"}}),
    GapElapsedStatus = Table.AddColumn(GapWallHours, "GapElapsedStatus", each if [PreviousShiftFinish] = null or [ShiftStart] = null then "Unknown"
        else if fnOvertimeRosterChecksClockSafe(List.Min({[PreviousShiftFinish], [ShiftStart]}), List.Max({[PreviousShiftFinish], [ShiftStart]})) then "Known" else "Unknown", type text),
    GapHours = Table.AddColumn(GapElapsedStatus, "GapHours", each if [GapElapsedStatus] = "Known" then [GapWallHours] else null, type nullable number),
    GapStatus = Table.AddColumn(GapHours, "GapStatus", each if [GapHours] <> null and [DutyStatus] = "Known" and [PreviousDutyStatus] = "Known" and [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [CoverageStatus] = "Declared"
        and [PreviousShiftFinish] >= [ExportCoverageStart] and [OverlapStatus] = "Clear" then "Known" else "Unknown", type text),
    MinimumBreakFlag = Table.AddColumn(GapStatus, "MinimumBreakFlag", each if [GapStatus] = "Known" and [RosterScopeStatus] = "Known" then [GapHours] < OvertimeRosterChecks_MinimumBreakHours else null, type nullable logical),
    MinBreakStatus = Table.AddColumn(MinimumBreakFlag, "MinBreakStatus", each if [MinimumBreakFlag] = null then "Unknown" else if [MinimumBreakFlag] then "Review" else "Clear", type text),
    ShortGapReviewFlag = Table.AddColumn(MinBreakStatus, "ShortGapReviewFlag", each if [EligibilityStatus] <> "Eligible" or [GapStatus] <> "Known" then null
        else if NormalizeKey([EmploymentType]) = "CASUAL" or [BrokenShiftException] = true then false
        else if [BrokenShiftException] = null then null else [GapHours] < OvertimeRosterChecks_ShortGapReviewHours, type nullable logical),
    ShortGapStatus = Table.AddColumn(ShortGapReviewFlag, "ShortGapStatus", each if [ShortGapReviewFlag] = null then "Unknown" else if [ShortGapReviewFlag] then "Review" else "Clear", type text),
    DailyExcessHours = Table.AddColumn(ShortGapStatus, "DailyExcessHours", each if [QualifyingNetHours] = null or [DailyHoursLimit] = null then null else List.Max({0, [QualifyingNetHours] - [DailyHoursLimit]}), type nullable number),
    DailyStatus = Table.AddColumn(DailyExcessHours, "DailyStatus", each if [EligibilityStatus] = "Exempt" or [EligibilityStatus] = "Outside scope" then "Not applicable"
        else if [EligibilityStatus] <> "Eligible" or [NetHoursStatus] <> "Known" or [MealTreatmentStatus] <> "Known" or [MatchStatus] <> "Unique" or [OverlapStatus] <> "Clear" or [CoverageStatus] <> "Declared" then "Unknown"
        else if [DailyExcessHours] = null then "Unknown" else if [DailyExcessHours] > 0 then "Exceeded" else "Clear", type text),
    DailyHoursFlag = Table.AddColumn(DailyStatus, "DailyHoursFlag", each if [DailyStatus] = "Exceeded" then true else if [DailyStatus] = "Clear" then false else null, type nullable logical),
    ClockStatus = Table.AddColumn(DailyHoursFlag, "ClockStatus", each if fnOvertimeRosterChecksClockSafe([ShiftStart], [ShiftFinish]) then "Known" else "Unknown", type text)
in
    ClockStatus,
OvertimeRosterChecks_WorkSegments = let
    Source = if OvertimeRosterChecks_MinimumFortnightDutyFreeDays <= OvertimeRosterChecks_FortnightLengthDays then OvertimeRosterChecks_ShiftPrepare
        else error "MinimumFortnightDutyFreeDays cannot exceed FortnightLengthDays.",
    SegmentLists = Table.AddColumn(Source, "WorkSegments", each
        let ShiftID = [SourceShiftID], Employee = [EmployeeKey], WorkHours = [QualifyingNetHours],
            Start = [ShiftStart], Finish = [ShiftFinish], MealStart = [MealStartDateTime], MealFinish = [MealFinishDateTime],
            ClockEqualsNet = WorkHours <> null and [ClockStatus] = "Known" and [GrossHours] <> null and Number.Abs(WorkHours - [GrossHours]) < 0.0000001,
            UnpaidMealKnown = WorkHours <> null and [ClockStatus] = "Known" and [MealTreatmentStatus] = "Known" and [MealAvailabilityRequired] = false and MealStart <> null and MealFinish <> null
                and Start <> null and Finish <> null and MealStart >= Start and MealFinish <= Finish and MealFinish > MealStart,
            MealReconciles = if not UnpaidMealKnown then false else Number.Abs(Duration.TotalHours(Finish - Start) - Duration.TotalHours(MealFinish - MealStart) - WorkHours) < 0.0000001,
            FirstHours = if MealReconciles then Duration.TotalHours(MealStart - Start) else null,
            Rows = if WorkHours = null then {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = null, WorkFinish = null, NetWorkStartHours = 0, NetWorkEndHours = null, TimingStatus = "Unknown"]}
                else if ClockEqualsNet then {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = Start, WorkFinish = Finish, NetWorkStartHours = 0, NetWorkEndHours = WorkHours, TimingStatus = "Known"]}
                else if MealReconciles then {
                    [SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = Start, WorkFinish = MealStart, NetWorkStartHours = 0, NetWorkEndHours = FirstHours, TimingStatus = "Known"],
                    [SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = MealFinish, WorkFinish = Finish, NetWorkStartHours = FirstHours, NetWorkEndHours = WorkHours, TimingStatus = "Known"]}
                else {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = null, WorkFinish = null, NetWorkStartHours = 0, NetWorkEndHours = WorkHours, TimingStatus = "Unknown"]}
        in Rows, type list),
    CombinedSegments = if Table.IsEmpty(Source) then {} else List.Combine(SegmentLists[WorkSegments]),
    SegmentTable = Table.FromRecords(CombinedSegments, type table [SourceShiftID = text, EmployeeKey = text, WorkStart = nullable datetime, WorkFinish = nullable datetime, NetWorkStartHours = number, NetWorkEndHours = nullable number, TimingStatus = text]),
    NonzeroSegments = Table.SelectRows(SegmentTable, each [NetWorkEndHours] = null or [NetWorkEndHours] > [NetWorkStartHours])
in
    NonzeroSegments,
OvertimeRosterChecks_FortnightParts = let
    Source = OvertimeRosterChecks_ShiftPrepare,
    Anchor = OvertimeRosterChecks_FortnightStartDate,
    FortnightDays = if OvertimeRosterChecks_MinimumFortnightDutyFreeDays <= OvertimeRosterChecks_FortnightLengthDays then OvertimeRosterChecks_FortnightLengthDays
        else error "MinimumFortnightDutyFreeDays cannot exceed FortnightLengthDays.",
    WorkSegments = Table.Buffer(OvertimeRosterChecks_WorkSegments),
    ShiftWorkJoin = Table.NestedJoin(Source, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
    PartLists = Table.AddColumn(ShiftWorkJoin, "FortnightParts", each
        if Anchor = null or [ShiftStart] = null or [ShiftFinish] = null or [ShiftFinish] <= [ShiftStart] then {} else
        let Duty = _, ShiftID = [SourceShiftID], Start = [ShiftStart], Finish = [ShiftFinish],
            FirstIndex = Number.RoundDown(Duration.Days(Date.From(Start) - Anchor) / FortnightDays),
            LastIndex = Number.RoundDown(Duration.Days(Date.From(Finish) - Anchor) / FortnightDays),
            Indices = {FirstIndex..LastIndex},
            ShiftSegments = [ShiftWork],
            TimingKnown = not Table.IsEmpty(ShiftSegments) and List.AllTrue(List.Transform(ShiftSegments[TimingStatus], each _ = "Known")),
            Rows = List.Transform(Indices, (index) =>
                let FNStart = Date.AddDays(Anchor, index * FortnightDays), FNEnd = Date.AddDays(FNStart, FortnightDays),
                    PartStart = List.Max({Start, DateTime.From(FNStart)}), PartFinish = List.Min({Finish, DateTime.From(FNEnd)}),
                    IsWholeShift = PartStart = Start and PartFinish = Finish,
                    HoursBeforePart = if IsWholeShift then 0 else if not TimingKnown then null else List.Sum(List.Transform(Table.ToRecords(ShiftSegments), (segment) =>
                        if segment[WorkStart] >= PartStart then 0 else Duration.TotalHours(List.Min({segment[WorkFinish], PartStart}) - segment[WorkStart]))),
                    NetPartHours = if PartFinish <= PartStart then 0 else if IsWholeShift then Duty[QualifyingNetHours]
                        else if not TimingKnown then null else List.Sum(List.Transform(Table.ToRecords(ShiftSegments), (segment) =>
                            let Left = List.Max({segment[WorkStart], PartStart}), Right = List.Min({segment[WorkFinish], PartFinish})
                            in if Right > Left then Duration.TotalHours(Right - Left) else 0)),
                    PartNetEnd = if NetPartHours = null or HoursBeforePart = null then null else HoursBeforePart + NetPartHours
                in [SourceShiftID = ShiftID, EmployeeKey = Duty[EmployeeKey], EmployeeID = Duty[EmployeeID], EmployerKey = Duty[EmployerKey],
                    PayrollFortnightStart = FNStart, PayrollFortnightEnd = FNEnd, PartStart = PartStart, PartFinish = PartFinish,
                    PartNetStartHours = HoursBeforePart, PartNetEndHours = PartNetEnd, NetPartHours = NetPartHours,
                    DailyHoursLimit = Duty[DailyHoursLimit], EligibilityStatus = Duty[EligibilityStatus], RosterScopeStatus = Duty[RosterScopeStatus], DutyStatus = Duty[DutyStatus], IdentityStatus = Duty[IdentityStatus], MatchStatus = Duty[MatchStatus],
                    MealTreatmentStatus = Duty[MealTreatmentStatus], CoverageStatus = Duty[CoverageStatus], ExportCoverageStart = Duty[ExportCoverageStart], ExportCoverageEnd = Duty[ExportCoverageEnd], OverlapStatus = Duty[OverlapStatus]])
        in Rows, type list),
    CombinedParts = if Table.IsEmpty(Source) then {} else List.Combine(PartLists[FortnightParts]),
    PartTable = Table.FromRecords(CombinedParts, type table [SourceShiftID = text, EmployeeKey = text, EmployeeID = nullable text, EmployerKey = nullable text,
        PayrollFortnightStart = date, PayrollFortnightEnd = date, PartStart = datetime, PartFinish = datetime, PartNetStartHours = nullable number, PartNetEndHours = nullable number, NetPartHours = nullable number,
        DailyHoursLimit = nullable number, EligibilityStatus = text, RosterScopeStatus = text, DutyStatus = text, IdentityStatus = text, MatchStatus = text, MealTreatmentStatus = text, CoverageStatus = text,
        ExportCoverageStart = nullable datetime, ExportCoverageEnd = nullable datetime, OverlapStatus = text]),
    PositiveParts = Table.SelectRows(PartTable, each [PartFinish] > [PartStart]),
    CounterColumns = {{"FortnightHoursBefore", type nullable number}, {"FortnightHoursAfter", type nullable number}, {"FortnightExcessHours", type nullable number}},
    EmptyCounterSource = Table.FirstN(PositiveParts, 0),
    CounterSchema = List.Accumulate(CounterColumns, EmptyCounterSource, (schema, column) => Table.AddColumn(schema, column{0}, each null, column{1})),
    CounterTableType = Value.Type(CounterSchema),
    AddCounter = (parts as table) as table =>
        let Ordered = if Table.RowCount(parts) < 2 then parts else Table.Sort(parts, {{"PartStart", Order.Ascending}, {"SourceShiftID", Order.Ascending}}),
            OrderedRows = List.Buffer(Table.ToRecords(Ordered)),
            Count = List.Count(OrderedRows),
            BeforeValues = if Count < 2 then {0} else List.Buffer(List.Generate(() => [Index = 0, Before = 0], each [Index] < Count,
                each let CurrentHours = OrderedRows{[Index]}[NetPartHours], NextBefore = if [Before] = null or CurrentHours = null then null else [Before] + CurrentHours
                in [Index = [Index] + 1, Before = NextBefore], each [Before])),
            CountedRows = List.Transform(List.Positions(OrderedRows), (index) =>
                let Current = OrderedRows{index}, Before = BeforeValues{index},
                    After = if Before = null or Current[NetPartHours] = null then null else Before + Current[NetPartHours],
                    Excess = if Before = null or After = null then null
                        else List.Max({0, After - OvertimeRosterChecks_FortnightHoursLimit}) - List.Max({0, Before - OvertimeRosterChecks_FortnightHoursLimit}),
                    Derived = [FortnightHoursBefore = Before, FortnightHoursAfter = After, FortnightExcessHours = Excess]
                in Record.Combine({Current, Derived})),
            CounterTable = Table.FromRecords(CountedRows, CounterTableType)
        in CounterTable,
    EmployeeFortnights = Table.Group(PositiveParts, {"EmployeeKey", "PayrollFortnightStart"}, {{"Parts", AddCounter, CounterTableType}}),
    CountedParts = if Table.IsEmpty(PositiveParts) then AddCounter(PositiveParts) else Table.Combine(EmployeeFortnights[Parts], CounterTableType)
in
    CountedParts,
OvertimeRosterChecks_FortnightResults = let
    Parts = OvertimeRosterChecks_FortnightParts,
    DutyColumns = Table.SelectColumns(OvertimeRosterChecks_ShiftPrepare, {"EmployeeKey", "SourceDate", "ShiftStart", "ShiftFinish", "IdentityStatus", "MatchStatus", "DutyStatus", "RosterScopeStatus"}),
    Duties = Table.Buffer(DutyColumns),
    GroupedParts = Table.Group(Parts, {"EmployeeKey", "PayrollFortnightStart", "PayrollFortnightEnd"}, {{"Parts", each _, type table}}),
    EmployerKey = Table.AddColumn(GroupedParts, "EmployerKey", each [Parts]{0}[EmployerKey], type nullable text),
    EmployeeID = Table.AddColumn(EmployerKey, "EmployeeID", each [Parts]{0}[EmployeeID], type nullable text),
    NetRosterHours = Table.AddColumn(EmployeeID, "NetRosterHours", each if List.NonNullCount([Parts][NetPartHours]) = Table.RowCount([Parts]) then List.Sum([Parts][NetPartHours]) else null, type nullable number),
    ExcessHours = Table.AddColumn(NetRosterHours, "FortnightExcessHours", each if [NetRosterHours] = null then null else List.Max({0, [NetRosterHours] - OvertimeRosterChecks_FortnightHoursLimit}), type nullable number),
    CoverageStatus = Table.AddColumn(ExcessHours, "CoverageStatus", each
        let Row = [Parts]{0}, Start = DateTime.From([PayrollFortnightStart]), Finish = DateTime.From([PayrollFortnightEnd])
        in if Row[ExportCoverageStart] <> null and Row[ExportCoverageEnd] <> null and Row[ExportCoverageStart] <= Start and Row[ExportCoverageEnd] >= Finish then "Declared" else "Unknown", type text),
    EmployeeDutyJoin = Table.NestedJoin(CoverageStatus, {"EmployeeKey"}, Duties, {"EmployeeKey"}, "EmployeeDutyRows", JoinKind.LeftOuter),
    EmployeeDuty = Table.AddColumn(EmployeeDutyJoin, "Duty", each
        let StartDate = [PayrollFortnightStart], FinishDate = [PayrollFortnightEnd], Start = DateTime.From(StartDate), Finish = DateTime.From(FinishDate)
        in Table.SelectRows([EmployeeDutyRows], each
            let ValidSpan = [ShiftStart] <> null and [ShiftFinish] <> null and [ShiftFinish] > [ShiftStart]
            in if ValidSpan then [ShiftStart] < Finish and [ShiftFinish] > Start
                else if [SourceDate] <> null then [SourceDate] >= StartDate and [SourceDate] < FinishDate
                else true), type table),
    DutyKnown = Table.AddColumn(EmployeeDuty, "DutyCoverageKnown", each [CoverageStatus] = "Declared" and not Table.IsEmpty([Duty])
        and List.AllTrue(List.Transform(Table.ToRecords([Duty]), each [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [DutyStatus] = "Known")), type logical),
    DutyFreeDays = Table.AddColumn(DutyKnown, "DutyFreeDays", each if not [DutyCoverageKnown] then null else
        let Days = List.Dates([PayrollFortnightStart], Duration.Days([PayrollFortnightEnd] - [PayrollFortnightStart]), #duration(1, 0, 0, 0)), Duty = [Duty],
            FreeDays = List.Select(Days, (day) =>
                let Start = DateTime.From(day), Finish = DateTime.From(Date.AddDays(day, 1)), Overlaps = Table.SelectRows(Duty, each [ShiftStart] < Finish and [ShiftFinish] > Start)
                in Table.IsEmpty(Overlaps))
        in List.Count(FreeDays), type nullable number),
    DaysShort = Table.AddColumn(DutyFreeDays, "DaysShort", each if [DutyFreeDays] = null then null else List.Max({0, OvertimeRosterChecks_MinimumFortnightDutyFreeDays - [DutyFreeDays]}), type nullable number),
    DaysOffFlag = Table.AddColumn(DaysShort, "DaysOffFlag", each if [DaysShort] = null or not List.AnyTrue(List.Transform([Duty][RosterScopeStatus], each _ = "Known")) then null else [DaysShort] > 0, type nullable logical),
    DaysOffStatus = Table.AddColumn(DaysOffFlag, "DaysOffStatus", each if [DaysOffFlag] = null then "Unknown" else if [DaysOffFlag] then "Review" else "Clear", type text),
    FortnightStatus = Table.AddColumn(DaysOffStatus, "FortnightStatus", each if not [DutyCoverageKnown] or [NetRosterHours] = null
        or not List.AllTrue(List.Transform(Table.ToRecords([Parts]), each [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [MealTreatmentStatus] = "Known" and [DutyStatus] = "Known" and [OverlapStatus] = "Clear")) then "Unknown"
        else if [FortnightExcessHours] > 0 then "Exceeded" else "Clear", type text),
    PublishedColumns = Table.SelectColumns(FortnightStatus, {"EmployeeKey", "EmployeeID", "EmployerKey", "PayrollFortnightStart", "PayrollFortnightEnd", "NetRosterHours", "FortnightExcessHours", "DutyFreeDays", "DaysShort", "DaysOffFlag", "CoverageStatus", "FortnightStatus", "DaysOffStatus"})
in
    PublishedColumns,
OvertimeRosterChecks_ExcessSegments = let
    // Keep provenance and unrelated check columns out of the per-part excess calculation.
    ShiftColumns = Table.SelectColumns(OvertimeRosterChecks_ShiftPrepare, {"SourceShiftID", "EmployeeKey", "QualifyingNetHours", "DailyHoursLimit", "EligibilityStatus", "MealTreatmentStatus", "MatchStatus", "DutyStatus", "CoverageStatus", "OverlapStatus", "DailyStatus"}),
    Parts = OvertimeRosterChecks_FortnightParts,
    // Materialise only scalar keys/status before nested joins revisit each shift's parts.
    FNStatusColumns = Table.SelectColumns(OvertimeRosterChecks_FortnightResults, {"EmployeeKey", "PayrollFortnightStart", "FortnightStatus"}),
    BufferedFNStatus = Table.Buffer(FNStatusColumns),
    FNStatusJoin = Table.NestedJoin(Parts, {"EmployeeKey", "PayrollFortnightStart"}, BufferedFNStatus, {"EmployeeKey", "PayrollFortnightStart"}, "FNStatus", JoinKind.LeftOuter),
    FNStatus = Table.ExpandTableColumn(FNStatusJoin, "FNStatus", {"FortnightStatus"}, {"PartFortnightStatus"}),
    PartsColumns = Table.SelectColumns(FNStatus, {"SourceShiftID", "EmployeeKey", "PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus"}),
    BufferedParts = Table.Buffer(PartsColumns),
    // Check position completeness once per shift, before flat expansion. If any part
    // is unresolved, the whole shift keeps one daily-position fallback row.
    PositionCompleteness = Table.Group(BufferedParts, {"SourceShiftID"}, {{"HasKnownPositions", each
        let Count = Table.RowCount(_)
        in List.NonNullCount([PartNetStartHours]) = Count and List.NonNullCount([PartNetEndHours]) = Count, type logical}}),
    CompletePositionKeys = Table.SelectColumns(Table.SelectRows(PositionCompleteness, each [HasKnownPositions]), {"SourceShiftID"}),
    CompletePartsJoin = Table.NestedJoin(BufferedParts, {"SourceShiftID"}, CompletePositionKeys, {"SourceShiftID"}, "CompletePosition", JoinKind.Inner),
    CompleteParts = Table.RemoveColumns(CompletePartsJoin, {"CompletePosition"}),
    MarkedParts = Table.AddColumn(CompleteParts, "HasPositionPart", each true, type logical),
    ShiftPartJoin = Table.NestedJoin(ShiftColumns, {"SourceShiftID"}, MarkedParts, {"SourceShiftID"}, "Parts", JoinKind.LeftOuter),
    ExpandedPositionParts = Table.ExpandTableColumn(ShiftPartJoin, "Parts", {"PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus", "HasPositionPart"}),
    PositionStart = Table.AddColumn(ExpandedPositionParts, "PositionStart", each if [HasPositionPart] = true then [PartNetStartHours] else 0, type number),
    PositionEnd = Table.AddColumn(PositionStart, "PositionEnd", each if [HasPositionPart] = true then [PartNetEndHours] else [QualifyingNetHours], type nullable number),
    PositionStatus = Table.AddColumn(PositionEnd, "PositionStatus", each if [HasPositionPart] = true then [PartFortnightStatus] else "Unknown", type nullable text),
    RemovedPositionHelpers = Table.RemoveColumns(PositionStatus, {"PartNetStartHours", "PartNetEndHours", "PartFortnightStatus", "HasPositionPart"}),
    NamedPositions = Table.RenameColumns(RemovedPositionHelpers, {{"PositionStart", "PartNetStartHours"}, {"PositionEnd", "PartNetEndHours"}, {"PositionStatus", "PartFortnightStatus"}}),
    ExpandedParts = Table.Buffer(NamedPositions),
    WorkSegments = Table.Buffer(OvertimeRosterChecks_WorkSegments),
    WorkJoin = Table.NestedJoin(ExpandedParts, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
    SegmentLists = Table.AddColumn(WorkJoin, "ExcessSegments", each
        let Duty = _, ShiftID = [SourceShiftID], Left = [PartNetStartHours], Right = [PartNetEndHours], DailyLimit = [DailyHoursLimit],
            FNBefore = [FortnightHoursBefore], FNLimitPosition = if FNBefore = null or Left = null then null else Left + List.Max({0, OvertimeRosterChecks_FortnightHoursLimit - FNBefore}),
            ThisWork = [ShiftWork],
            CandidateCuts = List.Combine({{Left, Right, DailyLimit, FNLimitPosition}, ThisWork[NetWorkStartHours], ThisWork[NetWorkEndHours]}),
            Cuts = if Left = null or Right = null then {} else List.Sort(List.Distinct(List.Select(List.RemoveNulls(CandidateCuts), each _ >= Left and _ <= Right))),
            Pairs = if List.Count(Cuts) < 2 then {} else List.Transform({0..List.Count(Cuts) - 2}, each [Left = Cuts{_}, Right = Cuts{_ + 1}]),
            LocateClock = (position as number, isFinish as logical) as nullable datetime =>
                let Matching = Table.SelectRows(ThisWork, each [TimingStatus] = "Known" and (if isFinish then position > [NetWorkStartHours] and position <= [NetWorkEndHours] else position >= [NetWorkStartHours] and position < [NetWorkEndHours])),
                    Located = if Table.RowCount(Matching) = 1 then Matching{0}[WorkStart] + #duration(0, 0, 0, (position - Matching{0}[NetWorkStartHours]) * 3600) else null
                in Located,
            Rows = List.Transform(Pairs, (pair) =>
                let Daily = DailyLimit <> null and pair[Left] >= DailyLimit, FN = FNLimitPosition <> null and pair[Left] >= FNLimitPosition,
                    Start = LocateClock(pair[Left], false), Finish = LocateClock(pair[Right], true),
                    Triggers = Text.Combine(List.RemoveNulls({if Daily then "Daily" else null, if FN then "Fortnight76" else null}), "; "),
                    Status = if Duty[EligibilityStatus] = "Eligible" and Duty[MealTreatmentStatus] = "Known" and Duty[MatchStatus] = "Unique" and Duty[DutyStatus] = "Known" and Duty[CoverageStatus] = "Declared" and Duty[OverlapStatus] = "Clear"
                        and (not Daily or Duty[DailyStatus] <> "Unknown") and (not FN or Duty[PartFortnightStatus] <> "Unknown") then "Review" else "Unknown"
                in [SourceShiftID = ShiftID, EmployeeKey = Duty[EmployeeKey], PayrollFortnightStart = Duty[PayrollFortnightStart], NetWorkStartHours = pair[Left], NetWorkEndHours = pair[Right], SegmentHours = pair[Right] - pair[Left],
                    SegmentStart = Start, SegmentFinish = Finish, DailyTrigger = Daily, FortnightTrigger = FN, TriggerSummary = Triggers, TimingStatus = if Start <> null and Finish <> null then "Known" else "Unknown", OvertimeStatus = Status])
        in List.Select(Rows, each [SegmentHours] > 0 and ([DailyTrigger] or [FortnightTrigger])), type list),
    CombinedSegments = if Table.IsEmpty(ExpandedParts) then {} else List.Combine(SegmentLists[ExcessSegments]),
    SegmentTable = Table.FromRecords(CombinedSegments, type table [SourceShiftID = text, EmployeeKey = text, PayrollFortnightStart = nullable date, NetWorkStartHours = number, NetWorkEndHours = number, SegmentHours = number,
        SegmentStart = nullable datetime, SegmentFinish = nullable datetime, DailyTrigger = logical, FortnightTrigger = logical, TriggerSummary = text, TimingStatus = text, OvertimeStatus = text])
in
    SegmentTable,
OvertimeRosterChecks_ShiftResults = let
    Source = OvertimeRosterChecks_ShiftPrepare,
    Parts = OvertimeRosterChecks_FortnightParts,
    FNJoin = Table.NestedJoin(Parts, {"EmployeeKey", "PayrollFortnightStart"}, OvertimeRosterChecks_FortnightResults, {"EmployeeKey", "PayrollFortnightStart"}, "FN", JoinKind.LeftOuter),
    ExpandedFN = Table.ExpandTableColumn(FNJoin, "FN", {"DutyFreeDays", "DaysShort", "DaysOffFlag", "DaysOffStatus", "FortnightStatus"}),
    ShiftPartJoin = Table.NestedJoin(Source, {"SourceShiftID"}, ExpandedFN, {"SourceShiftID"}, "ShiftParts", JoinKind.LeftOuter),
    PayrollFortnightCount = Table.AddColumn(ShiftPartJoin, "PayrollFortnightCount", each Table.RowCount([ShiftParts]), Int64.Type),
    PayrollFortnightStart = Table.AddColumn(PayrollFortnightCount, "PayrollFortnightStart", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[PayrollFortnightStart] else null, type nullable date),
    PayrollFortnightEnd = Table.AddColumn(PayrollFortnightStart, "PayrollFortnightEnd", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[PayrollFortnightEnd] else null, type nullable date),
    FortnightHoursBefore = Table.AddColumn(PayrollFortnightEnd, "FortnightHoursBefore", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[FortnightHoursBefore] else null, type nullable number),
    FortnightHoursAfter = Table.AddColumn(FortnightHoursBefore, "FortnightHoursAfter", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[FortnightHoursAfter] else null, type nullable number),
    FortnightExcessHours = Table.AddColumn(FortnightHoursAfter, "FortnightExcessHours", each if [PayrollFortnightCount] = 0 or List.NonNullCount([ShiftParts][FortnightExcessHours]) <> [PayrollFortnightCount] then null else List.Sum([ShiftParts][FortnightExcessHours]), type nullable number),
    FortnightStatus = Table.AddColumn(FortnightExcessHours, "FortnightStatus", each if [PayrollFortnightCount] = 0 or List.Contains([ShiftParts][FortnightStatus], "Unknown") then "Unknown"
        else if [FortnightExcessHours] > 0 then "Exceeded" else "Clear", type text),
    FortnightHoursFlag = Table.AddColumn(FortnightStatus, "FortnightHoursFlag", each if [EligibilityStatus] <> "Eligible" or [FortnightStatus] = "Unknown" then null else [FortnightExcessHours] > 0, type nullable logical),
    DutyFreeDays = Table.AddColumn(FortnightHoursFlag, "DutyFreeDays", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[DutyFreeDays] else null, type nullable number),
    DaysShort = Table.AddColumn(DutyFreeDays, "DaysShort", each if [PayrollFortnightCount] = 0 or List.NonNullCount([ShiftParts][DaysShort]) <> [PayrollFortnightCount] then null else List.Max([ShiftParts][DaysShort]), type nullable number),
    DaysOffFlag = Table.AddColumn(DaysShort, "DaysOffFlag", each if [PayrollFortnightCount] = 0 then null else if List.Contains([ShiftParts][DaysOffFlag], true) then true
        else if List.Contains([ShiftParts][DaysOffFlag], null) then null else false, type nullable logical),
    DaysOffStatus = Table.AddColumn(DaysOffFlag, "DaysOffStatus", each if [DaysOffFlag] = null then "Unknown" else if [DaysOffFlag] then "Review" else "Clear", type text),
    DaysOffOvertimeReviewFlag = Table.AddColumn(DaysOffStatus, "DaysOffOvertimeReviewFlag", each if [EligibilityStatus] <> "Eligible" or [DaysOffFlag] = null then null
        else if Text.Upper(Text.Remove([EmploymentType], {" ", "-", "_"})) = "CASUAL" then false else [DaysOffFlag], type nullable logical),
    WorkTiming = Table.Group(OvertimeRosterChecks_WorkSegments, {"SourceShiftID"}, {{"MealTimingStatus", each if List.AllTrue(List.Transform([TimingStatus], each _ = "Known")) then "Known" else "Unknown", type text}}),
    WorkTimingJoin = Table.NestedJoin(DaysOffOvertimeReviewFlag, {"SourceShiftID"}, WorkTiming, {"SourceShiftID"}, "WorkTiming", JoinKind.LeftOuter),
    MealTimingStatus = Table.ExpandTableColumn(WorkTimingJoin, "WorkTiming", {"MealTimingStatus"}),
    TimingStatusDefaults = Table.TransformColumns(MealTimingStatus, {{"MealTimingStatus", each if _ = null then "Unknown" else _, type text}}),
    OvertimeJoin = Table.NestedJoin(TimingStatusDefaults, {"SourceShiftID"}, OvertimeRosterChecks_ExcessSegments, {"SourceShiftID"}, "ExcessSegments", JoinKind.LeftOuter),
    OvertimeHours = Table.AddColumn(OvertimeJoin, "OvertimeHours", each if [QualifyingNetHours] = null or [DailyHoursLimit] = null or [FortnightStatus] = "Unknown" then null else List.Sum({0} & [ExcessSegments][SegmentHours]), type nullable number),
    OvertimeStatus = Table.AddColumn(OvertimeHours, "OvertimeStatus", each if [EligibilityStatus] = "Exempt" or [EligibilityStatus] = "Outside scope" then "Not applicable"
        else if [EligibilityStatus] <> "Eligible" or [DailyStatus] = "Unknown" or [FortnightStatus] = "Unknown" then "Unknown"
        else if [OvertimeHours] > 0 or [ShortGapReviewFlag] = true or [DaysOffOvertimeReviewFlag] = true then "Review" else "Clear", type text),
    AnyRuleFlag = Table.AddColumn(OvertimeStatus, "AnyRuleFlag", each
        let Flags = {[DailyHoursFlag], [MinimumBreakFlag], [ShortGapReviewFlag], [DaysOffFlag], [FortnightHoursFlag]}
        in if List.Contains(Flags, true) then true else if List.Contains(Flags, null) then null else false, type nullable logical),
    TriggerSummary = Table.AddColumn(AnyRuleFlag, "TriggerSummary", each Text.Combine(List.RemoveNulls({if [DailyHoursFlag] = true then "Daily" else null,
        if [MinimumBreakFlag] = true then "MinimumBreak" else null, if [ShortGapReviewFlag] = true then "ShortGapReview" else null,
        if [DaysOffFlag] = true then "DaysOffRoster" else null, if [DaysOffOvertimeReviewFlag] = true then "DaysOffOvertimeReview" else null,
        if [FortnightHoursFlag] = true then "Fortnight76" else null}), "; "), type text),
    ResultReasons = Table.AddColumn(TriggerSummary, "ResultReasons", each Text.Combine(List.RemoveNulls({
        if [IdentityStatus] <> "Verified" then "Employee, allocation-alias or facility identity unresolved" else null,
        if [MatchStatus] <> "Unique" then "Shift match or conflicting duplicate unresolved" else null,
        if [SourceRecordCount] > 1 and [SourceFactCount] = 1 then "Identical published export copies collapsed" else null,
        if [PublicationStatus] <> "Known" then "Publication flag missing" else null,
        if [DutyStatus] <> "Known" then "Duty status unresolved" else null,
        if [PolicyStatus] <> "Known" then "Roster and worker policy facts conflict" else null,
        if [CoverageStatus] <> "Declared" then "Export coverage is not declared across all configured units" else null,
        if [AgreementStream] = null then "Agreement stream missing" else null,
        if [Grade] = null then "Grade missing" else null,
        if [OvertimeExempt] = null then "Overtime exemption unresolved" else null,
        if [BrokenShiftException] = null then "Broken-shift exception unresolved" else null,
        if [NetHoursStatus] <> "Known" then "Net roster hours or shift span invalid or conflicting" else null,
        if [MealTreatmentStatus] <> "Known" then "Meal availability treatment unresolved" else null,
        if [MealTimingStatus] <> "Known" then "Exact meal or overtime clock timing unresolved" else null,
        if [PreviousSourceShiftID] = null then "First supplied roster shift has no preceding shift" else null,
        if [GapElapsedStatus] <> "Known" and [PreviousSourceShiftID] <> null then "Elapsed gap cannot be established across local clock ambiguity" else null,
        if [OverlapStatus] = "Overlap" then "Overlapping roster shifts require reconciliation" else null,
        if OvertimeRosterChecks_FortnightStartDate = null then "Fortnight start is unknown because Settings permutation dates are empty" else null,
        if [FortnightStatus] = "Unknown" and OvertimeRosterChecks_FortnightStartDate <> null then "Roster-based fortnight has incomplete coverage or unresolved qualifying hours" else null}), "; "), type text),
    PublishedColumns = Table.SelectColumns(ResultReasons, {"SourceShiftID", "SourceUnit", "SourceExportUnit", "SourceUnits", "SourceFile", "SourceRow", "SourceName", "SourceRole", "SourceDate", "EmployerKey", "WorkerEmployeeID", "AllocationCode", "EmployeeID", "EmployeeKey", "Facility", "Role", "ShiftStart", "ShiftFinish", "RosterShiftPeriod", "GrossHours", "BreakMinutes", "NetRosterHours", "QualifyingNetHours", "EmploymentType", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException", "IdentityStatus", "MatchStatus", "PolicyStatus", "PublicationStatus", "DutyStatus", "RosterScopeStatus", "EligibilityStatus", "CoverageStatus", "NetHoursStatus", "MealTreatmentStatus", "MealTimingStatus", "ClockStatus", "PreviousSourceShiftID", "PreviousShiftFinish", "GapWallHours", "GapHours", "GapElapsedStatus", "GapStatus", "OverlapStatus", "MinimumBreakFlag", "MinBreakStatus", "ShortGapReviewFlag", "ShortGapStatus", "DailyHoursLimit", "DailyExcessHours", "DailyHoursFlag", "DailyStatus", "PayrollFortnightCount", "PayrollFortnightStart", "PayrollFortnightEnd", "FortnightHoursBefore", "FortnightHoursAfter", "FortnightExcessHours", "FortnightHoursFlag", "FortnightStatus", "DutyFreeDays", "DaysShort", "DaysOffFlag", "DaysOffStatus", "DaysOffOvertimeReviewFlag", "OvertimeHours", "OvertimeStatus", "AnyRuleFlag", "TriggerSummary", "ResultReasons"}),
    OrderedOutput = Table.Sort(PublishedColumns, {{"SourceUnit", Order.Ascending}, {"EmployeeKey", Order.Ascending}, {"ShiftStart", Order.Ascending}, {"SourceShiftID", Order.Ascending}})
in
    OrderedOutput
in [Path = OvertimeRosterChecks_Path,
    Anchor = OvertimeRosterChecks_FortnightStartDate,
    Shifts = OvertimeRosterChecks_ShiftResults,
    Fortnights = OvertimeRosterChecks_FortnightResults,
    Segments = OvertimeRosterChecks_ExcessSegments],
Fixtures = (Evaluate as function) as table =>
let
    Anchor = #date(2026, 11, 2),
    WindowStart = #datetime(2026, 11, 2, 0, 0, 0),
    WindowEnd = #datetime(2026, 11, 16, 0, 0, 0),
    RosterColumns = {"Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time",
        "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "Published", "Non Attended", "Status", "MealAvailabilityRequired", "MealStart", "MealEnd",
        "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    WorkerColumns = {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours",
        "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    AsRawSheet = (rows as list, columns as list) as table => Table.DemoteHeaders(Table.FromRecords(rows, columns, MissingField.UseNull)),
    Worker = (employee as text, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = employee, #"Position Description" = "AIN", #"Employee Status Description" = "Permanent Full-Time",
         #"Start Date" = #date(2020, 1, 1), #"Termination Date" = null, #"Contracted FN Hours" = 76,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Shift = (day as number, start as time, finish as time, net as number, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = "E1", #"Employee_Code" = "A1", #"Pay Company" = "Synthetic employer", #"Employee Roster Name" = "Synthetic worker",
         Role = "AIN", Location = "BD", Date = Date.AddDays(Anchor, day), #"Start Time" = start, #"End Time" = finish,
         #"Break Length Minutes" = 0, #"Shift Net Length" = net, #"Shift Type" = null, #"Employment Type" = "Permanent Full-Time",
         Published = true, #"Non Attended" = false, Status = "Published",
         MealAvailabilityRequired = false, MealStart = null, MealEnd = null, CoverageStart = WindowStart, CoverageEnd = WindowEnd,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Carrier = (unit as text) as record => Shift(0, #time(1, 0, 0), #time(2, 0, 0), 1,
        [#"Employee Code" = "C" & unit, #"Employee_Code" = "AC" & unit, #"Employee Roster Name" = "Coverage carrier " & unit, Location = unit]),
    Thresholds = #table(type table [Title = text, Value = any, Unit = text, Notes = nullable text], {
        {"DailyHoursLimit", 10, "hours", null}, {"CSHospitalityHoursLimit", 8, "hours", null}, {"CSPersonalCareHoursLimit", 10, "hours", null},
        {"MinimumBreakHours", 8, "hours", null}, {"ShortGapReviewHours", 10, "hours", null}, {"MinimumFortnightDutyFreeDays", 4, "days", null},
        {"FortnightHoursLimit", 76, "hours", null}, {"FortnightLengthDays", 14, "days", null}}),
    Periods = #table(type table [ShiftPeriod = text, StartDay = number], {{"AM", 6 / 24}, {"PM", 14 / 24}, {"NIGHT", 22 / 24}}),
    Dates = (start as date, count as number) as table => #table(type table [Date = date],
        List.Transform(List.Numbers(0, count), each {Date.AddDays(start, _)})),
    Settings = (dates as table) as table => #table(type table [Item = text, Kind = text, Data = table], {
        {"ShiftPeriod", "Table", Periods}, {"PermutationDimensions", "Table", dates}}),
    Roles = #table(type table [#"Roster Roles" = text, #"Role Group" = text], {{"AIN", "AIN"}, {"EN", "EN"}, {"RN", "RN"}, {"CS", "CS"}}),
    Facilities = #table(type table [Location = text, #"Facility-Abbrev" = text], {{"BD", "BD"}, {"TE", "TE"}, {"JH-RY", "JH-RY"}}),
    Inputs = (rows as list, optional changes as nullable record) as record =>
        let
            Base = [
                PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\OvertimeRosterChecks.xlsx"}}),
                PathMapping = #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {{"https://synthetic.invalid/residentialcare", "C:\SyntheticResidentialCare"}}),
                RosterBD = AsRawSheet(rows & {Carrier("BD")}, RosterColumns),
                RosterTE = AsRawSheet({Carrier("TE")}, RosterColumns),
                RosterJHRY = AsRawSheet({Carrier("JH-RY")}, RosterColumns),
                WorkersBD = AsRawSheet({Worker("E1"), Worker("CBD")}, WorkerColumns),
                WorkersTE = AsRawSheet({Worker("E1"), Worker("CTE")}, WorkerColumns),
                WorkersJHRY = AsRawSheet({Worker("E1"), Worker("CJH-RY")}, WorkerColumns),
                SettingsBD = Settings(Dates(Anchor, 14)), SettingsTE = Settings(Dates(Anchor, 14)), SettingsJHRY = Settings(Dates(Anchor, 14)),
                Thresholds = Thresholds, Roles = Roles, Facilities = Facilities],
            Changed = Record.Combine({Base, if changes = null then [] else changes}),
            // A worker report is a current employee fact. Keep the same synthetic
            // employee fact across unit reports unless a case explicitly varies it.
            ChangedWorkers = if changes <> null and Record.HasFields(changes, "WorkersBD") then Table.PromoteHeaders(changes[WorkersBD]) else null,
            TargetWorkerRows = if ChangedWorkers = null then {} else Table.ToRecords(Table.SelectRows(ChangedWorkers, each [Employee Code] = "E1")),
            ConsistentTE = if ChangedWorkers <> null and not Record.HasFields(changes, "WorkersTE") then
                [WorkersTE = AsRawSheet(TargetWorkerRows & {Worker("CTE")}, WorkerColumns)] else [],
            ConsistentJHRY = if ChangedWorkers <> null and not Record.HasFields(changes, "WorkersJHRY") then
                [WorkersJHRY = AsRawSheet(TargetWorkerRows & {Worker("CJH-RY")}, WorkerColumns)] else [],
            Result = Record.Combine({Changed, ConsistentTE, ConsistentJHRY})
        in Result,
    WithCoverage = (input as record, start as datetime, finish as datetime) as record =>
        let SetCoverage = (raw as table) as table => Table.DemoteHeaders(Table.TransformColumns(Table.PromoteHeaders(raw),
                {{"CoverageStart", each start}, {"CoverageEnd", each finish}})),
            Calendar = Settings(Dates(Date.From(start), Duration.Days(Date.From(finish) - Date.From(start))))
        in Record.Combine({Record.TransformFields(input, {{"RosterBD", SetCoverage}, {"RosterTE", SetCoverage}, {"RosterJHRY", SetCoverage}}),
            [SettingsBD = Calendar, SettingsTE = Calendar, SettingsJHRY = Calendar]}),
    WithRosterDatesOffset = (input as record, days as number) as record =>
        let OffsetDates = (raw as table) as table => Table.DemoteHeaders(Table.TransformColumns(Table.PromoteHeaders(raw),
                {{"Date", each Date.AddDays(Date.From(_), days)}}))
        in Record.TransformFields(input, {{"RosterBD", OffsetDates}, {"RosterTE", OffsetDates}, {"RosterJHRY", OffsetDates}}),
    ChangeThreshold = (title as text, value as any) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(Thresholds), each if [Title] = title then
            Record.Combine({_, [Value = value]}) else _), Value.Type(Thresholds)),
    Near = (actual as any, expected as number) as logical => Value.Is(actual, type number) and Number.Abs(actual - expected) < 0.00000001,
    TargetShifts = (result as record) as table => Table.SelectRows(result[Shifts], each [SourceName] = "Synthetic worker"),
    TargetFortnights = (result as record) as table => Table.SelectRows(result[Fortnights], each [EmployeeID] = "E1"),
    OneShift = (result as record) as record => TargetShifts(result){0},
    LaterShift = (result as record) as record => Table.SelectRows(TargetShifts(result), each [GapHours] <> null){0},
    Unknown = (value as any) as logical => Value.Is(value, type text) and Text.Contains(Text.Upper(value), "UNKNOWN"),
    Check = (name as text, input as record, predicate as function) as record =>
        let
            Attempt = try Evaluate(input),
            Assert = if Attempt[HasError] then null else try predicate(Attempt[Value]),
            Passed = if Attempt[HasError] then false else if Assert[HasError] then false else Assert[Value] = true,
            Details = if Attempt[HasError] then try Attempt[Error][Message] otherwise "Production evaluation error"
                else if Assert[HasError] then try Assert[Error][Message] otherwise "Assertion evaluation error"
                else if Passed then null else "Actual production result differs from the independent fixture expectation"
        in [Case = name, Passed = Passed, Details = Details],
    Reject = (name as text, input as record, force as function, expectedSetting as text) as record =>
        let
            Attempt = try force(Evaluate(input)),
            Message = if Attempt[HasError] then try Attempt[Error][Message] otherwise "Evaluation rejected" else null,
            Passed = Attempt[HasError] and Message <> null and Text.Contains(Message, expectedSetting, Comparer.OrdinalIgnoreCase)
        in [Case = name, Passed = Passed, Details = if Passed then null else if Attempt[HasError] then Message else "Invalid configuration was not rejected"],
    GapInputs = (laterDay as number, laterStart as time, laterFinish as time, laterNet as number, optional changes as nullable record) as record =>
        Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(laterDay, laterStart, laterFinish, laterNet)}, changes),
    DayShifts = (count as number) as list => List.Transform(List.Numbers(0, count), each Shift(_, #time(6, 0, 0), #time(14, 0, 0), 8)),
    CompleteTenDays = Inputs(DayShifts(10)),
    CompleteElevenDays = Inputs(DayShifts(11)),
    SeventySixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(10, 0, 0), 4)},
    SeventyFourPlusSixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(8, 0, 0), 2), Shift(10, #time(6, 0, 0), #time(12, 0, 0), 6)},
    FiftyFiveHours = Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9), Shift(2, #time(22, 0, 0), #time(6, 0, 0), 8)}),
    DailyFortnightOverlap = Inputs(DayShifts(8) & {Shift(8, #time(0, 0, 0), #time(14, 0, 0), 14)}),
    Cases = {
        Check("Daily exactly ten hours has no excess", Inputs({Shift(0, #time(6, 0, 0), #time(16, 0, 0), 10)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Daily eleven net hours has one excess hour", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and OneShift(r)[DailyHoursFlag] = true),
        Check("CS Hospitality exactly eight hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [Role = "CS", AgreementStream = "Hospitality"])},
                [WorkersBD = AsRawSheet({Worker("E1", [AgreementStream = "Hospitality"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("CS Hospitality eight and one-quarter hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 15, 0), 8.25, [Role = "CS", AgreementStream = "Hospitality"])},
                [WorkersBD = AsRawSheet({Worker("E1", [AgreementStream = "Hospitality"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Near(OneShift(r)[DailyExcessHours], 0.25) and OneShift(r)[DailyHoursFlag] = true),
        Check("Conflicting roster and worker agreement streams remain unknown", Inputs({Shift(0, #time(6, 0, 0), #time(14, 15, 0), 8.25,
                [Role = "CS", AgreementStream = "Hospitality"])}),
            (r) => Unknown(OneShift(r)[PolicyStatus]) and OneShift(r)[DailyHoursFlag] = null),
        Check("CS Personal Care uses ten-hour limit", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = "Personal Care"])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Unpaid meal is excluded once", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0)),
        Check("Required client availability during meal counts as work", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealAvailabilityRequired = true, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1)),
        Check("Known overtime clock segments never include an unpaid meal", Inputs({Shift(0, #time(6, 0, 0), #time(21, 0, 0), 14,
                [#"Break Length Minutes" = 60, MealStart = #time(17, 0, 0), MealEnd = #time(18, 0, 0)])}),
            (r) => Table.RowCount(r[Segments]) = 2 and Near(List.Sum(r[Segments][SegmentHours]), 4)
                and Near(List.Sum(List.Transform(Table.ToRecords(r[Segments]), each Duration.TotalHours([SegmentFinish] - [SegmentStart]))), 4)
                and Table.IsEmpty(Table.SelectRows(r[Segments], each [SegmentStart] < #datetime(2026, 11, 2, 18, 0, 0) and [SegmentFinish] > #datetime(2026, 11, 2, 17, 0, 0)))),
        Check("Unknown meal timing preserves known daily excess quantity", Inputs({Shift(0, #time(6, 0, 0), #time(18, 0, 0), 11,
                [#"Break Length Minutes" = 60])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and Table.RowCount(Table.SelectRows(r[Segments], each Unknown([TimingStatus]))) > 0),
        Check("Seven hours fifty-nine minutes fails minimum rest", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 7 + 59 / 60) and LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly eight hours passes minimum rest but needs short-gap review", GapInputs(0, #time(22, 0, 0), #time(23, 0, 0), 1),
            (r) => Near(LaterShift(r)[GapHours], 8) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly ten hours has no short-gap trigger", GapInputs(1, #time(0, 0, 0), #time(2, 0, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 10) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = false),
        Check("First observed shift does not invent a predecessor", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => OneShift(r)[GapHours] = null and Unknown(OneShift(r)[GapStatus])),
        Check("Cross-unit duty participates in rest chronology", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(22, 0, 0), #time(23, 0, 0), 1, [Location = "TE"])}, RosterColumns)]),
            (r) => Table.RowCount(TargetShifts(r)) = 2 and Near(LaterShift(r)[GapHours], 8)),
        Check("Casual short rest remains a roster breach", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Employment Type" = "Casual"]),
                Shift(0, #time(21, 59, 0), #time(23, 59, 0), 2, [#"Employment Type" = "Casual"])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Employee Status Description" = "Casual", #"Contracted FN Hours" = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Documented broken-shift exception affects overtime review only", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8),
                Shift(0, #time(21, 0, 0), #time(23, 0, 0), 2, [BrokenShiftException = true])}),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Four complete duty-free dates passes days-off check", CompleteTenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 4) and Near(TargetFortnights(r){0}[DaysShort], 0)),
        Check("Three complete duty-free dates flags one day short", CompleteElevenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3) and Near(TargetFortnights(r){0}[DaysShort], 1)),
        Check("Missing one unit's declared coverage leaves days off unknown", Inputs(DayShifts(10),
                [RosterTE = AsRawSheet({Record.Combine({Carrier("TE"), [CoverageStart = null, CoverageEnd = null]})}, RosterColumns)]),
            (r) => Unknown(TargetFortnights(r){0}[DaysOffStatus]) and TargetFortnights(r){0}[DutyFreeDays] = null),
        Check("Overnight duty occupies both touched calendar dates", Inputs({Shift(0, #time(22, 0, 0), #time(6, 0, 0), 8)} &
                List.Transform({2..10}, each Shift(_, #time(6, 0, 0), #time(7, 0, 0), 1))),
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3)),
        Check("Monday fifteen to Wednesday twenty-two gives fifty-five elapsed gap hours", FiftyFiveHours,
            (r) => Near(LaterShift(r)[GapHours], 55)),
        Check("Removed pre-night fields and trigger are absent from published outputs", FiftyFiveHours,
            (r) => List.IsEmpty(List.Intersect({Table.ColumnNames(r[Shifts]), {"PreviousRosterShiftPeriod", "NightTransitionFlag", "NightDutyFreeDays",
                    "NightDaysFlag", "NightGapStatus", "NightStatus", "NightTransitionStatus"}}))
                and List.AllTrue(List.Transform(r[Shifts][TriggerSummary], each not Text.Contains(_, "BeforeNight")))
                and List.AllTrue(List.Transform(r[Segments][TriggerSummary], each not Text.Contains(_, "BeforeNight")))),
        Check("Exactly seventy-six net hours has no fortnight excess", Inputs(SeventySixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 76) and Near(TargetFortnights(r){0}[FortnightExcessHours], 0)),
        Check("Seventy-four plus six identifies four excess hours", Inputs(SeventyFourPlusSixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)
                and Near(List.Sum(TargetShifts(r)[FortnightExcessHours]), 4)),
        Check("Daily overtime hours still enter the net fortnight counter once", Inputs(List.Transform({0..5}, each Shift(_, #time(0, 0, 0), #time(10, 0, 0), 10)) &
                {Shift(6, #time(0, 0, 0), #time(14, 0, 0), 14), Shift(7, #time(0, 0, 0), #time(6, 0, 0), 6)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)),
        Check("Daily and fortnight overlapping excess is counted once", DailyFortnightOverlap,
            (r) => Near(TargetFortnights(r){0}[FortnightExcessHours], 2) and Near(List.Sum(r[Segments][SegmentHours]), 4)
                and Table.RowCount(Table.SelectRows(r[Segments], each [DailyTrigger] and [FortnightTrigger])) > 0),
        Check("A known boundary-crossing shift allocates two and six hours to adjacent fortnights", WithCoverage(Inputs({Shift(13, #time(22, 0, 0), #time(6, 0, 0), 8)}),
                WindowStart, #datetime(2026, 11, 30, 0, 0, 0)),
            (r) => List.Sort(TargetFortnights(r)[NetRosterHours]) = {2, 6}),
        Check("Unknown meal placement across a fortnight boundary is not prorated", WithCoverage(Inputs({Shift(13, #time(20, 0, 0), #time(8, 0, 0), 11,
                [#"Break Length Minutes" = 60])}), WindowStart, #datetime(2026, 11, 30, 0, 0, 0)),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and Near(List.Sum(r[Segments][SegmentHours]), 1)),
        Check("Invalid equal start and finish record cannot leave a mixed fortnight clear", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(1, #time(6, 0, 0), #time(6, 0, 0), 8)}),
            (r) => List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and List.AllTrue(List.Transform(TargetFortnights(r)[DaysOffStatus], Unknown))),
        Check("Missing start and finish record cannot disappear from mixed fortnight coverage", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8),
                Shift(1, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Start Time" = null, #"End Time" = null])}),
            (r) => List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))
                and List.AllTrue(List.Transform(TargetFortnights(r)[DaysOffStatus], Unknown))),
        Check("Exact duplicate source rows never double known net hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 8) or Unknown(TargetFortnights(r){0}[FortnightStatus])),
        Check("Overlapping cross-unit duties are explicitly unknown", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(10, 0, 0), #time(18, 0, 0), 8, [Location = "TE"])}, RosterColumns)]),
            (r) => List.AllTrue(List.Transform(TargetShifts(r)[OverlapStatus], each Unknown(_) or _ = "Overlap"))),
        Check("Conflicting worker rows cannot silently choose employment type", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [WorkersBD = AsRawSheet({Worker("E1"), Worker("E1", [#"Employee Status Description" = "Casual"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[IdentityStatus]) or Unknown(OneShift(r)[DailyStatus])),
        Check("Missing CS agreement stream remains unknown", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Position Description" = "CS", AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) and OneShift(r)[DailyHoursFlag] = null),
        Reject("Invalid nonblank daily threshold is explicitly rejected", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2,
                [Thresholds = ChangeThreshold("DailyHoursLimit", "ten")]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Reject("Duplicate threshold title is explicitly rejected", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = Table.Combine({Thresholds, #table(type table [Title = text, Value = any, Unit = text, Notes = nullable text], {{"DailyHoursLimit", 12, "hours", null}})})]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Reject("Missing required threshold title is explicitly rejected", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = Table.SelectRows(Thresholds, each [Title] <> "DailyHoursLimit")]),
            (r) => OneShift(r)[DailyExcessHours], "DailyHoursLimit"),
        Check("Empty Settings dates leave fortnight unknown without a roster-date fallback", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [SettingsBD = Settings(Dates(Anchor, 0)), SettingsTE = Settings(Dates(Anchor, 0)), SettingsJHRY = Settings(Dates(Anchor, 0))]),
            (r) => r[Anchor] = null and Near(OneShift(r)[DailyExcessHours], 1) and Unknown(OneShift(r)[FortnightStatus])),
        Check("Settings minimum anchors the fortnight when every published shift starts later", WithRosterDatesOffset(
                Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}), 3),
            (r) => r[Anchor] = Anchor and Date.From(OneShift(r)[ShiftStart]) = Date.AddDays(Anchor, 3)
                and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Check("Unordered duplicate Settings dates still derive the earliest date", Inputs({Shift(3, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [SettingsBD = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)})),
                 SettingsTE = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)})),
                 SettingsJHRY = Settings(Table.Combine({Table.ReverseRows(Dates(Anchor, 14)), Table.FirstN(Dates(Anchor, 14), 3)}))]),
            (r) => r[Anchor] = Anchor and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Check("Anchor is the minimum across unit Settings even when one unit has no dates", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [SettingsBD = Settings(Dates(Date.AddDays(Anchor, 5), 14)), SettingsTE = Settings(Dates(Anchor, 0)), SettingsJHRY = Settings(Dates(Anchor, 14))]),
            (r) => r[Anchor] = Anchor and TargetFortnights(r){0}[PayrollFortnightStart] = Anchor),
        Reject("Invalid Settings calendar date is explicitly rejected", Inputs({},
                [SettingsBD = Settings(#table({"Date"}, {{"not-a-date"}}))]),
            (r) => r[Anchor], "Date"),
        Reject("Missing Settings calendar Date column is explicitly rejected", Inputs({},
                [SettingsBD = Settings(#table({"Day"}, {{1}}))]),
            (r) => r[Anchor], "Date"),
        Check("Missing payroll eligibility remains explicit", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11,
                [OvertimeExempt = null, Grade = null, AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [OvertimeExempt = null, Grade = null, AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) or Unknown(OneShift(r)[OvertimeStatus])),
        Check("Unresolved DST offset does not publish a known elapsed rest result", WithCoverage(Inputs({
                Shift(-30, #time(15, 0, 0), #time(23, 0, 0), 8), Shift(-29, #time(9, 0, 0), #time(11, 0, 0), 2)}),
                #datetime(2026, 10, 1, 0, 0, 0), WindowEnd),
            (r) => Unknown(Table.SelectRows(TargetShifts(r), each [GapWallHours] <> null){0}[GapStatus])
                and Table.SelectRows(TargetShifts(r), each [GapWallHours] <> null){0}[MinimumBreakFlag] = null),
        Check("Source shift keys are stable under reordered roster rows", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(2, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => List.Sort(TargetShifts(r)[SourceShiftID]) = List.Sort(TargetShifts(Evaluate(Inputs({
                Shift(2, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)})))[SourceShiftID])),
        Check("Duplicated roster publications do not create extra actual shift identities", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}, RosterColumns)]),
            (r) => List.Count(List.Distinct(TargetShifts(r)[SourceShiftID])) = 1 and
                (Near(TargetFortnights(r){0}[NetRosterHours], 8) or Unknown(TargetFortnights(r){0}[FortnightStatus]))),
        Check("Worker and allocation codes retain separate identifier meanings", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => OneShift(r)[EmployeeID] = "E1" and OneShift(r)[AllocationCode] = "A1" and not Unknown(OneShift(r)[IdentityStatus])),
        Check("Allocation code alone is not guessed to be a worker identifier", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Employee Code" = null])}),
            (r) => Unknown(Table.SelectRows(r[Shifts], each [AllocationCode] = "A1"){0}[IdentityStatus])),
        Check("Overtime exemption does not suppress the minimum-rest roster check", Inputs({
                Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [OvertimeExempt = true]),
                Shift(0, #time(21, 0, 0), #time(23, 0, 0), 2, [OvertimeExempt = true])}),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Off-role duty still occupies a target employee's calendar day", Inputs(DayShifts(10) & {
                Shift(10, #time(6, 0, 0), #time(7, 0, 0), 1, [Role = "Other role"])},
                [Roles = Table.Combine({Roles, #table(type table [#"Roster Roles" = text, #"Role Group" = text], {{"Other role", "OTHER"}})})]),
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3)),
        Check("Resolver accepts CELL filename brackets at the approved nested target", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\[OvertimeRosterChecks.xlsx]Sheet1"}})]),
            (r) => Comparer.OrdinalIgnoreCase(r[Path][ClientPath], "C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon") = 0),
        Check("Resolver uses case-insensitive longest SharePoint prefix", Inputs({}, [
                PathInput = #table({"FilePath"}, {{"https://SYNTHETIC.invalid/residentialcare/client/datexx-whiddon/2. calculations/overtimerosterchecks/overtimerosterchecks.xlsx"}}),
                PathMapping = #table({"SharepointRootUrl", "SyncedFolderRootPath"}, {
                    {"https://synthetic.invalid/residentialcare", "C:\SyntheticResidentialCare"},
                    {"https://synthetic.invalid/residentialcare/CLIENT/DATExx-Whiddon", "C:\LongestAuthoritativeWhiddon"}})]),
            (r) => Comparer.OrdinalIgnoreCase(r[Path][ClientPath], "C:\LongestAuthoritativeWhiddon") = 0),
        Reject("Resolver refuses a false local prefix boundary", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCareCollision\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\OvertimeRosterChecks.xlsx"}})]),
            (r) => r[Path][ClientPath], "CentriSyncPaths"),
        Reject("Resolver refuses an alternative workbook filename", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks\RosterChecksCopy.xlsx"}})]),
            (r) => r[Path][ClientPath], "OvertimeRosterChecks.xlsx"),
        Reject("Resolver refuses the organisation calculation parent folder", Inputs({},
                [PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\OvertimeRosterChecks.xlsx"}})]),
            (r) => r[Path][ClientPath], "OvertimeRosterChecks")
    },
    Results = Table.FromRecords(Cases, type table [Case = text, Passed = logical, Details = nullable text])
in
    Results
in Fixtures(Evaluate);
