// Power Query from: Demand-MasterRoster Manual Read.xlsx
// Pathname: c:\Users\alexp\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\Demand-MasterRoster Manual Read.xlsx
// Extracted: 2026-08-29T20:58:10.990Z

section Section1;

shared #"IMPORT Master" = let
    Source = Excel.Workbook(File.Contents("C:\Users\alexp\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\ResidentialCare\Whiddon\DATExx\UNITS\Unit1\1. Input\Master Roster.xlsx"), null, true),
    Combined_Sheet = Source{[Item="Combined",Kind="Sheet"]}[Data],
    #"Promoted Headers1" = Table.PromoteHeaders(Combined_Sheet, [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers1",{{"Master Template", type text}, {"Template", type text}, {"Location", type text}, {"Department", type text}, {"Role", type text}, {"Area", type text}, {"Employee Code", Int64.Type}, {"Employee Name", type text}, {"Week No", Int64.Type}, {"Week Day", type text}, {"Start Time", type time}, {"End Time", type time}, {"Roster Hours", type number}, {"Cost", type number}, {"MinRosterHours", type number}, {"MaxRosterHours", Int64.Type}, {"Event", type text}, {"Break Length", Int64.Type}, {"Break Start Time", type time}, {"Paid Break Length", Int64.Type}, {"Paid Break Start Time", type time}, {"Shift Definition", type text}, {"Shift Net Length", type number}, {"Shift Type", type text}, {"Non Attended", type logical}})
in
    #"Changed Type1";

shared #"INPUT SHiftEnd" = let
    Source = Excel.CurrentWorkbook(){[Name="Table7"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shift", type text}, {"End", type time}})
in
    #"Changed Type";

shared #"ShftEnd Prepare" = let
    Source = #"INPUT SHiftEnd",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type)
in
    #"Added Index";

// Query: INPUT MinuteWorkers
// Purpose: Reads the configured MinuteWorker roles and their Direct Care percentages.
// Inputs: Current-workbook table MinuteWorkersTable.
// Output: Role, QFR Category, and Direct Care %.
shared #"INPUT MinuteWorkers" =
let
    Source = Excel.CurrentWorkbook(){[Name = "MinuteWorkersTable"]}[Content],
    #"Assigned Required Types" = Table.TransformColumnTypes(
        Source,
        {{"Role", type text}, {"QFR Category", type text}, {"Direct Care %", type number}}
    )
in
    #"Assigned Required Types";

// Query: MW Abbreviate Role
// Purpose: Converts both original master-roster roles and configured MinuteWorker names to one role key.
// Inputs: A nullable role name from IMPORT Master or INPUT MinuteWorkers.
// Output: RN, AIN, EN, CCCM, RSM, CSM, WCSO, CAM, RGM, TA, or the trimmed source value.
// Notes: Medication-competent AIN shifts belong to the configured AIN role; CAM and RGM remain valid keys even when history is absent.
shared #"MW Abbreviate Role" = (RoleName as nullable text) as nullable text =>
let
    TrimmedRole = if RoleName = null then null else Text.Trim(RoleName),
    UpperRole = if TrimmedRole = null then null else Text.Upper(TrimmedRole),
    AbbreviatedRole =
        if UpperRole = null or UpperRole = "" then
            null
        else if
            UpperRole = "RN" or
            Text.Contains(UpperRole, "REGN") or
            Text.Contains(UpperRole, "REGISTERED NURSE")
        then
            "RN"
        else if
            UpperRole = "AIN" or UpperRole = "AINC4" or
            Text.Contains(UpperRole, "ASST IN NURSING") or
            Text.Contains(UpperRole, "ASSISTANT IN NURSING")
        then
            "AIN"
        else if UpperRole = "EN" or Text.Contains(UpperRole, "ENROLLED NURSE") then
            "EN"
        else if UpperRole = "CCCM" or Text.Contains(UpperRole, "CLINICAL CARE COORDINATOR") then
            "CCCM"
        else if UpperRole = "RSM" or Text.Contains(UpperRole, "RESIDENTIAL SERVICES MANAGER") then
            "RSM"
        else if
            UpperRole = "CSM" or
            Text.Contains(UpperRole, "CARE SERVICE MANAGER") or
            Text.Contains(UpperRole, "CARE SERVICES MANAGER")
        then
            "CSM"
        else if
            UpperRole = "WCSO" or UpperRole = "WLO" or
            Text.Contains(UpperRole, "WELLBEING & CARE SUPPORT OFFICER") or
            Text.Contains(UpperRole, "WELLBEING AND CARE SUPPORT OFFICER") or
            Text.Contains(UpperRole, "WELLBEING & LIFESTYLE OFFICER") or
            Text.Contains(UpperRole, "WELLBEING AND LIFESTYLE OFFICER")
        then
            "WCSO"
        else if
            UpperRole = "CAM" or
            Text.Contains(UpperRole, "CARE & ASSESSMENT MANAGER") or
            Text.Contains(UpperRole, "CARE AND ASSESSMENT MANAGER")
        then
            "CAM"
        else if UpperRole = "RGM" or Text.Contains(UpperRole, "REGIONAL GENERAL MANAGER") then
            "RGM"
        else if UpperRole = "TA" or Text.Contains(UpperRole, "THERAPY ASSISTANT") then
            "TA"
        else
            TrimmedRole
in
    AbbreviatedRole;

// Query: MW MinuteWorkers Prepare
// Purpose: Normalises configured MinuteWorkers to the same abbreviations used by historical calculations.
// Inputs: INPUT MinuteWorkers and MW Abbreviate Role.
// Output: One row per configured role with SourceRole, abbreviated Role, RoleKey, and MinuteCategory.
shared #"MW MinuteWorkers Prepare" =
let
    Source = #"INPUT MinuteWorkers",
    #"Trimmed Role Fields" = Table.TransformColumns(
        Source,
        {
            {"Role", each if _ = null then null else Text.Trim(_), type nullable text},
            {"QFR Category", each if _ = null then null else Text.Trim(_), type nullable text}
        }
    ),
    #"Preserved Source Role" = Table.DuplicateColumn(
        #"Trimmed Role Fields",
        "Role",
        "SourceRole"
    ),
    #"Abbreviated Role" = Table.TransformColumns(
        #"Preserved Source Role",
        {{"Role", each #"MW Abbreviate Role"(_), type nullable text}}
    ),
    #"Added Role Key" = Table.AddColumn(
        #"Abbreviated Role",
        "RoleKey",
        each if [Role] = null then null else Text.Upper([Role]),
        type nullable text
    ),
    // Only the canonical RN role receives the dedicated RN target. All other
    // configured MinuteWorkers, including RN-qualified managers, use OTHERS.
    #"Added Minute Category" = Table.AddColumn(
        #"Added Role Key",
        "MinuteCategory",
        each if [RoleKey] = "RN" then "RN" else "OTHERS",
        type text
    )
in
    #"Added Minute Category";

// Query: Master Prepare
// Purpose: Normalises original master-roster roles, retains configured MinuteWorkers, and assigns shifts.
// Inputs: IMPORT Master, MW Abbreviate Role, MW MinuteWorkers Prepare, ShftEnd Prepare, and shift-boundary scalars.
// Output: Historical roster rows using the same abbreviated Role keys as INPUT MinuteWorkers.
shared #"Master Prepare" =
let
    Source = #"IMPORT Master",
    EligibleMinuteWorkerRoles = List.Buffer(
        List.Distinct(
            List.RemoveNulls(#"MW MinuteWorkers Prepare"[RoleKey])
        )
    ),
    #"Normalised Location" = Table.TransformColumns(
        Source,
        {
            {
                "Location",
                each if _ = null then null else Text.Upper(Text.Start(Text.Trim(_), 2)),
                type nullable text
            }
        }
    ),
    #"Selected Historical Columns" = Table.SelectColumns(
        #"Normalised Location",
        {"Location", "Role", "Employee Code", "Week No", "Week Day", "Start Time", "End Time", "Roster Hours"}
    ),
    #"Renamed Original Role" = Table.RenameColumns(
        #"Selected Historical Columns",
        {{"Role", "OriginalRole"}}
    ),
    // The same abbreviation function is applied on both sides of the later
    // history join, preventing renamed roles from disappearing from Shift%.
    #"Added Abbreviated Role" = Table.AddColumn(
        #"Renamed Original Role",
        "Role",
        each #"MW Abbreviate Role"([OriginalRole]),
        type nullable text
    ),
    // The configured MinuteWorker list replaces the old incomplete hard-coded
    // role filter. Downstream historical totals and Shift% therefore reset here.
    #"Filtered To MinuteWorker Roles" = Table.SelectRows(
        #"Added Abbreviated Role",
        each List.Contains(EligibleMinuteWorkerRoles, [Role])
    ),
    #"Removed Original Role" = Table.RemoveColumns(
        #"Filtered To MinuteWorker Roles",
        {"OriginalRole"}
    ),
    #"Filtered Facilities" = Table.SelectRows(
        #"Removed Original Role",
        each [Location] <> "NR"
    ),
    #"Assigned Shift" = Table.AddColumn(
        #"Filtered Facilities",
        "Shift",
        each
            if [Start Time] <= StartAM then "NS"
            else if [Start Time] <= StartPM then "AM"
            else if [Start Time] <= StartNS then "PM"
            else if [Start Time] <= #time(23, 59, 59) then "NS"
            else "ERROR",
        type text
    ),
    #"Merged Shift Index" = Table.NestedJoin(
        #"Assigned Shift",
        {"Shift"},
        #"ShftEnd Prepare",
        {"Shift"},
        "ShftEnd Prepare",
        JoinKind.LeftOuter
    ),
    #"Expanded ShftEnd Prepare" = Table.ExpandTableColumn(
        #"Merged Shift Index",
        "ShftEnd Prepare",
        {"Index"},
        {"ShiftIndex"}
    )
in
    #"Expanded ShftEnd Prepare";

shared LocRoleWeekDaysHours = let
    Source = #"Master Prepare",
    #"added DAYOFWEEKS#" = Table.AddColumn(Source, "DayOfWeek", each List.PositionOf(
    {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},
    [Week Day]
) + 1),
    #"Sorted Rows" = Table.Sort(#"added DAYOFWEEKS#",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Grouped Rows" = Table.Group(#"Sorted Rows", {"Location", "Week No", "Role", "DayOfWeek", "Week Day", "Shift", "ShiftIndex"}, {{"Hours", each List.Sum([Roster Hours]), type nullable number}}),
    #"Sorted Rows1" = Table.Sort(#"Grouped Rows",{{"DayOfWeek", Order.Ascending}, {"Location", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Inserted Merged Column1" = Table.AddColumn(#"Sorted Rows1", "WeekDayShift", each Text.Combine({Text.From([DayOfWeek], "en-AU"), [Week Day], [Shift]}, ""), type text),
    #"Inserted Merged Column" = Table.AddColumn(#"Inserted Merged Column1", "WeekAndDay", each Text.Combine({Text.From([Week No], "en-AU"), [Week Day]}, ""), type text),
    #"Inserted Merged Column2" = Table.AddColumn(#"Inserted Merged Column", "DayShift", each Text.Combine({[Week Day], [Shift]}, ""), type text)
in
    #"Inserted Merged Column2";

shared StartPM = let
    Source = #"ShftEnd Prepare",
    End = Source{0}[End]
in
    End;

shared StartNS = let
    Source = #"ShftEnd Prepare",
    End = Source{1}[End]
in
    End;

shared StartAM = let
    Source = #"ShftEnd Prepare",
    End = Source{2}[End]
in
    End;

// Query: INPUT TargetMinutes
// Purpose: Reads RN and ALL daily target minutes for every facility column.
// Inputs: Current-workbook table TargetMinutes.
// Output: MinuteType plus dynamically typed numeric facility columns.
// Notes: Facility columns are dynamic; MW TargetMinutes Prepare converts daily values to 14-day targets.
shared #"INPUT TargetMinutes" =
let
    Source = Excel.CurrentWorkbook(){[Name = "TargetMinutes"]}[Content],
    ColumnNames = Table.ColumnNames(Source),
    RequiredSchema =
        if not List.Contains(ColumnNames, "MinuteType") then
            error "TargetMinutes must contain a MinuteType column."
        else if List.Count(ColumnNames) = 1 then
            error "TargetMinutes must contain at least one facility column."
        else
            Source,
    FacilityColumns = List.RemoveItems(ColumnNames, {"MinuteType"}),
    #"Assigned Required Types" = Table.TransformColumnTypes(
        RequiredSchema,
        {{"MinuteType", type text}} &
            List.Transform(FacilityColumns, each {_, type number})
    )
in
    #"Assigned Required Types";

shared LocRoleWeekDaysHoursMATRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "WeekDayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[WeekDayShift]), "WeekDayShift", "Hours", List.Sum),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVEMATTRIX = let
    Source = LocRoleWeekDaysHours,
    #"Sorted Rows" = Table.Sort(Source,{{"Week No", Order.Ascending}, {"DayOfWeek", Order.Ascending}, {"ShiftIndex", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"Location", "Role", "Hours", "DayShift"}),
    #"Pivoted Column1" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[DayShift]), "DayShift", "Hours", List.Average),
    #"Sorted Rows2" = Table.Sort(#"Pivoted Column1",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows2";

shared LocRoleWeekDaysShiftAVETABLE = let
    Source = LocRoleWeekDaysShiftAVEMATTRIX,
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Location", "Role"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

shared LocRoleTOTAL = let
    Source = LocRoleWeekDaysShiftAVETABLE,
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Hours", each List.Sum([Value]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Location", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared #"LocRoleDayShift%" = let
    Source = Table.NestedJoin(LocRoleWeekDaysShiftAVETABLE, {"Location", "Role"}, LocRoleTOTAL, {"Location", "Role"}, "LocRoleTOTAL", JoinKind.LeftOuter),
    #"Expanded LocRoleTOTAL" = Table.ExpandTableColumn(Source, "LocRoleTOTAL", {"Hours"}, {"Hours"}),
    #"Inserted Division" = Table.AddColumn(#"Expanded LocRoleTOTAL", "LocRoleDayShift%", each [Value] / [Hours], type number),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Division",{{"LocRoleDayShift%", Percentage.Type}})
in
    #"Changed Type";

shared #"LocRoleDayShift%CHECK" = let
    Source = #"LocRoleDayShift%",
    #"Grouped Rows" = Table.Group(Source, {"Location", "Role"}, {{"Check", each List.Sum([#"LocRoleDayShift%"]), type number}})
in
    #"Grouped Rows";

shared #"LocRoleDatShift%MATRIX" = let
    Source = #"LocRoleDayShift%",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Value", "Hours"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Attribute]), "Attribute", "LocRoleDayShift%", List.Sum)
in
    #"Pivoted Column";

shared #"LocRoleWeekDaysHours (2)" = let
    Source = LocRoleWeekDaysHours
in
    Source;

// Query: MW Check Table
// Purpose: Converts validation records to the standard MinuteWorker check-table schema.
// Inputs: A list of validation records.
// Output: Check, Severity, Facility, MinuteCategory, Role, Actual, Expected, and Message.
shared #"MW Check Table" = (Rows as list) as table =>
let
    EmptyCheckTable = #table(
        type table [
            Check = nullable text,
            Severity = nullable text,
            Facility = nullable text,
            MinuteCategory = nullable text,
            Role = nullable text,
            Actual = nullable number,
            Expected = nullable number,
            Message = nullable text
        ],
        {}
    ),
    Result =
        if List.Count(Rows) = 0 then
            EmptyCheckTable
        else
            // Check records share a fixed interface. Missing optional context
            // fields are intentionally represented as null in the check output.
            Table.FromRecords(
                Rows,
                type table [
                    Check = nullable text,
                    Severity = nullable text,
                    Facility = nullable text,
                    MinuteCategory = nullable text,
                    Role = nullable text,
                    Actual = nullable number,
                    Expected = nullable number,
                    Message = nullable text
                ],
                MissingField.UseNull
            )
in
    Result;

// Query: MW TargetMinutes Prepare
// Purpose: Converts daily RN and ALL inputs to daily and 14-day RN/OTHERS category targets.
// Inputs: INPUT TargetMinutes.
// Output: Two rows per facility with daily and fortnight target minutes for RN and OTHERS.
shared #"MW TargetMinutes Prepare" =
let
    Source = #"INPUT TargetMinutes",
    TargetPeriodDays = 14,
    FacilityColumns = List.RemoveItems(Table.ColumnNames(Source), {"MinuteType"}),
    #"Unpivoted Facility Targets" = Table.Unpivot(
        Source,
        FacilityColumns,
        "Facility",
        "TargetMinutes"
    ),
    #"Normalised Target Keys" = Table.TransformColumns(
        #"Unpivoted Facility Targets",
        {
            {"MinuteType", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text},
            {"Facility", each Text.Upper(Text.Trim(_)), type text}
        }
    ),
    #"Summarised Facility Targets" = Table.Group(
        #"Normalised Target Keys",
        {"Facility"},
        {
            {"RNCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "RN")), Int64.Type},
            {"ALLCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "ALL")), Int64.Type},
            {"RNNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "RN" and [TargetMinutes] = null)), Int64.Type},
            {"ALLNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "ALL" and [TargetMinutes] = null)), Int64.Type},
            {
                "RNDailyTargetMinutes",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "RN")[TargetMinutes]
                    in
                        if List.IsEmpty(Values) then null else List.Sum(Values),
                type nullable number
            },
            {
                "ALLDailyTargetMinutes",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "ALL")[TargetMinutes]
                    in
                        if List.IsEmpty(Values) then null else List.Sum(Values),
                type nullable number
            },
            {
                "UnexpectedTypeCount",
                each Table.RowCount(
                    Table.SelectRows(
                        _,
                        each [MinuteType] = null or ([MinuteType] <> "RN" and [MinuteType] <> "ALL")
                    )
                ),
                Int64.Type
            }
        }
    ),
    // TargetMinutes contains one day's productive minutes. Normalise it to the
    // 14-day roster period before role and day/shift allocation.
    #"Added RN Fortnight Target" = Table.AddColumn(
        #"Summarised Facility Targets",
        "RNFortnightTargetMinutes",
        each
            if [RNDailyTargetMinutes] = null then
                null
            else
                [RNDailyTargetMinutes] * TargetPeriodDays,
        type nullable number
    ),
    #"Added ALL Fortnight Target" = Table.AddColumn(
        #"Added RN Fortnight Target",
        "ALLFortnightTargetMinutes",
        each
            if [ALLDailyTargetMinutes] = null then
                null
            else
                [ALLDailyTargetMinutes] * TargetPeriodDays,
        type nullable number
    ),
    // ALL includes RN, so the productive-minute pool available to all other
    // MinuteWorker roles is ALL less RN at both the daily and fortnight grain.
    #"Added Category Target Records" = Table.AddColumn(
        #"Added ALL Fortnight Target",
        "CategoryTargets",
        each {
            [
                MinuteCategory = "RN",
                CategoryDailyTargetMinutes = [RNDailyTargetMinutes],
                CategoryTargetMinutes = [RNFortnightTargetMinutes]
            ],
            [
                MinuteCategory = "OTHERS",
                CategoryDailyTargetMinutes =
                    if [RNDailyTargetMinutes] = null or [ALLDailyTargetMinutes] = null then
                        null
                    else
                        [ALLDailyTargetMinutes] - [RNDailyTargetMinutes],
                CategoryTargetMinutes =
                    if [RNFortnightTargetMinutes] = null or [ALLFortnightTargetMinutes] = null then
                        null
                    else
                        [ALLFortnightTargetMinutes] - [RNFortnightTargetMinutes]
            ]
        },
        type list
    ),
    #"Expanded Category Target Rows" = Table.ExpandListColumn(
        #"Added Category Target Records",
        "CategoryTargets"
    ),
    #"Expanded Category Target Values" = Table.ExpandRecordColumn(
        #"Expanded Category Target Rows",
        "CategoryTargets",
        {"MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"},
        {"MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"}
    )
in
    #"Expanded Category Target Values";

// Query: MW Historical DayShift
// Purpose: Filters the historical day/shift distribution to configured MinuteWorker roles.
// Inputs: LocRoleDayShift%, LocRoleWeekDaysHours, and MW MinuteWorkers Prepare.
// Output: One historical row per facility, role, weekday, and shift.
shared #"MW Historical DayShift" =
let
    Source = #"LocRoleDayShift%",
    #"Renamed Historical Columns" = Table.RenameColumns(
        Source,
        {
            {"Location", "Facility"},
            {"Attribute", "DayShift"},
            {"Value", "HistoricalRosterHours"},
            {"Hours", "RoleHistoricalRosterHours"}
        }
    ),
    #"Normalised Facility" = Table.TransformColumns(
        #"Renamed Historical Columns",
        {{"Facility", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text}}
    ),
    DayShiftMap = Table.Distinct(
        Table.SelectColumns(
            LocRoleWeekDaysHours,
            {"DayShift", "DayOfWeek", "Week Day", "Shift", "ShiftIndex"}
        )
    ),
    #"Merged DayShift Details" = Table.NestedJoin(
        #"Normalised Facility",
        {"DayShift"},
        DayShiftMap,
        {"DayShift"},
        "DayShiftDetails",
        JoinKind.LeftOuter
    ),
    #"Expanded DayShift Details" = Table.ExpandTableColumn(
        #"Merged DayShift Details",
        "DayShiftDetails",
        {"DayOfWeek", "Week Day", "Shift", "ShiftIndex"},
        {"DayOfWeek", "Week Day", "Shift", "ShiftIndex"}
    ),
    #"Added Role Key" = Table.AddColumn(
        #"Expanded DayShift Details",
        "RoleKey",
        each if [Role] = null then null else Text.Upper(Text.Trim([Role])),
        type nullable text
    ),
    // The inner join is the MinuteWorker filter: historical roles not listed in
    // INPUT MinuteWorkers do not enter the new analysis.
    #"Filtered To MinuteWorkers" = Table.NestedJoin(
        #"Added Role Key",
        {"RoleKey"},
        #"MW MinuteWorkers Prepare",
        {"RoleKey"},
        "MinuteWorker",
        JoinKind.Inner
    ),
    #"Expanded MinuteWorker Attributes" = Table.ExpandTableColumn(
        #"Filtered To MinuteWorkers",
        "MinuteWorker",
        {"QFR Category", "Direct Care %", "MinuteCategory"},
        {"QFR Category", "Direct Care %", "MinuteCategory"}
    ),
    #"Selected Historical Columns" = Table.SelectColumns(
        #"Expanded MinuteWorker Attributes",
        {
            "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "Shift", "ShiftIndex", "DayShift",
            "HistoricalRosterHours", "RoleHistoricalRosterHours", "LocRoleDayShift%"
        }
    )
in
    #"Selected Historical Columns";

// Query: MW Role Distribution
// Purpose: Calculates each role's share of category productive minutes from historical roster hours.
// Inputs: MW Historical DayShift.
// Output: One row per facility and MinuteWorker role with RoleHistoryDistribution%.
shared #"MW Role Distribution" =
let
    Source = #"MW Historical DayShift",
    #"Selected Distinct Role History" = Table.Distinct(
        Table.SelectColumns(
            Source,
            {
                "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category",
                "Direct Care %", "RoleHistoricalRosterHours"
            }
        )
    ),
    // Only the Direct Care portion of historical roster hours contributes to
    // the role distribution used to allocate target minutes.
    #"Added Historical Minute Hours" = Table.AddColumn(
        #"Selected Distinct Role History",
        "HistoricalMinuteHours",
        each [RoleHistoricalRosterHours] * [#"Direct Care %"],
        type number
    ),
    #"Summarised Category History" = Table.Group(
        #"Added Historical Minute Hours",
        {"Facility", "MinuteCategory"},
        {
            {
                "CategoryHistoricalMinuteHours",
                each List.Sum([HistoricalMinuteHours]),
                type number
            }
        }
    ),
    #"Merged Category History" = Table.NestedJoin(
        #"Added Historical Minute Hours",
        {"Facility", "MinuteCategory"},
        #"Summarised Category History",
        {"Facility", "MinuteCategory"},
        "CategoryHistory",
        JoinKind.LeftOuter
    ),
    #"Expanded Category History" = Table.ExpandTableColumn(
        #"Merged Category History",
        "CategoryHistory",
        {"CategoryHistoricalMinuteHours"},
        {"CategoryHistoricalMinuteHours"}
    ),
    // RoleHistoryDistribution% is calculated within facility and MinuteCategory.
    #"Added Role History Distribution" = Table.AddColumn(
        #"Expanded Category History",
        "RoleHistoryDistribution%",
        each
            if [CategoryHistoricalMinuteHours] = null or [CategoryHistoricalMinuteHours] = 0 then
                null
            else
                [HistoricalMinuteHours] / [CategoryHistoricalMinuteHours],
        Percentage.Type
    )
in
    #"Added Role History Distribution";

// Query: MW Role Targets
// Purpose: Allocates each facility/category fortnight target across eligible MinuteWorker roles.
// Inputs: MW Role Distribution and MW TargetMinutes Prepare.
// Output: One row per facility and role with daily and 14-day RoleTargetMinutes.
shared #"MW Role Targets" =
let
    Source = #"MW Role Distribution",
    #"Merged Category Targets" = Table.NestedJoin(
        Source,
        {"Facility", "MinuteCategory"},
        #"MW TargetMinutes Prepare",
        {"Facility", "MinuteCategory"},
        "CategoryTarget",
        JoinKind.Inner
    ),
    #"Expanded Category Targets" = Table.ExpandTableColumn(
        #"Merged Category Targets",
        "CategoryTarget",
        {
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RNDailyTargetMinutes", "ALLDailyTargetMinutes",
            "RNFortnightTargetMinutes", "ALLFortnightTargetMinutes"
        },
        {
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RNDailyTargetMinutes", "ALLDailyTargetMinutes",
            "RNFortnightTargetMinutes", "ALLFortnightTargetMinutes"
        }
    ),
    #"Added Role Daily Target Minutes" = Table.AddColumn(
        #"Expanded Category Targets",
        "RoleDailyTargetMinutes",
        each [CategoryDailyTargetMinutes] * [#"RoleHistoryDistribution%"],
        type number
    ),
    // Each role receives its Direct Care-adjusted share of the 14-day category target.
    #"Added Role Target Minutes" = Table.AddColumn(
        #"Added Role Daily Target Minutes",
        "RoleTargetMinutes",
        each [CategoryTargetMinutes] * [#"RoleHistoryDistribution%"],
        type number
    )
in
    #"Added Role Target Minutes";

// Query: MW DayShift Allocation
// Purpose: Distributes role targets across the representative week and converts roster minutes to FTE.
// Inputs: MW Historical DayShift and MW Role Targets.
// Output: One row per facility, role, weekday, and shift with unrounded FTE.
shared #"MW DayShift Allocation" =
let
    Source = #"MW Historical DayShift",
    #"Merged Role Targets" = Table.NestedJoin(
        Source,
        {"Facility", "MinuteCategory", "RoleKey"},
        #"MW Role Targets",
        {"Facility", "MinuteCategory", "RoleKey"},
        "RoleTarget",
        JoinKind.Inner
    ),
    #"Expanded Role Targets" = Table.ExpandTableColumn(
        #"Merged Role Targets",
        "RoleTarget",
        {
            "HistoricalMinuteHours", "CategoryHistoricalMinuteHours",
            "RoleHistoryDistribution%", "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleTargetMinutes"
        },
        {
            "HistoricalMinuteHours", "CategoryHistoricalMinuteHours",
            "RoleHistoryDistribution%", "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleTargetMinutes"
        }
    ),
    // LocRoleDayShift% describes one representative week. Dividing the
    // fortnight role target by two produces the week-after-week FTE profile.
    #"Added WeekdayShift Target Minutes" = Table.AddColumn(
        #"Expanded Role Targets",
        "WeekdayShiftTargetMinutes",
        each [RoleTargetMinutes] * [#"LocRoleDayShift%"] / 2,
        type number
    ),
    // Target minutes represent productive Direct Care time. Divide by the
    // role's Direct Care percentage to obtain rostered minutes.
    #"Added WeekdayShift Roster Minutes" = Table.AddColumn(
        #"Added WeekdayShift Target Minutes",
        "WeekdayShiftRosterMinutes",
        each [WeekdayShiftTargetMinutes] / [#"Direct Care %"],
        type number
    ),
    // One FTE shift is 7.6 hours, or 456 rostered minutes.
    #"Added FTE" = Table.AddColumn(
        #"Added WeekdayShift Roster Minutes",
        "FTE",
        each [WeekdayShiftRosterMinutes] / 456,
        type number
    ),
    #"Added WeekdayShift Key" = Table.AddColumn(
        #"Added FTE",
        "WeekdayShift",
        each Text.From([DayOfWeek], "en-AU") & [#"Week Day"] & [Shift],
        type text
    ),
    #"Selected Output Columns" = Table.SelectColumns(
        #"Added WeekdayShift Key",
        {
            "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "Shift", "ShiftIndex", "WeekdayShift",
            "HistoricalRosterHours", "RoleHistoricalRosterHours", "HistoricalMinuteHours",
            "CategoryHistoricalMinuteHours", "RoleHistoryDistribution%", "LocRoleDayShift%",
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleTargetMinutes", "WeekdayShiftTargetMinutes",
            "WeekdayShiftRosterMinutes", "FTE"
        }
    ),
    #"Sorted Allocation" = Table.Sort(
        #"Selected Output Columns",
        {
            {"Facility", Order.Ascending},
            {"MinuteCategory", Order.Ascending},
            {"Role", Order.Ascending},
            {"DayOfWeek", Order.Ascending},
            {"ShiftIndex", Order.Ascending}
        }
    )
in
    #"Sorted Allocation";

// Query: MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK
// Purpose: Proves that variable daily role distributions still allocate the complete weekly productive-minute target.
// Inputs: MW DayShift Allocation.
// Output: Seven rows per facility and role, including zero-allocation days, daily variability, cumulative minutes, and weekly/fortnight reconciliation.
shared MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK =
let
    Tolerance = 0.000001,
    Source = #"MW DayShift Allocation",
    Weekdays = #table(
        type table [DayOfWeek = Int64.Type, #"Week Day" = text],
        {
            {1, "Monday"},
            {2, "Tuesday"},
            {3, "Wednesday"},
            {4, "Thursday"},
            {5, "Friday"},
            {6, "Saturday"},
            {7, "Sunday"}
        }
    ),
    RoleGrain = Table.Distinct(
        Table.SelectColumns(
            Source,
            {
                "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
                "RoleDailyTargetMinutes", "RoleTargetMinutes"
            }
        )
    ),
    // Cross-joining each role to the seven weekdays makes absent historical
    // days visible as zero rather than silently omitting them from the proof.
    #"Added Complete Week" = Table.AddColumn(
        RoleGrain,
        "Weekdays",
        each Weekdays,
        type table [DayOfWeek = Int64.Type, #"Week Day" = text]
    ),
    #"Expanded Complete Week" = Table.ExpandTableColumn(
        #"Added Complete Week",
        "Weekdays",
        {"DayOfWeek", "Week Day"},
        {"DayOfWeek", "Week Day"}
    ),
    // Shift rows are collapsed to a daily role total. DayDistribution% is
    // deliberately allowed to vary; only its seven-day sum must equal 100%.
    DailyAllocation = Table.Group(
        Source,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        {
            {
                "DayDistribution%",
                each List.Sum([#"LocRoleDayShift%"]),
                Percentage.Type
            },
            {
                "AllocatedDayProductiveMinutes",
                each List.Sum([WeekdayShiftTargetMinutes]),
                type number
            },
            {
                "AllocatedDayRosterMinutes",
                each List.Sum([WeekdayShiftRosterMinutes]),
                type number
            },
            {
                "DayFTEShiftTotal",
                each List.Sum([FTE]),
                type number
            }
        }
    ),
    #"Merged Daily Allocation" = Table.NestedJoin(
        #"Expanded Complete Week",
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        DailyAllocation,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        "DailyAllocation",
        JoinKind.LeftOuter
    ),
    #"Expanded Daily Allocation" = Table.ExpandTableColumn(
        #"Merged Daily Allocation",
        "DailyAllocation",
        {
            "DayDistribution%", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal"
        },
        {
            "DayDistribution%", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal"
        }
    ),
    #"Replaced Missing Days With Zero" = Table.ReplaceValue(
        #"Expanded Daily Allocation",
        null,
        0,
        Replacer.ReplaceValue,
        {
            "DayDistribution%", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal"
        }
    ),
    WeeklyTotals = Table.Group(
        #"Replaced Missing Days With Zero",
        {"Facility", "MinuteCategory", "Role"},
        {
            {
                "WeeklyDistributionTotal%",
                each List.Sum([#"DayDistribution%"]),
                Percentage.Type
            },
            {
                "AllocatedWeeklyProductiveMinutes",
                each List.Sum([AllocatedDayProductiveMinutes]),
                type number
            },
            {
                "AllocatedWeeklyRosterMinutes",
                each List.Sum([AllocatedDayRosterMinutes]),
                type number
            },
            {
                "WeeklyFTEShiftTotal",
                each List.Sum([DayFTEShiftTotal]),
                type number
            },
            {
                "DaysWithAllocatedMinutes",
                each Table.RowCount(
                    Table.SelectRows(_, each [AllocatedDayProductiveMinutes] <> 0)
                ),
                Int64.Type
            }
        }
    ),
    #"Merged Weekly Totals" = Table.NestedJoin(
        #"Replaced Missing Days With Zero",
        {"Facility", "MinuteCategory", "Role"},
        WeeklyTotals,
        {"Facility", "MinuteCategory", "Role"},
        "WeeklyTotals",
        JoinKind.LeftOuter
    ),
    #"Expanded Weekly Totals" = Table.ExpandTableColumn(
        #"Merged Weekly Totals",
        "WeeklyTotals",
        {
            "WeeklyDistributionTotal%", "AllocatedWeeklyProductiveMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "DaysWithAllocatedMinutes"
        },
        {
            "WeeklyDistributionTotal%", "AllocatedWeeklyProductiveMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "DaysWithAllocatedMinutes"
        }
    ),
    // Daily targets cover each calendar day, so seven daily targets are the
    // productive minutes that the representative-week distribution must absorb.
    #"Added Expected Weekly Minutes" = Table.AddColumn(
        #"Expanded Weekly Totals",
        "ExpectedWeeklyProductiveMinutes",
        each [RoleDailyTargetMinutes] * 7,
        type number
    ),
    #"Added Weekly Variance" = Table.AddColumn(
        #"Added Expected Weekly Minutes",
        "WeeklyVarianceMinutes",
        each [AllocatedWeeklyProductiveMinutes] - [ExpectedWeeklyProductiveMinutes],
        type number
    ),
    // Doubling the representative week reconstructs the original 14-day role target.
    #"Added Reconstructed Fortnight Minutes" = Table.AddColumn(
        #"Added Weekly Variance",
        "ReconstructedFortnightProductiveMinutes",
        each [AllocatedWeeklyProductiveMinutes] * 2,
        type number
    ),
    #"Added Fortnight Variance" = Table.AddColumn(
        #"Added Reconstructed Fortnight Minutes",
        "FortnightVarianceMinutes",
        each [ReconstructedFortnightProductiveMinutes] - [RoleTargetMinutes],
        type number
    ),
    #"Added Day Versus Average" = Table.AddColumn(
        #"Added Fortnight Variance",
        "DayVsAverageDailyTarget%",
        each
            if [RoleDailyTargetMinutes] = null or [RoleDailyTargetMinutes] = 0 then
                if [AllocatedDayProductiveMinutes] = 0 then 0 else null
            else
                [AllocatedDayProductiveMinutes] / [RoleDailyTargetMinutes],
        Percentage.Type
    ),
    #"Added Check Status" = Table.AddColumn(
        #"Added Day Versus Average",
        "Status",
        each
            if
                Number.Abs([#"WeeklyDistributionTotal%"] - 1) <= Tolerance and
                Number.Abs([WeeklyVarianceMinutes]) <= Tolerance and
                Number.Abs([FortnightVarianceMinutes]) <= Tolerance
            then
                "PASS"
            else
                "ERROR",
        type text
    ),
    #"Sorted For Cumulative Check" = Table.Sort(
        #"Added Check Status",
        {
            {"Facility", Order.Ascending},
            {"MinuteCategory", Order.Ascending},
            {"Role", Order.Ascending},
            {"DayOfWeek", Order.Ascending}
        }
    ),
    // The cumulative column ends at AllocatedWeeklyProductiveMinutes on Sunday,
    // making the weekly reconciliation easy to chart and visually inspect.
    #"Added Cumulative Week Minutes" = Table.AddColumn(
        #"Sorted For Cumulative Check",
        "CumulativeWeekProductiveMinutes",
        each
            let
                CurrentFacility = [Facility],
                CurrentCategory = [MinuteCategory],
                CurrentRole = [Role],
                CurrentDay = [DayOfWeek]
            in
                List.Sum(
                    Table.SelectRows(
                        #"Sorted For Cumulative Check",
                        (CheckRow) =>
                            CheckRow[Facility] = CurrentFacility and
                            CheckRow[MinuteCategory] = CurrentCategory and
                            CheckRow[Role] = CurrentRole and
                            CheckRow[DayOfWeek] <= CurrentDay
                    )[AllocatedDayProductiveMinutes]
                ),
        type number
    ),
    #"Added Cumulative Weekly Target Percentage" = Table.AddColumn(
        #"Added Cumulative Week Minutes",
        "CumulativeWeeklyTarget%",
        each
            if [ExpectedWeeklyProductiveMinutes] = null or [ExpectedWeeklyProductiveMinutes] = 0 then
                if [CumulativeWeekProductiveMinutes] = 0 then 0 else null
            else
                [CumulativeWeekProductiveMinutes] / [ExpectedWeeklyProductiveMinutes],
        Percentage.Type
    ),
    #"Added Check Message" = Table.AddColumn(
        #"Added Cumulative Weekly Target Percentage",
        "CheckMessage",
        each
            if [Status] = "PASS" then
                "Daily percentages vary, but the seven-day and reconstructed 14-day productive minutes reconcile."
            else
                "Weekly distribution or productive-minute totals do not reconcile; review the role's daily rows.",
        type text
    ),
    #"Selected Check Columns" = Table.SelectColumns(
        #"Added Check Message",
        {
            "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "DayDistribution%", "DayVsAverageDailyTarget%",
            "AllocatedDayProductiveMinutes", "CumulativeWeekProductiveMinutes",
            "CumulativeWeeklyTarget%", "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
            "RoleDailyTargetMinutes", "ExpectedWeeklyProductiveMinutes",
            "AllocatedWeeklyProductiveMinutes", "WeeklyVarianceMinutes",
            "WeeklyDistributionTotal%", "DaysWithAllocatedMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "RoleTargetMinutes",
            "ReconstructedFortnightProductiveMinutes", "FortnightVarianceMinutes",
            "Status", "CheckMessage"
        }
    )
in
    #"Selected Check Columns";

// Query: MW Input Check
// Purpose: Validates configured roles, Direct Care percentages, and RN/ALL target inputs.
// Inputs: MW MinuteWorkers Prepare and MW TargetMinutes Prepare.
// Output: Error records for invalid required inputs.
shared #"MW Input Check" =
let
    MinuteWorkers = #"MW MinuteWorkers Prepare",
    RoleCounts = Table.Group(
        MinuteWorkers,
        {"RoleKey"},
        {{"RoleCount", each Table.RowCount(_), Int64.Type}}
    ),
    InvalidRoleCounts = Table.SelectRows(
        RoleCounts,
        each [RoleKey] = null or [RoleKey] = "" or [RoleCount] <> 1
    ),
    RoleChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(InvalidRoleCounts),
            each [
                Check = "MinuteWorker role uniqueness",
                Severity = "Error",
                Facility = null,
                MinuteCategory = null,
                Role = [RoleKey],
                Actual = Number.From([RoleCount]),
                Expected = 1,
                Message = "Each non-blank MinuteWorker role must occur exactly once."
            ]
        )
    ),
    InvalidDirectCare = Table.SelectRows(
        MinuteWorkers,
        each [#"Direct Care %"] = null or [#"Direct Care %"] <= 0 or [#"Direct Care %"] > 1
    ),
    DirectCareChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(InvalidDirectCare),
            each [
                Check = "Direct Care percentage",
                Severity = "Error",
                Facility = null,
                MinuteCategory = [MinuteCategory],
                Role = [Role],
                Actual = [#"Direct Care %"],
                Expected = 1,
                Message = "Direct Care % must be greater than 0 and no greater than 100%."
            ]
        )
    ),
    TargetSummary = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {
                "Facility", "RNCount", "ALLCount", "RNNullCount", "ALLNullCount",
                "RNDailyTargetMinutes", "ALLDailyTargetMinutes", "UnexpectedTypeCount"
            }
        )
    ),
    TargetChecks = #"MW Check Table"(
        List.Combine(
            List.Transform(
                Table.ToRecords(TargetSummary),
                each List.RemoveNulls({
                    if [RNCount] <> 1 then [
                        Check = "RN target row count",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = Number.From([RNCount]),
                        Expected = 1,
                        Message = "Each facility must have exactly one RN target row."
                    ] else null,
                    if [ALLCount] <> 1 then [
                        Check = "ALL target row count",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = Number.From([ALLCount]),
                        Expected = 1,
                        Message = "Each facility must have exactly one ALL target row."
                    ] else null,
                    if [UnexpectedTypeCount] <> 0 then [
                        Check = "Target minute types",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = Number.From([UnexpectedTypeCount]),
                        Expected = 0,
                        Message = "MinuteType values must be RN or ALL."
                    ] else null,
                    if [RNNullCount] <> 0 or [RNDailyTargetMinutes] = null or [RNDailyTargetMinutes] < 0 then [
                        Check = "RN target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNDailyTargetMinutes],
                        Expected = 0,
                        Message = "RN daily target minutes must be numeric and non-negative."
                    ] else null,
                    if [ALLNullCount] <> 0 or [ALLDailyTargetMinutes] = null or [ALLDailyTargetMinutes] < 0 then [
                        Check = "ALL target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = [ALLDailyTargetMinutes],
                        Expected = 0,
                        Message = "ALL daily target minutes must be numeric and non-negative."
                    ] else null,
                    if
                        [RNDailyTargetMinutes] <> null and
                        [ALLDailyTargetMinutes] <> null and
                        [RNDailyTargetMinutes] > [ALLDailyTargetMinutes]
                    then [
                        Check = "RN target not greater than ALL",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNDailyTargetMinutes],
                        Expected = [ALLDailyTargetMinutes],
                        Message = "RN daily target minutes cannot exceed ALL daily target minutes."
                    ] else null
                })
            )
        )
    ),
    Result = Table.Combine({RoleChecks, DirectCareChecks, TargetChecks})
in
    Result;

// Query: MW Distribution Check
// Purpose: Validates history coverage and the role and day/shift distribution denominators.
// Inputs: MW Role Distribution, MW Historical DayShift, and MW TargetMinutes Prepare.
// Output: Pass, warning, and error records for historical allocation readiness.
shared #"MW Distribution Check" =
let
    Tolerance = 0.000001,
    RoleDistribution = #"MW Role Distribution",
    HistoricalDayShift = #"MW Historical DayShift",
    TargetCategories = #"MW TargetMinutes Prepare",

    RoleDistributionTotals = Table.Group(
        RoleDistribution,
        {"Facility", "MinuteCategory"},
        {{"Actual", each List.Sum([#"RoleHistoryDistribution%"]), type nullable number}}
    ),
    RoleDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(RoleDistributionTotals),
            each [
                Check = "Role history distribution totals 100%",
                Severity =
                    if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [Actual],
                Expected = 1,
                Message = "RoleHistoryDistribution% must total 100% for each facility/category."
            ]
        )
    ),

    DayShiftDistributionTotals = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory", "Role"},
        {{"Actual", each List.Sum([#"LocRoleDayShift%"]), type nullable number}}
    ),
    DayShiftDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(DayShiftDistributionTotals),
            each [
                Check = "Role day/shift distribution totals 100%",
                Severity =
                    if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = [Role],
                Actual = [Actual],
                Expected = 1,
                Message = "LocRoleDayShift% must total 100% for each facility/role."
            ]
        )
    ),

    CategoryHistory = Table.Distinct(
        Table.SelectColumns(
            RoleDistribution,
            {"Facility", "MinuteCategory", "CategoryHistoricalMinuteHours"}
        )
    ),
    TargetHistoryJoin = Table.ExpandTableColumn(
        Table.NestedJoin(
            TargetCategories,
            {"Facility", "MinuteCategory"},
            CategoryHistory,
            {"Facility", "MinuteCategory"},
            "CategoryHistory",
            JoinKind.LeftOuter
        ),
        "CategoryHistory",
        {"CategoryHistoricalMinuteHours"},
        {"CategoryHistoricalMinuteHours"}
    ),
    MissingHistory = Table.SelectRows(
        TargetHistoryJoin,
        each
            [CategoryTargetMinutes] > 0 and
            ([CategoryHistoricalMinuteHours] = null or [CategoryHistoricalMinuteHours] <= 0)
    ),
    MissingHistoryChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(MissingHistory),
            each [
                Check = "Positive target has history",
                Severity = "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual =
                    if [CategoryHistoricalMinuteHours] = null then
                        0
                    else
                        [CategoryHistoricalMinuteHours],
                Expected = [CategoryTargetMinutes],
                Message =
                    "Facility " & [Facility] & ", category " & [MinuteCategory] &
                    " has a positive 14-day target of " &
                    Text.From([CategoryTargetMinutes], "en-AU") &
                    " minutes but no eligible historical MinuteWorker hours."
            ]
        )
    ),

    MasterRoleKeys = Table.Distinct(
        Table.SelectColumns(
            Table.AddColumn(
                LocRoleWeekDaysHours,
                "RoleKey",
                each if [Role] = null then null else Text.Upper(Text.Trim([Role])),
                type nullable text
            ),
            {"RoleKey"}
        )
    ),
    UnmatchedRoles = Table.NestedJoin(
        #"MW MinuteWorkers Prepare",
        {"RoleKey"},
        MasterRoleKeys,
        {"RoleKey"},
        "HistoryRole",
        JoinKind.LeftAnti
    ),
    UnmatchedRoleChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(UnmatchedRoles),
            each [
                Check = "MinuteWorker abbreviation has history",
                Severity = "Warning",
                Facility = null,
                MinuteCategory = [MinuteCategory],
                Role = [RoleKey],
                Actual = 0,
                Expected = 1,
                Message =
                    "Configured MinuteWorker '" &
                    (if [SourceRole] = null then "" else [SourceRole]) &
                    "' maps to " &
                    (if [RoleKey] = null then "<blank>" else [RoleKey]) &
                    " but no eligible original IMPORT Master role maps to that abbreviation."
            ]
        )
    ),

    TargetFacilities = Table.Distinct(Table.SelectColumns(TargetCategories, {"Facility"})),
    HistoricalFacilities = Table.Distinct(Table.SelectColumns(CategoryHistory, {"Facility"})),
    TargetOnlyFacilities = Table.NestedJoin(
        TargetFacilities,
        {"Facility"},
        HistoricalFacilities,
        {"Facility"},
        "HistoryFacility",
        JoinKind.LeftAnti
    ),
    TargetFacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(TargetOnlyFacilities),
            each [
                Check = "Target facility has MinuteWorker history",
                Severity = "Warning",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message = "The target facility has no matching MinuteWorker history."
            ]
        )
    ),
    HistoryOnlyFacilities = Table.NestedJoin(
        HistoricalFacilities,
        {"Facility"},
        TargetFacilities,
        {"Facility"},
        "TargetFacility",
        JoinKind.LeftAnti
    ),
    HistoricalFacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(HistoryOnlyFacilities),
            each [
                Check = "Historical facility has target",
                Severity = "Warning",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message = "The historical facility has no TargetMinutes column and is omitted from the FTE output."
            ]
        )
    ),

    Result = Table.Combine(
        {
            RoleDistributionChecks,
            DayShiftDistributionChecks,
            MissingHistoryChecks,
            UnmatchedRoleChecks,
            TargetFacilityChecks,
            HistoricalFacilityChecks
        }
    )
in
    Result;

// Query: MW PreAllocation Check
// Purpose: Combines all checks that must pass before FTE outputs can be published.
// Inputs: MW Input Check and MW Distribution Check.
// Output: Validation rows; Severity Error blocks TABLE and MATRIX.
shared #"MW PreAllocation Check" =
let
    Result = Table.Combine({#"MW Input Check", #"MW Distribution Check"})
in
    Result;

// Query: MinuteWorkersFTE_TABLE
// Purpose: Publishes the validated representative-week MinuteWorker FTE allocation.
// Inputs: MW DayShift Allocation and MW PreAllocation Check.
// Output: One row per facility, role, weekday, and shift with unrounded FTE.
shared MinuteWorkersFTE_TABLE =
let
    Validation = #"MW PreAllocation Check",
    FatalValidation = Table.SelectRows(Validation, each [Severity] = "Error"),
    // Include the failing keys in the thrown message so Excel identifies the
    // exact input or history row without requiring the nested Details table.
    #"Added Fatal Error Context" = Table.AddColumn(
        FatalValidation,
        "ErrorContext",
        each Text.Combine(
            List.RemoveNulls(
                {
                    [Check],
                    if [Facility] = null then null else "Facility=" & [Facility],
                    if [MinuteCategory] = null then null else "MinuteCategory=" & [MinuteCategory],
                    if [Role] = null then null else "Role=" & [Role],
                    [Message]
                }
            ),
            " | "
        ),
        type text
    ),
    FatalMessage = Text.Combine(
        List.Distinct(#"Added Fatal Error Context"[ErrorContext]),
        "#(lf)"
    ),
    Result =
        if Table.RowCount(FatalValidation) > 0 then
            error Error.Record(
                "MinuteWorkersFTE validation failed",
                FatalMessage,
                FatalValidation
            )
        else
            #"MW DayShift Allocation"
in
    Result;

// Query: MinuteWorkersFTE_MATRIX
// Purpose: Pivots the validated table to one FTE column per weekday and shift.
// Inputs: MinuteWorkersFTE_TABLE.
// Output: One row per facility, MinuteCategory, and role.
shared MinuteWorkersFTE_MATRIX =
let
    Source = MinuteWorkersFTE_TABLE,
    MatrixSource = Table.SelectColumns(
        Source,
        {
            "Facility", "MinuteCategory", "Role", "QFR Category",
            "Direct Care %", "WeekdayShift", "FTE"
        }
    ),
    OrderedWeekdayShifts = List.Distinct(
        Table.Sort(
            Table.SelectColumns(Source, {"DayOfWeek", "ShiftIndex", "WeekdayShift"}),
            {
                {"DayOfWeek", Order.Ascending},
                {"ShiftIndex", Order.Ascending}
            }
        )[WeekdayShift]
    ),
    Result = Table.Pivot(
        MatrixSource,
        OrderedWeekdayShifts,
        "WeekdayShift",
        "FTE",
        List.Sum
    )
in
    Result;

// Query: MinuteWorkersFTE_CHECK
// Purpose: Reports input, distribution, category, and facility reconciliation results.
// Inputs: MW PreAllocation Check, MW DayShift Allocation, and MW TargetMinutes Prepare.
// Output: Validation rows proving productive minutes reconcile to RN, OTHERS, and ALL.
shared MinuteWorkersFTE_CHECK =
let
    Tolerance = 0.000001,
    PreAllocationChecks = #"MW PreAllocation Check",
    FatalValidation = Table.SelectRows(
        PreAllocationChecks,
        each [Severity] = "Error"
    ),
    Allocation = #"MW DayShift Allocation",

    // Reconstruct fortnight productive minutes from the representative-week FTE.
    AllocationWithProductiveMinutes = Table.AddColumn(
        Allocation,
        "ReconciledProductiveMinutes",
        each [FTE] * 456 * [#"Direct Care %"] * 2,
        type number
    ),
    CategoryReconciliation = Table.Group(
        AllocationWithProductiveMinutes,
        {"Facility", "MinuteCategory"},
        {
            {
                "Actual",
                each List.Sum([ReconciledProductiveMinutes]),
                type number
            },
            {
                "Expected",
                each List.Max([CategoryTargetMinutes]),
                type number
            }
        }
    ),
    CategoryChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryReconciliation),
            each [
                Check = "Allocated productive minutes reconcile to category target",
                Severity =
                    if Number.Abs([Actual] - [Expected]) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [Actual],
                Expected = [Expected],
                Message = "FTE x 456 x Direct Care % x 2 must equal the fortnight category target."
            ]
        )
    ),

    FacilityActual = Table.Group(
        AllocationWithProductiveMinutes,
        {"Facility"},
        {
            {
                "Actual",
                each List.Sum([ReconciledProductiveMinutes]),
                type number
            }
        }
    ),
    FacilityTargets = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {"Facility", "ALLFortnightTargetMinutes"}
        )
    ),
    FacilityReconciliation = Table.ExpandTableColumn(
        Table.NestedJoin(
            FacilityTargets,
            {"Facility"},
            FacilityActual,
            {"Facility"},
            "Allocation",
            JoinKind.LeftOuter
        ),
        "Allocation",
        {"Actual"},
        {"Actual"}
    ),
    FacilityChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(FacilityReconciliation),
            each [
                Check = "Facility productive minutes reconcile to ALL target",
                Severity =
                    if Number.Abs(
                        (if [Actual] = null then 0 else [Actual]) - [ALLFortnightTargetMinutes]
                    ) <= Tolerance then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = null,
                Role = null,
                Actual = if [Actual] = null then 0 else [Actual],
                Expected = [ALLFortnightTargetMinutes],
                Message = "RN plus OTHERS productive minutes must equal the facility ALL fortnight target."
            ]
        )
    ),

    Result =
        if Table.RowCount(FatalValidation) > 0 then
            PreAllocationChecks
        else
            Table.Combine({PreAllocationChecks, CategoryChecks, FacilityChecks})
in
    Result;
