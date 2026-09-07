// Power Query from: Demand-MasterRoster Manual Read.xlsx
// Pathname: c:\Users\alexp\CentriNOTSYNC\ResidentialCare \CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\Demand-MasterRoster Manual Read.xlsx
// Extracted: 2026-08-29T20:58:10.990Z

section Section1;

shared #"IMPORT Master" = let
    Source = Excel.Workbook(File.Contents("C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\1. Input\Master Roster.xlsx"), null, true),
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

// Query: MinuteWorkerRoleAssignments_TABLE
// Purpose: Lists every original Master Roster role that maps to a configured MinuteWorker role.
// Inputs: IMPORT Master, MW Abbreviate Role, and MW MinuteWorkers Prepare.
// Output: One row per original and assigned role with configuration, row-count, and facility coverage.
// Notes: Unmatched Master Roster roles are excluded; use this table to audit the applied role mappings.
shared MinuteWorkerRoleAssignments_TABLE =
let
    Source = Table.SelectColumns(#"IMPORT Master", {"Location", "Role"}),
    #"Normalised Master Role Fields" = Table.TransformColumns(
        Source,
        {
            {
                "Location",
                each if _ = null then null else Text.Upper(Text.Start(Text.Trim(_), 2)),
                type nullable text
            },
            {"Role", each if _ = null then null else Text.Trim(_), type nullable text}
        }
    ),
    #"Renamed Original Role" = Table.RenameColumns(
        #"Normalised Master Role Fields",
        {{"Role", "MasterRosterRole"}}
    ),
    #"Added Assigned Role Key" = Table.AddColumn(
        #"Renamed Original Role",
        "AssignedRoleKey",
        each
            let
                AssignedRole = #"MW Abbreviate Role"([MasterRosterRole])
            in
                if AssignedRole = null then null else Text.Upper(Text.Trim(AssignedRole)),
        type nullable text
    ),
    // The inner join retains only roles that are actually configured in INPUT
    // MinuteWorkers; the expanded Role is the authoritative assigned role.
    #"Merged Configured MinuteWorkers" = Table.NestedJoin(
        #"Added Assigned Role Key",
        {"AssignedRoleKey"},
        #"MW MinuteWorkers Prepare",
        {"RoleKey"},
        "MinuteWorker",
        JoinKind.Inner
    ),
    #"Expanded MinuteWorker Assignment" = Table.ExpandTableColumn(
        #"Merged Configured MinuteWorkers",
        "MinuteWorker",
        {"Role", "MinuteCategory", "QFR Category", "Direct Care %"},
        {"AssignedMinuteWorkerRole", "MinuteCategory", "QFR Category", "Direct Care %"}
    ),
    #"Summarised Role Assignments" = Table.Group(
        #"Expanded MinuteWorker Assignment",
        {
            "MasterRosterRole", "AssignedMinuteWorkerRole", "MinuteCategory",
            "QFR Category", "Direct Care %"
        },
        {
            {"MasterRosterRowCount", each Table.RowCount(_), Int64.Type},
            {
                "FacilityCount",
                each List.Count(List.Distinct(List.RemoveNulls([Location]))),
                Int64.Type
            },
            {
                "Facilities",
                each Text.Combine(List.Sort(List.Distinct(List.RemoveNulls([Location]))), ", "),
                type text
            }
        }
    ),
    #"Sorted Role Assignments" = Table.Sort(
        #"Summarised Role Assignments",
        {
            {"AssignedMinuteWorkerRole", Order.Ascending},
            {"MasterRosterRole", Order.Ascending}
        }
    )
in
    #"Sorted Role Assignments";

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
// Purpose: Reads RN and ALL productive-care target hours per fortnight for every facility column.
// Inputs: Current-workbook table TargetMinutes.
// Output: MinuteType plus dynamically typed numeric facility columns.
// Notes: The existing TargetMinutes table name is retained; its values are hours per fortnight, confirmed 2026-09-07.
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
// Purpose: Converts RN and ALL productive-care hours per fortnight to daily and fortnight minute targets.
// Inputs: INPUT TargetMinutes.
// Output: Two rows per facility with original RN/ALL fortnight hours and daily/fortnight minutes for RN and OTHERS.
shared #"MW TargetMinutes Prepare" =
let
    Source = #"INPUT TargetMinutes",
    TargetPeriodDays = 14,
    MinutesPerHour = 60,
    FacilityColumns = List.RemoveItems(Table.ColumnNames(Source), {"MinuteType"}),
    // Expand every facility cell explicitly: Table.Unpivot drops null cells,
    // which would hide missing targets and wholly blank facility columns.
    #"Unpivoted Facility Targets" = Table.Combine(
        List.Transform(
            FacilityColumns,
            (FacilityName as text) as table =>
                Table.AddColumn(
                    Table.RenameColumns(
                        Table.SelectColumns(Source, {"MinuteType", FacilityName}),
                        {{FacilityName, "TargetFortnightHours"}}
                    ),
                    "Facility",
                    each FacilityName,
                    type text
                )
        )
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
            {"RNNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "RN" and [TargetFortnightHours] = null)), Int64.Type},
            {"ALLNullCount", each Table.RowCount(Table.SelectRows(_, each [MinuteType] = "ALL" and [TargetFortnightHours] = null)), Int64.Type},
            {
                "RNFortnightTargetHours",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "RN")[TargetFortnightHours]
                    in
                        if List.IsEmpty(Values) then null else List.Sum(Values),
                type nullable number
            },
            {
                "ALLFortnightTargetHours",
                each
                    let
                        Values = Table.SelectRows(_, each [MinuteType] = "ALL")[TargetFortnightHours]
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
    // Convert hours to minutes once, before any daily allocation. These are
    // productive-care targets; the role Direct Care % is applied downstream.
    #"Added RN Fortnight Target" = Table.AddColumn(
        #"Summarised Facility Targets",
        "RNFortnightTargetMinutes",
        each
            if [RNFortnightTargetHours] = null then
                null
            else
                [RNFortnightTargetHours] * MinutesPerHour,
        type nullable number
    ),
    #"Added ALL Fortnight Target" = Table.AddColumn(
        #"Added RN Fortnight Target",
        "ALLFortnightTargetMinutes",
        each
            if [ALLFortnightTargetHours] = null then
                null
            else
                [ALLFortnightTargetHours] * MinutesPerHour,
        type nullable number
    ),
    // Retain the fortnight target divided by 14 as an average daily reference.
    // Actual weekday allocations vary with whole-period historical weights.
    #"Added RN Daily Target" = Table.AddColumn(
        #"Added ALL Fortnight Target",
        "RNDailyTargetMinutes",
        each [RNFortnightTargetMinutes] / TargetPeriodDays,
        type nullable number
    ),
    #"Added ALL Daily Target" = Table.AddColumn(
        #"Added RN Daily Target",
        "ALLDailyTargetMinutes",
        each [ALLFortnightTargetMinutes] / TargetPeriodDays,
        type nullable number
    ),
    // ALL includes RN, so the productive-minute pool available to all other
    // MinuteWorker roles is ALL less RN at both the daily and fortnight grain.
    #"Added Category Target Records" = Table.AddColumn(
        #"Added ALL Daily Target",
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

// Query: MW Historical WeekDayShift
// Purpose: Preserves each historical week's roster hours and FTE before matching weekdays are averaged.
// Inputs: Master Prepare, MW MinuteWorkers Prepare, and ShftEnd Prepare.
// Output: One row per observed facility/role, original Week No, weekday and AM/PM/NS shift.
// Notes: Complete weeks have explicit zero cells; missing cells in incomplete weeks remain null. No target scaling is applied.
shared #"MW Historical WeekDayShift" =
let
    History = Table.Buffer(Table.AddColumn(
        Table.RenameColumns(#"Master Prepare", {{"Location", "Facility"}}),
        "DayOfWeek",
        each List.PositionOf({"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}, [#"Week Day"]) + 1,
        Int64.Type
    )),
    ValidDays = Table.SelectRows(History, each [Week No] <> null and [DayOfWeek] >= 1 and [DayOfWeek] <= 7),
    // Coverage is measured at facility/week grain before adding absent cells.
    // Week No is retained unchanged so the two Mondays can be compared directly.
    WeekCoverage = Table.Group(ValidDays, {"Facility", "Week No"}, {
        {"HistoricalDaysPresent", each List.Count(List.Distinct([DayOfWeek])), Int64.Type}
    }),
    Roles = Table.Distinct(Table.SelectColumns(History, {"Facility", "Role"})),
    RoleAttributes = Table.ExpandTableColumn(
        Table.NestedJoin(Roles, {"Role"}, #"MW MinuteWorkers Prepare", {"Role"}, "Configuration", JoinKind.Inner),
        "Configuration", {"MinuteCategory", "QFR Category", "Direct Care %"},
        {"MinuteCategory", "QFR Category", "Direct Care %"}
    ),
    RoleWeeks = Table.ExpandTableColumn(
        Table.NestedJoin(RoleAttributes, {"Facility"}, WeekCoverage, {"Facility"}, "Weeks", JoinKind.Inner),
        "Weeks", {"Week No", "HistoricalDaysPresent"}, {"Week No", "HistoricalDaysPresent"}
    ),
    Weekdays = #table(type table [DayOfWeek = Int64.Type, #"Week Day" = text], {
        {1, "Monday"}, {2, "Tuesday"}, {3, "Wednesday"}, {4, "Thursday"},
        {5, "Friday"}, {6, "Saturday"}, {7, "Sunday"}
    }),
    RoleWeekDays = Table.ExpandTableColumn(
        Table.AddColumn(RoleWeeks, "Days", each Weekdays), "Days",
        {"DayOfWeek", "Week Day"}, {"DayOfWeek", "Week Day"}
    ),
    ShiftNames = #table(type table [Shift = text], {{"AM"}, {"PM"}, {"NS"}}),
    ShiftKeys = Table.ExpandTableColumn(
        Table.NestedJoin(ShiftNames, {"Shift"}, #"ShftEnd Prepare", {"Shift"}, "ShiftOrder", JoinKind.LeftOuter),
        "ShiftOrder", {"Index"}, {"ShiftIndex"}
    ),
    CompleteGrain = Table.ExpandTableColumn(
        Table.AddColumn(RoleWeekDays, "Shifts", each ShiftKeys), "Shifts",
        {"Shift", "ShiftIndex"}, {"Shift", "ShiftIndex"}
    ),
    ObservedCells = Table.Group(ValidDays, {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, {
        {"ObservedRosterHours", each List.Sum([Roster Hours]), type nullable number},
        {"SourceRowCount", each Table.RowCount(_), Int64.Type},
        {"InvalidRosterRowCount", each Table.RowCount(Table.SelectRows(_, each [Roster Hours] = null or [Roster Hours] < 0)), Int64.Type}
    }),
    JoinedCells = Table.ExpandTableColumn(
        Table.NestedJoin(CompleteGrain,
            {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, ObservedCells,
            {"Facility", "Role", "Week No", "DayOfWeek", "Shift"}, "Observed", JoinKind.LeftOuter),
        "Observed", {"ObservedRosterHours", "SourceRowCount", "InvalidRosterRowCount"},
        {"ObservedRosterHours", "SourceRowCount", "InvalidRosterRowCount"}
    ),
    CellCounts = Table.ReplaceValue(JoinedCells, null, 0, Replacer.ReplaceValue, {"SourceRowCount", "InvalidRosterRowCount"}),
    AddedCoverageStatus = Table.AddColumn(CellCounts, "HistoricalCoverageStatus",
        each if [HistoricalDaysPresent] = 7 then "PASS" else "INCOMPLETE", type text),
    AddedCellStatus = Table.AddColumn(AddedCoverageStatus, "HistoricalCellStatus",
        each if [InvalidRosterRowCount] > 0 then "ERROR"
        else if [SourceRowCount] > 0 then "OBSERVED"
        else if [HistoricalDaysPresent] = 7 then "ZERO" else "MISSING", type text),
    AddedRosterHours = Table.AddColumn(AddedCellStatus, "HistoricalRosterHours",
        each if [HistoricalCellStatus] = "ERROR" or [HistoricalCellStatus] = "MISSING" then null
        else if [SourceRowCount] = 0 then 0 else [ObservedRosterHours], type nullable number),
    // Historical FTE measures actual roster hours, without Direct Care % or target adjustment.
    AddedRosterFTE = Table.AddColumn(AddedRosterHours, "HistoricalRosterFTE",
        each [HistoricalRosterHours] / 7.6, type nullable number),
    AddedProductiveHours = Table.AddColumn(AddedRosterFTE, "HistoricalProductiveHours",
        each [HistoricalRosterHours] * [#"Direct Care %"], type nullable number),
    AddedDayShift = Table.AddColumn(AddedProductiveHours, "DayShift", each [#"Week Day"] & [Shift], type text),
    Result = Table.Sort(Table.RemoveColumns(AddedDayShift, {"ObservedRosterHours"}), {
        {"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Week No", Order.Ascending},
        {"DayOfWeek", Order.Ascending}, {"ShiftIndex", Order.Ascending}
    })
in
    Result;

// Query: MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE
// Purpose: Exposes unaveraged historical FTE for comparing corresponding weekdays across the roster weeks.
// Inputs: MW Historical WeekDayShift.
// Output: Facility, role, original Week No, weekday, shift, historical hours/FTE, source counts and coverage flags.
// Notes: Graph weekday against HistoricalRosterFTE with Week No as the series; filter facility, role and shift. Retains all supplied weeks.
shared MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE =
let
    Result = #"MW Historical WeekDayShift"
in
    Result;

// Query: MW Historical DayShift
// Purpose: Averages corresponding historical weekdays and calculates shares across the complete representative week.
// Inputs: MW Historical WeekDayShift and MW MinuteWorkers Prepare.
// Output: One row per facility, role, weekday, and shift with role/day and category/day denominators.
// Notes: This branch does not use LocRoleTOTAL or the legacy week-normalised LocRoleDayShift%.
shared #"MW Historical DayShift" =
let
    Source = Table.SelectColumns(
        Table.RenameColumns(#"MW Historical WeekDayShift", {{"Facility", "Location"}, {"HistoricalRosterHours", "Hours"}}),
        {
            "Location", "Week No", "Role", "DayOfWeek", "Week Day",
            "Shift", "ShiftIndex", "DayShift", "Hours"
        }
    ),
    #"Normalised Facility" = Table.TransformColumns(
        Table.RenameColumns(Source, {{"Location", "Facility"}}),
        {{"Facility", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text}}
    ),
    #"Filtered Valid Historical Days" = Table.SelectRows(
        #"Normalised Facility",
        each [DayOfWeek] >= 1 and [DayOfWeek] <= 7
    ),
    BufferedHistoricalDays = Table.Buffer(#"Filtered Valid Historical Days"),
    // Count each facility's complete historical periods before grouping by
    // weekday or role. Absent role/shift cells in complete weeks contribute
    // zero rather than shrinking the denominator. Incomplete weeks and invalid
    // raw hours block publication through MW Distribution Check.
    FacilityPeriodCounts = Table.Group(
        Table.Distinct(
            Table.SelectColumns(
                BufferedHistoricalDays,
                {"Facility", "Week No"}
            )
        ),
        {"Facility"},
        {{"HistoricalWeeksInAverage", each Table.RowCount(_), Int64.Type}}
    ),
    HistoricalTotals = Table.Group(
        BufferedHistoricalDays,
        {
            "Facility", "Role", "DayOfWeek", "Week Day", "Shift",
            "ShiftIndex", "DayShift"
        },
        {
            {
                "HistoricalRosterHoursTotal",
                each List.Sum(List.RemoveNulls([Hours])),
                type number
            }
        }
    ),
    // Zero-hour cells have no historical weight and are omitted rather than
    // published as a role/day with an undefined shift denominator.
    PositiveHistoricalTotals = Table.SelectRows(
        HistoricalTotals,
        each [HistoricalRosterHoursTotal] > 0
    ),
    #"Merged Facility Period Counts" = Table.NestedJoin(
        PositiveHistoricalTotals,
        {"Facility"},
        FacilityPeriodCounts,
        {"Facility"},
        "FacilityPeriodCount",
        JoinKind.LeftOuter
    ),
    #"Expanded Facility Period Counts" = Table.ExpandTableColumn(
        #"Merged Facility Period Counts",
        "FacilityPeriodCount",
        {"HistoricalWeeksInAverage"},
        {"HistoricalWeeksInAverage"}
    ),
    #"Added Historical Roster Hours Average" = Table.AddColumn(
        #"Expanded Facility Period Counts",
        "HistoricalRosterHours",
        each
            if [HistoricalWeeksInAverage] = null or [HistoricalWeeksInAverage] = 0 then
                null
            else
                [HistoricalRosterHoursTotal] / [HistoricalWeeksInAverage],
        type nullable number
    ),
    #"Added Role Key" = Table.AddColumn(
        #"Added Historical Roster Hours Average",
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
    #"Selected Historical Base Columns" = Table.SelectColumns(
        #"Expanded MinuteWorker Attributes",
        {
            "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "Shift", "ShiftIndex", "DayShift",
            "HistoricalWeeksInAverage", "HistoricalRosterHours"
        }
    ),
    #"Added Historical Productive Hours" = Table.AddColumn(
        #"Selected Historical Base Columns",
        "HistoricalProductiveHours",
        each [HistoricalRosterHours] * [#"Direct Care %"],
        type number
    ),
    BufferedHistory = Table.Buffer(#"Added Historical Productive Hours"),
    RoleDayTotals = Table.Group(
        BufferedHistory,
        {"Facility", "MinuteCategory", "RoleKey", "DayOfWeek"},
        {
            {
                "RoleDayHistoricalRosterHours",
                each List.Sum([HistoricalRosterHours]),
                type number
            },
            {
                "RoleDayHistoricalProductiveHours",
                each List.Sum([HistoricalProductiveHours]),
                type number
            }
        }
    ),
    CategoryDayTotals = Table.Group(
        BufferedHistory,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        {
            {
                "CategoryDayHistoricalProductiveHours",
                each List.Sum([HistoricalProductiveHours]),
                type number
            }
        }
    ),
    #"Merged Role Day Totals" = Table.NestedJoin(
        BufferedHistory,
        {"Facility", "MinuteCategory", "RoleKey", "DayOfWeek"},
        RoleDayTotals,
        {"Facility", "MinuteCategory", "RoleKey", "DayOfWeek"},
        "RoleDayTotals",
        JoinKind.LeftOuter
    ),
    #"Expanded Role Day Totals" = Table.ExpandTableColumn(
        #"Merged Role Day Totals",
        "RoleDayTotals",
        {"RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours"},
        {"RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours"}
    ),
    #"Merged Category Day Totals" = Table.NestedJoin(
        #"Expanded Role Day Totals",
        {"Facility", "MinuteCategory", "DayOfWeek"},
        CategoryDayTotals,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        "CategoryDayTotals",
        JoinKind.LeftOuter
    ),
    #"Expanded Category Day Totals" = Table.ExpandTableColumn(
        #"Merged Category Day Totals",
        "CategoryDayTotals",
        {"CategoryDayHistoricalProductiveHours"},
        {"CategoryDayHistoricalProductiveHours"}
    ),
    // Within-day shares remain useful diagnostics. Allocation also uses the
    // weekday's share of the whole category week, calculated below.
    #"Added Role Day History Distribution" = Table.AddColumn(
        #"Expanded Category Day Totals",
        "RoleDayHistoryDistribution%",
        each
            if
                [CategoryDayHistoricalProductiveHours] = null or
                [CategoryDayHistoricalProductiveHours] = 0
            then
                null
            else
                [RoleDayHistoricalProductiveHours] /
                [CategoryDayHistoricalProductiveHours],
        Percentage.Type
    ),
    // Shift mix is conditional on role/day. Its absolute target can change
    // when another weekday changes the whole-period distribution denominator.
    #"Added Role Day Shift Distribution" = Table.AddColumn(
        #"Added Role Day History Distribution",
        "RoleDayShiftDistribution%",
        each
            if
                [RoleDayHistoricalRosterHours] = null or
                [RoleDayHistoricalRosterHours] = 0
            then
                null
            else
                [HistoricalRosterHours] / [RoleDayHistoricalRosterHours],
        Percentage.Type
    ),
    #"Added Category Day Role Shift Distribution" = Table.AddColumn(
        #"Added Role Day Shift Distribution",
        "CategoryDayRoleShiftDistribution%",
        each
            if
                [CategoryDayHistoricalProductiveHours] = null or
                [CategoryDayHistoricalProductiveHours] = 0
            then
                null
            else
                [HistoricalProductiveHours] /
                [CategoryDayHistoricalProductiveHours],
        Percentage.Type
    ),
    CategoryWeekTotals = Table.Group(BufferedHistory, {"Facility", "MinuteCategory"}, {
        {"CategoryWeeklyHistoricalProductiveHours", each List.Sum([HistoricalProductiveHours]), type number}
    }),
    ExpandedCategoryWeek = Table.ExpandTableColumn(
        Table.NestedJoin(#"Added Category Day Role Shift Distribution", {"Facility", "MinuteCategory"},
            CategoryWeekTotals, {"Facility", "MinuteCategory"}, "CategoryWeek", JoinKind.LeftOuter),
        "CategoryWeek", {"CategoryWeeklyHistoricalProductiveHours"}, {"CategoryWeeklyHistoricalProductiveHours"}
    ),
    // Each cell's denominator spans every role, weekday and shift within its
    // facility/category. Shares sum to one across the entire representative week.
    AddedWholeWeekShare = Table.AddColumn(ExpandedCategoryWeek, "CategoryWeekRoleDayShiftDistribution%",
        each if [CategoryWeeklyHistoricalProductiveHours] = null or [CategoryWeeklyHistoricalProductiveHours] <= 0
        then null else [HistoricalProductiveHours] / [CategoryWeeklyHistoricalProductiveHours], Percentage.Type),
    AddedWeekdayShare = Table.AddColumn(AddedWholeWeekShare, "CategoryWeekdayDistribution%",
        each if [CategoryWeeklyHistoricalProductiveHours] = null or [CategoryWeeklyHistoricalProductiveHours] <= 0
        then null else [CategoryDayHistoricalProductiveHours] / [CategoryWeeklyHistoricalProductiveHours], Percentage.Type),
    #"Selected Historical Columns" = Table.SelectColumns(
        AddedWeekdayShare,
        {
            "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "Shift", "ShiftIndex", "DayShift",
            "HistoricalWeeksInAverage", "HistoricalRosterHours", "HistoricalProductiveHours",
            "RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours",
            "CategoryDayHistoricalProductiveHours", "RoleDayHistoryDistribution%",
            "RoleDayShiftDistribution%", "CategoryDayRoleShiftDistribution%",
            "CategoryWeeklyHistoricalProductiveHours", "CategoryWeekdayDistribution%", "CategoryWeekRoleDayShiftDistribution%"
        }
    )
in
    #"Selected Historical Columns";

// Query: MW Role Distribution
// Purpose: Exposes each role's productive-minute share independently within each weekday.
// Inputs: MW Historical DayShift.
// Output: One row per facility, MinuteCategory, role, and weekday with RoleDayHistoryDistribution%.
shared #"MW Role Distribution" =
let
    Source = #"MW Historical DayShift",
    #"Selected Role Day Distribution" = Table.Distinct(
        Table.SelectColumns(
            Source,
            {
                "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category",
                "Direct Care %", "DayOfWeek", "Week Day",
                "RoleDayHistoricalRosterHours",
                "RoleDayHistoricalProductiveHours", "CategoryDayHistoricalProductiveHours",
                "RoleDayHistoryDistribution%", "CategoryWeekdayDistribution%"
            }
        )
    )
in
    #"Selected Role Day Distribution";

// Query: MW Category Weekday Targets
// Purpose: Calculates variable weekday category targets from whole-period productive-hour shares.
// Inputs: MW TargetMinutes Prepare and MW Historical DayShift.
// Output: Seven rows per facility/category with average daily minutes, weekday shares and weighted weekday target minutes.
// Notes: CategoryDailyTargetMinutes remains the seven-day average; CategoryWeekdayTargetMinutes is the actual weekday allocation target.
shared #"MW Category Weekday Targets" =
let
    Targets = Table.SelectColumns(#"MW TargetMinutes Prepare",
        {"Facility", "MinuteCategory", "CategoryDailyTargetMinutes", "CategoryTargetMinutes"}),
    Weekdays = #table(type table [DayOfWeek = Int64.Type, #"Week Day" = text], {
        {1, "Monday"}, {2, "Tuesday"}, {3, "Wednesday"}, {4, "Thursday"},
        {5, "Friday"}, {6, "Saturday"}, {7, "Sunday"}
    }),
    CompleteDays = Table.ExpandTableColumn(Table.AddColumn(Targets, "Days", each Weekdays),
        "Days", {"DayOfWeek", "Week Day"}, {"DayOfWeek", "Week Day"}),
    HistoricalDayShares = Table.Distinct(Table.SelectColumns(#"MW Historical DayShift",
        {"Facility", "MinuteCategory", "DayOfWeek", "CategoryWeekdayDistribution%"})),
    ExpandedShares = Table.ExpandTableColumn(
        Table.NestedJoin(CompleteDays, {"Facility", "MinuteCategory", "DayOfWeek"}, HistoricalDayShares,
            {"Facility", "MinuteCategory", "DayOfWeek"}, "DayShare", JoinKind.LeftOuter),
        "DayShare", {"CategoryWeekdayDistribution%"}, {"CategoryWeekdayDistribution%"}),
    // An absent category/day has zero weight. Missing complete source days or
    // a positive category target with no period history fail pre-allocation checks.
    FilledShares = Table.ReplaceValue(ExpandedShares, null, 0, Replacer.ReplaceValue, {"CategoryWeekdayDistribution%"}),
    Result = Table.AddColumn(FilledShares, "CategoryWeekdayTargetMinutes",
        each [CategoryTargetMinutes] / 2 * [#"CategoryWeekdayDistribution%"], type nullable number)
in
    Result;

// Query: MW Role Targets
// Purpose: Allocates half the fortnight category target across weekday/role combinations using whole-period historical shares.
// Inputs: MW Role Distribution and MW TargetMinutes Prepare.
// Output: One row per facility, role, and weekday with day-specific, weekly, and 14-day role targets.
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
        each [CategoryTargetMinutes] / 2 * [#"CategoryWeekdayDistribution%"] * [#"RoleDayHistoryDistribution%"],
        type number
    ),
    BufferedRoleDayTargets = Table.Buffer(#"Added Role Daily Target Minutes"),
    RoleWeeklyTargets = Table.Group(
        BufferedRoleDayTargets,
        {"Facility", "MinuteCategory", "RoleKey"},
        {
            {
                "RoleWeeklyTargetMinutes",
                each List.Sum([RoleDailyTargetMinutes]),
                type number
            }
        }
    ),
    #"Merged Role Weekly Targets" = Table.NestedJoin(
        BufferedRoleDayTargets,
        {"Facility", "MinuteCategory", "RoleKey"},
        RoleWeeklyTargets,
        {"Facility", "MinuteCategory", "RoleKey"},
        "RoleWeeklyTarget",
        JoinKind.LeftOuter
    ),
    #"Expanded Role Weekly Targets" = Table.ExpandTableColumn(
        #"Merged Role Weekly Targets",
        "RoleWeeklyTarget",
        {"RoleWeeklyTargetMinutes"},
        {"RoleWeeklyTargetMinutes"}
    ),
    #"Added Role Average Daily Target" = Table.AddColumn(
        #"Expanded Role Weekly Targets",
        "RoleAverageDailyTargetMinutes",
        each [RoleWeeklyTargetMinutes] / 7,
        type number
    ),
    // The representative week repeats twice in the 14-day target period.
    #"Added Role Target Minutes" = Table.AddColumn(
        #"Added Role Average Daily Target",
        "RoleTargetMinutes",
        each [RoleWeeklyTargetMinutes] * 2,
        type number
    )
in
    #"Added Role Target Minutes";

// Query: MW DayShift Allocation
// Purpose: Distributes the fortnight target across roles, weekdays and shifts, then outputs required average weekday roster FTE.
// Inputs: MW Historical DayShift and MW Role Targets.
// Output: One row per facility, role, weekday, and shift with unrounded FTE.
shared #"MW DayShift Allocation" =
let
    Source = #"MW Historical DayShift",
    #"Merged Role Targets" = Table.NestedJoin(
        Source,
        {"Facility", "MinuteCategory", "RoleKey", "DayOfWeek"},
        #"MW Role Targets",
        {"Facility", "MinuteCategory", "RoleKey", "DayOfWeek"},
        "RoleTarget",
        JoinKind.Inner
    ),
    #"Expanded Role Targets" = Table.ExpandTableColumn(
        #"Merged Role Targets",
        "RoleTarget",
        {
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleWeeklyTargetMinutes",
            "RoleAverageDailyTargetMinutes", "RoleTargetMinutes"
        },
        {
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleWeeklyTargetMinutes",
            "RoleAverageDailyTargetMinutes", "RoleTargetMinutes"
        }
    ),
    // Whole-week shares allocate the complete fortnight. Divide by two to
    // display the average of corresponding weekdays rather than their sum.
    #"Added WeekdayShift Target Minutes" = Table.AddColumn(
        #"Expanded Role Targets",
        "WeekdayShiftTargetMinutes",
        each
            [CategoryTargetMinutes] / 2 *
            [#"CategoryWeekRoleDayShiftDistribution%"],
        type number
    ),
    #"Added Two Stage Allocation Variance" = Table.AddColumn(
        #"Added WeekdayShift Target Minutes",
        "TwoStageAllocationVarianceMinutes",
        each
            [WeekdayShiftTargetMinutes] -
            ([RoleDailyTargetMinutes] * [#"RoleDayShiftDistribution%"]),
        type number
    ),
    // Target minutes represent productive Direct Care time. Divide by the
    // role's Direct Care percentage to obtain rostered minutes.
    #"Added WeekdayShift Roster Minutes" = Table.AddColumn(
        #"Added Two Stage Allocation Variance",
        "WeekdayShiftRosterMinutes",
        each
            if [#"Direct Care %"] = null or [#"Direct Care %"] <= 0 then
                null
            else
                [WeekdayShiftTargetMinutes] / [#"Direct Care %"],
        type nullable number
    ),
    // FTE here means 7.6-hour shift equivalents (456 rostered minutes),
    // not weekly employee FTE. Sum AM/PM/NS to obtain a full day's equivalent.
    #"Added FTE" = Table.AddColumn(
        #"Added WeekdayShift Roster Minutes",
        "FTE",
        each
            if [WeekdayShiftRosterMinutes] = null then
                null
            else
                [WeekdayShiftRosterMinutes] / 456,
        type nullable number
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
            "HistoricalWeeksInAverage", "HistoricalRosterHours", "HistoricalProductiveHours",
            "RoleDayHistoricalRosterHours", "RoleDayHistoricalProductiveHours",
            "CategoryDayHistoricalProductiveHours", "RoleDayHistoryDistribution%",
            "RoleDayShiftDistribution%", "CategoryDayRoleShiftDistribution%",
            "CategoryWeekdayDistribution%", "CategoryWeekRoleDayShiftDistribution%",
            "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
            "RoleDailyTargetMinutes", "RoleAverageDailyTargetMinutes",
            "RoleWeeklyTargetMinutes", "RoleTargetMinutes", "WeekdayShiftTargetMinutes",
            "TwoStageAllocationVarianceMinutes",
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
// Purpose: Proves every role/day allocation matches its whole-period weighted target and reconciles across the representative week and fortnight.
// Inputs: MW DayShift Allocation.
// Output: Seven rows per facility and role with daily shift-distribution, daily target, weekly, and fortnight checks.
shared MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK =
let
    Tolerance = 0.000001,
    // This check consumes the allocation more than once. Buffer only the
    // required scalar columns so the external history pipeline is evaluated
    // once without retaining unused shift-level fields in memory.
    Source = Table.Buffer(
        Table.SelectColumns(
            #"MW DayShift Allocation",
            {
                "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
                "RoleDailyTargetMinutes", "RoleAverageDailyTargetMinutes",
                "RoleWeeklyTargetMinutes", "RoleTargetMinutes", "DayOfWeek",
                "RoleDayHistoryDistribution%", "RoleDayShiftDistribution%",
                "WeekdayShiftTargetMinutes", "TwoStageAllocationVarianceMinutes",
                "WeekdayShiftRosterMinutes", "FTE"
            }
        )
    ),
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
                "RoleAverageDailyTargetMinutes", "RoleWeeklyTargetMinutes", "RoleTargetMinutes"
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
    // Shift rows are collapsed to a daily role total. Each populated role/day
    // shift distribution must total 100% independently of every other day.
    DailyAllocation = Table.Group(
        Source,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        {
            {
                "DailyShiftDistributionTotal%",
                each List.Sum([#"RoleDayShiftDistribution%"]),
                Percentage.Type
            },
            {
                "RoleDayHistoryDistribution%",
                each List.Max([#"RoleDayHistoryDistribution%"]),
                Percentage.Type
            },
            {
                "RoleDailyTargetMinutes",
                each List.Max([RoleDailyTargetMinutes]),
                type number
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
            },
            {
                "MaximumAbsoluteTwoStageVarianceMinutes",
                each List.Max(List.Transform([TwoStageAllocationVarianceMinutes], Number.Abs)),
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
            "DailyShiftDistributionTotal%", "RoleDayHistoryDistribution%",
            "RoleDailyTargetMinutes", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
            "MaximumAbsoluteTwoStageVarianceMinutes"
        },
        {
            "DailyShiftDistributionTotal%", "RoleDayHistoryDistribution%",
            "RoleDailyTargetMinutes", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
            "MaximumAbsoluteTwoStageVarianceMinutes"
        }
    ),
    #"Replaced Missing Days With Zero" = Table.ReplaceValue(
        #"Expanded Daily Allocation",
        null,
        0,
        Replacer.ReplaceValue,
        {
            "DailyShiftDistributionTotal%", "RoleDayHistoryDistribution%",
            "RoleDailyTargetMinutes", "AllocatedDayProductiveMinutes",
            "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
            "MaximumAbsoluteTwoStageVarianceMinutes"
        }
    ),
    // The completed daily table is small (seven rows per allocated role) and
    // feeds both the weekly totals and the detail join. Buffer it once to avoid
    // rebuilding the cross-join and daily aggregation for each branch.
    BufferedDailyAllocation = Table.Buffer(#"Replaced Missing Days With Zero"),
    WeeklyTotals = Table.Group(
        BufferedDailyAllocation,
        {"Facility", "MinuteCategory", "Role"},
        {
            {
                "RoleHistoryDistributionAcrossWeek%",
                each List.Sum([#"RoleDayHistoryDistribution%"]),
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
        BufferedDailyAllocation,
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
            "RoleHistoryDistributionAcrossWeek%", "AllocatedWeeklyProductiveMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "DaysWithAllocatedMinutes"
        },
        {
            "RoleHistoryDistributionAcrossWeek%", "AllocatedWeeklyProductiveMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "DaysWithAllocatedMinutes"
        }
    ),
    // Role mix can differ by weekday, so the expected week is the sum of the
    // seven weekday-specific role targets.
    #"Added Expected Weekly Minutes" = Table.AddColumn(
        #"Expanded Weekly Totals",
        "ExpectedWeeklyProductiveMinutes",
        each [RoleWeeklyTargetMinutes],
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
    #"Added Day Allocation Variance" = Table.AddColumn(
        #"Added Fortnight Variance",
        "DayAllocationVarianceMinutes",
        each [AllocatedDayProductiveMinutes] - [RoleDailyTargetMinutes],
        type number
    ),
    #"Added Day Versus Role Day Target" = Table.AddColumn(
        #"Added Day Allocation Variance",
        "DayVsRoleDayTarget%",
        each
            if [RoleDailyTargetMinutes] = null or [RoleDailyTargetMinutes] = 0 then
                if [AllocatedDayProductiveMinutes] = 0 then 0 else null
            else
                [AllocatedDayProductiveMinutes] / [RoleDailyTargetMinutes],
        Percentage.Type
    ),
    #"Added Day Versus Average" = Table.AddColumn(
        #"Added Day Versus Role Day Target",
        "DayVsAverageDailyTarget%",
        each
            if
                [RoleAverageDailyTargetMinutes] = null or
                [RoleAverageDailyTargetMinutes] = 0
            then
                if [AllocatedDayProductiveMinutes] = 0 then 0 else null
            else
                [AllocatedDayProductiveMinutes] / [RoleAverageDailyTargetMinutes],
        Percentage.Type
    ),
    #"Added Check Status" = Table.AddColumn(
        #"Added Day Versus Average",
        "Status",
        each
            if
                (
                    (
                        Number.Abs([#"RoleDayHistoryDistribution%"]) <= Tolerance and
                        Number.Abs([#"DailyShiftDistributionTotal%"]) <= Tolerance
                    ) or
                    (
                        [#"RoleDayHistoryDistribution%"] > Tolerance and
                        Number.Abs([#"DailyShiftDistributionTotal%"] - 1) <= Tolerance
                    )
                ) and
                Number.Abs([DayAllocationVarianceMinutes]) <= Tolerance and
                [MaximumAbsoluteTwoStageVarianceMinutes] <= Tolerance and
                Number.Abs([WeeklyVarianceMinutes]) <= Tolerance and
                Number.Abs([FortnightVarianceMinutes]) <= Tolerance
            then
                "PASS"
            else
                "ERROR",
        type text
    ),
    CumulativeGroupKeys = {"Facility", "MinuteCategory", "Role"},
    CumulativeInputColumns = Table.ColumnNames(#"Added Check Status"),
    CumulativeExpandColumns =
        List.RemoveItems(CumulativeInputColumns, CumulativeGroupKeys) &
        {"CumulativeWeekProductiveMinutes"},
    // Calculate the running total inside each seven-row role group. The former
    // row-by-row self-filter scanned the complete result for every output row
    // and could repeatedly trigger the full external-history dependency chain.
    #"Grouped For Cumulative Check" = Table.Group(
        #"Added Check Status",
        CumulativeGroupKeys,
        {
            {
                "RoleDays",
                (RoleRows as table) as table =>
                    let
                        SortedRoleDays = Table.Sort(
                            RoleRows,
                            {{"DayOfWeek", Order.Ascending}}
                        ),
                        DailyProductiveMinutes = List.Buffer(
                            SortedRoleDays[AllocatedDayProductiveMinutes]
                        ),
                        IndexedRoleDays = Table.AddIndexColumn(
                            SortedRoleDays,
                            "CumulativeDayIndex",
                            1,
                            1,
                            Int64.Type
                        ),
                        AddedRoleCumulativeMinutes = Table.AddColumn(
                            IndexedRoleDays,
                            "CumulativeWeekProductiveMinutes",
                            each List.Sum(
                                List.FirstN(
                                    DailyProductiveMinutes,
                                    [CumulativeDayIndex]
                                )
                            ),
                            type number
                        ),
                        Result = Table.RemoveColumns(
                            AddedRoleCumulativeMinutes,
                            {"CumulativeDayIndex"}
                        )
                    in
                        Result
            }
        }
    ),
    #"Expanded Cumulative Check" = Table.ExpandTableColumn(
        #"Grouped For Cumulative Check",
        "RoleDays",
        CumulativeExpandColumns,
        CumulativeExpandColumns
    ),
    #"Sorted Cumulative Check" = Table.Sort(
        #"Expanded Cumulative Check",
        {
            {"Facility", Order.Ascending},
            {"MinuteCategory", Order.Ascending},
            {"Role", Order.Ascending},
            {"DayOfWeek", Order.Ascending}
        }
    ),
    #"Added Cumulative Weekly Target Percentage" = Table.AddColumn(
        #"Sorted Cumulative Check",
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
                "The role/day target, daily shift distribution, representative week, and reconstructed fortnight reconcile."
            else
                "A role/day shift distribution or productive-minute total does not reconcile; review this weekday and the role totals.",
        type text
    ),
    #"Selected Check Columns" = Table.SelectColumns(
        #"Added Check Message",
        {
            "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "RoleDayHistoryDistribution%",
            "DailyShiftDistributionTotal%", "DayVsRoleDayTarget%",
            "DayVsAverageDailyTarget%", "AllocatedDayProductiveMinutes",
            "RoleDailyTargetMinutes", "DayAllocationVarianceMinutes",
            "MaximumAbsoluteTwoStageVarianceMinutes", "CumulativeWeekProductiveMinutes",
            "CumulativeWeeklyTarget%", "AllocatedDayRosterMinutes", "DayFTEShiftTotal",
            "RoleAverageDailyTargetMinutes", "ExpectedWeeklyProductiveMinutes",
            "AllocatedWeeklyProductiveMinutes", "WeeklyVarianceMinutes",
            "RoleHistoryDistributionAcrossWeek%", "DaysWithAllocatedMinutes",
            "AllocatedWeeklyRosterMinutes", "WeeklyFTEShiftTotal", "RoleTargetMinutes",
            "ReconstructedFortnightProductiveMinutes", "FortnightVarianceMinutes",
            "Status", "CheckMessage"
        }
    )
in
    #"Selected Check Columns";

// Query: MinuteWorkersFTE_CATEGORY_DAILY_CHECK
// Purpose: Reconciles RN and OTHERS daily productive minutes after summing all roles in each category.
// Inputs: MW TargetMinutes Prepare, MW Category Weekday Targets, and MW DayShift Allocation.
// Output: Seven rows per facility/category, proving weekday allocations match historical weights and the week/fortnight match the budget.
// Notes: Use this category-grain output instead of averaging role/day rows; the average is retained only as a secondary weekly audit.
shared MinuteWorkersFTE_CATEGORY_DAILY_CHECK =
let
    Tolerance = 0.000001,
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
    CategoryTargets = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {
                "Facility", "MinuteCategory", "CategoryDailyTargetMinutes",
                "CategoryTargetMinutes"
            }
        )
    ),
    // Build the complete seven-day category grain before joining allocations,
    // so an absent allocation is zero and fails if its weighted target is positive.
    #"Added Complete Week" = Table.AddColumn(
        CategoryTargets,
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
    #"Expanded Weighted Weekday Targets" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Expanded Complete Week", {"Facility", "MinuteCategory", "DayOfWeek"},
            #"MW Category Weekday Targets", {"Facility", "MinuteCategory", "DayOfWeek"}, "WeightedTarget", JoinKind.LeftOuter),
        "WeightedTarget", {"CategoryWeekdayDistribution%", "CategoryWeekdayTargetMinutes"},
        {"CategoryWeekdayDistribution%", "CategoryWeekdayTargetMinutes"}
    ),
    Allocation = Table.SelectColumns(
        #"MW DayShift Allocation",
        {
            "Facility", "MinuteCategory", "DayOfWeek",
            "WeekdayShiftTargetMinutes"
        }
    ),
    // Sum shifts and roles before calculating an average. Averaging the raw
    // role/day rows divides OTHERS by its role count and understates the target.
    DailyCategoryAllocation = Table.Group(
        Allocation,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        {
            {
                "AllocatedDayProductiveMinutes",
                each List.Sum([WeekdayShiftTargetMinutes]),
                type number
            }
        }
    ),
    #"Merged Daily Category Allocation" = Table.NestedJoin(
        #"Expanded Weighted Weekday Targets",
        {"Facility", "MinuteCategory", "DayOfWeek"},
        DailyCategoryAllocation,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        "DailyCategoryAllocation",
        JoinKind.LeftOuter
    ),
    #"Expanded Daily Category Allocation" = Table.ExpandTableColumn(
        #"Merged Daily Category Allocation",
        "DailyCategoryAllocation",
        {"AllocatedDayProductiveMinutes"},
        {"AllocatedDayProductiveMinutes"}
    ),
    #"Replaced Missing Allocation With Zero" = Table.ReplaceValue(
        #"Expanded Daily Category Allocation",
        null,
        0,
        Replacer.ReplaceValue,
        {"AllocatedDayProductiveMinutes"}
    ),
    BufferedCategoryDays = Table.Buffer(#"Replaced Missing Allocation With Zero"),
    WeeklyCategoryTotals = Table.Group(
        BufferedCategoryDays,
        {"Facility", "MinuteCategory"},
        {
            {
                "AllocatedWeeklyProductiveMinutes",
                each List.Sum([AllocatedDayProductiveMinutes]),
                type number
            }
        }
    ),
    #"Merged Weekly Category Totals" = Table.NestedJoin(
        BufferedCategoryDays,
        {"Facility", "MinuteCategory"},
        WeeklyCategoryTotals,
        {"Facility", "MinuteCategory"},
        "WeeklyCategoryTotals",
        JoinKind.LeftOuter
    ),
    #"Expanded Weekly Category Totals" = Table.ExpandTableColumn(
        #"Merged Weekly Category Totals",
        "WeeklyCategoryTotals",
        {"AllocatedWeeklyProductiveMinutes"},
        {"AllocatedWeeklyProductiveMinutes"}
    ),
    #"Added Expected Weekly Minutes" = Table.AddColumn(
        #"Expanded Weekly Category Totals",
        "ExpectedWeeklyProductiveMinutes",
        each
            if [CategoryDailyTargetMinutes] = null then
                null
            else
                [CategoryDailyTargetMinutes] * 7,
        type nullable number
    ),
    #"Added Daily Variance" = Table.AddColumn(
        #"Added Expected Weekly Minutes",
        "DailyVarianceMinutes",
        each
            if [CategoryWeekdayTargetMinutes] = null then
                null
            else
                [AllocatedDayProductiveMinutes] - [CategoryWeekdayTargetMinutes],
        type nullable number
    ),
    #"Added Average Allocated Daily Minutes" = Table.AddColumn(
        #"Added Daily Variance",
        "AverageAllocatedDailyProductiveMinutes",
        each [AllocatedWeeklyProductiveMinutes] / 7,
        type number
    ),
    #"Added Average Daily Variance" = Table.AddColumn(
        #"Added Average Allocated Daily Minutes",
        "AverageDailyVarianceMinutes",
        each
            if [CategoryDailyTargetMinutes] = null then
                null
            else
                [AverageAllocatedDailyProductiveMinutes] - [CategoryDailyTargetMinutes],
        type nullable number
    ),
    // Doubling the representative week must also reconstruct the original
    // 14-day RN or OTHERS category target.
    #"Added Reconstructed Fortnight Minutes" = Table.AddColumn(
        #"Added Average Daily Variance",
        "ReconstructedFortnightProductiveMinutes",
        each [AllocatedWeeklyProductiveMinutes] * 2,
        type number
    ),
    #"Added Fortnight Variance" = Table.AddColumn(
        #"Added Reconstructed Fortnight Minutes",
        "FortnightVarianceMinutes",
        each
            if [CategoryTargetMinutes] = null then
                null
            else
                [ReconstructedFortnightProductiveMinutes] - [CategoryTargetMinutes],
        type nullable number
    ),
    // Retain this existing field as day versus the seven-day average, not as a
    // pass/fail ratio: legitimately busier weekdays can exceed 100%.
    #"Added Day Versus Daily Target" = Table.AddColumn(
        #"Added Fortnight Variance",
        "DayVsDailyTarget%",
        each
            if [CategoryDailyTargetMinutes] = null or [CategoryDailyTargetMinutes] = 0 then
                if [AllocatedDayProductiveMinutes] = 0 then 0 else null
            else
                [AllocatedDayProductiveMinutes] / [CategoryDailyTargetMinutes],
        Percentage.Type
    ),
    #"Added Status" = Table.AddColumn(
        #"Added Day Versus Daily Target",
        "Status",
        each
            if
                [DailyVarianceMinutes] <> null and
                [AverageDailyVarianceMinutes] <> null and
                [FortnightVarianceMinutes] <> null and
                Number.Abs([DailyVarianceMinutes]) <= Tolerance and
                Number.Abs([AverageDailyVarianceMinutes]) <= Tolerance and
                Number.Abs([FortnightVarianceMinutes]) <= Tolerance
            then
                "PASS"
            else
                "ERROR",
        type text
    ),
    #"Added Check Message" = Table.AddColumn(
        #"Added Status",
        "CheckMessage",
        each
            if [Status] = "PASS" then
                "This weekday reconciles to its historical share of the period target; the weekly average and fortnight also reconcile."
            else
                "This weekday, the representative week, or the reconstructed fortnight does not reconcile; review missing history and category allocations.",
        type text
    ),
    #"Selected Check Columns" = Table.SelectColumns(
        #"Added Check Message",
        {
            "Facility", "MinuteCategory", "DayOfWeek", "Week Day",
            "AllocatedDayProductiveMinutes", "DayVsDailyTarget%",
            "DailyVarianceMinutes",
            "CategoryDailyTargetMinutes", "CategoryWeekdayDistribution%", "CategoryWeekdayTargetMinutes", "AverageAllocatedDailyProductiveMinutes",
            "AverageDailyVarianceMinutes", "ExpectedWeeklyProductiveMinutes",
            "AllocatedWeeklyProductiveMinutes", "CategoryTargetMinutes",
            "ReconstructedFortnightProductiveMinutes", "FortnightVarianceMinutes",
            "Status", "CheckMessage"
        }
    ),
    #"Sorted Category Daily Check" = Table.Sort(
        #"Selected Check Columns",
        {
            {"Facility", Order.Ascending},
            {"MinuteCategory", Order.Ascending},
            {"DayOfWeek", Order.Ascending}
        }
    )
in
    #"Sorted Category Daily Check";

// Query: MinuteWorkersFTE_HISTORICAL_DAY_CHECK
// Purpose: Compares historical daily worker counts and roster-hour FTE with the target FTE allocation at the same facility, role, and weekday grain.
// Inputs: Master Prepare and MW DayShift Allocation.
// Output: Seven rows per allocated facility/role with historical people and roster FTE beside AM, PM, NS, and total target FTE.
// Notes: Historical distinct workers are people; historical and target roster FTE are 7.6-hour equivalents. Compare FTE with FTE, not a full-day worker count with one shift's FTE.
// Notes: Week No is the available historical period key; confirm that it uniquely identifies roster weeks in the source.
shared MinuteWorkersFTE_HISTORICAL_DAY_CHECK =
let
    Tolerance = 0.000001,
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
    // Buffer one narrow allocation source because it feeds both the role grain
    // and the target aggregation. This avoids reevaluating the role-target and
    // historical-distribution chain through two independent branches.
    AllocationSource = Table.Buffer(
        Table.SelectColumns(
            #"MW DayShift Allocation",
            {
                "Facility", "MinuteCategory", "Role", "QFR Category",
                "Direct Care %", "RoleDailyTargetMinutes", "DayOfWeek", "Shift",
                "HistoricalRosterHours", "WeekdayShiftTargetMinutes",
                "WeekdayShiftRosterMinutes", "FTE"
            }
        )
    ),
    RoleGrain =
        Table.Distinct(
            Table.SelectColumns(
                AllocationSource,
                {
                    "Facility", "MinuteCategory", "Role", "QFR Category",
                    "Direct Care %"
                }
            )
        ),
    HistoricalSource =
        Table.SelectColumns(
            #"Master Prepare",
            {
                "Location", "Role", "Employee Code", "Week No", "Week Day",
                "Shift", "Roster Hours"
            }
        ),
    #"Renamed Historical Facility" = Table.RenameColumns(
        HistoricalSource,
        {{"Location", "Facility"}}
    ),
    #"Added Historical Day Number" = Table.AddColumn(
        #"Renamed Historical Facility",
        "DayOfWeek",
        each
            List.PositionOf(
                {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"},
                [#"Week Day"]
            ) + 1,
        Int64.Type
    ),
    #"Filtered Valid Historical Days" = Table.SelectRows(
        #"Added Historical Day Number",
        each [DayOfWeek] >= 1 and [DayOfWeek] <= 7
    ),
    BufferedHistoricalDays = Table.Buffer(#"Filtered Valid Historical Days"),
    FacilityWeeks =
        Table.Distinct(
            Table.SelectColumns(
                BufferedHistoricalDays,
                {"Facility", "Week No"}
            )
        ),
    FacilityWeekDayCoverage = Table.Group(
        BufferedHistoricalDays,
        {"Facility", "Week No"},
        {
            {
                "HistoricalFacilityDaysPresent",
                each List.Count(List.Distinct([DayOfWeek])),
                Int64.Type
            }
        }
    ),
    FacilityCoverageSummary = Table.Group(
        FacilityWeekDayCoverage,
        {"Facility"},
        {
            {"HistoricalFacilityWeeks", each Table.RowCount(_), Int64.Type},
            {
                "HistoricalIncompleteFacilityWeeks",
                each Table.RowCount(Table.SelectRows(_, each [HistoricalFacilityDaysPresent] < 7)),
                Int64.Type
            },
            {
                "HistoricalMinimumFacilityDaysPresent",
                each List.Min([HistoricalFacilityDaysPresent]),
                Int64.Type
            }
        }
    ),
    CountDistinctPositiveWorkers = (Rows as table, ShiftName as nullable text) as number =>
        let
            EligibleRows = Table.SelectRows(
                Rows,
                (HistoryRow) =>
                    HistoryRow[#"Employee Code"] <> null and
                    HistoryRow[#"Roster Hours"] <> null and
                    HistoryRow[#"Roster Hours"] > 0 and
                    (ShiftName = null or HistoryRow[Shift] = ShiftName)
            )
        in
            List.Count(List.Distinct(EligibleRows[#"Employee Code"])),
    // Count distinct employees once across the full day, and separately inside
    // each shift. Roster rows are retained because split rows can make a simple
    // row count differ from the number of people actually rostered.
    ObservedHistoricalWeekDays = Table.Group(
        BufferedHistoricalDays,
        {"Facility", "Role", "Week No", "DayOfWeek"},
        {
            {
                "HistoricalRosterRows",
                each Table.RowCount(_),
                Int64.Type
            },
            {
                "HistoricalPositiveRosterRows",
                each
                    Table.RowCount(
                        Table.SelectRows(
                            _,
                            each [Roster Hours] <> null and [Roster Hours] > 0
                        )
                    ),
                Int64.Type
            },
            {
                "HistoricalDistinctWorkers",
                each CountDistinctPositiveWorkers(_, null),
                Int64.Type
            },
            {
                "HistoricalRowsMissingEmployeeCode",
                each
                    Table.RowCount(
                        Table.SelectRows(
                            _,
                            each
                                [Employee Code] = null and
                                [Roster Hours] <> null and
                                [Roster Hours] > 0
                        )
                    ),
                Int64.Type
            },
            {
                "HistoricalRosterHours",
                each List.Sum(List.RemoveNulls([Roster Hours])),
                type number
            },
            {
                "HistoricalAMDistinctWorkers",
                each CountDistinctPositiveWorkers(_, "AM"),
                Int64.Type
            },
            {
                "HistoricalPMDistinctWorkers",
                each CountDistinctPositiveWorkers(_, "PM"),
                Int64.Type
            },
            {
                "HistoricalNSDistinctWorkers",
                each CountDistinctPositiveWorkers(_, "NS"),
                Int64.Type
            },
            {
                "HistoricalAMRosterHours",
                each List.Sum(List.RemoveNulls(Table.SelectRows(_, each [Shift] = "AM")[Roster Hours])),
                type number
            },
            {
                "HistoricalPMRosterHours",
                each List.Sum(List.RemoveNulls(Table.SelectRows(_, each [Shift] = "PM")[Roster Hours])),
                type number
            },
            {
                "HistoricalNSRosterHours",
                each List.Sum(List.RemoveNulls(Table.SelectRows(_, each [Shift] = "NS")[Roster Hours])),
                type number
            }
        }
    ),
    #"Merged Facility Weeks" = Table.NestedJoin(
        RoleGrain,
        {"Facility"},
        FacilityWeeks,
        {"Facility"},
        "FacilityWeeks",
        JoinKind.Inner
    ),
    #"Expanded Facility Weeks" = Table.ExpandTableColumn(
        #"Merged Facility Weeks",
        "FacilityWeeks",
        {"Week No"},
        {"Week No"}
    ),
    #"Added Complete Historical Week" = Table.AddColumn(
        #"Expanded Facility Weeks",
        "Weekdays",
        each Weekdays,
        type table [DayOfWeek = Int64.Type, #"Week Day" = text]
    ),
    #"Expanded Complete Historical Week" = Table.ExpandTableColumn(
        #"Added Complete Historical Week",
        "Weekdays",
        {"DayOfWeek", "Week Day"},
        {"DayOfWeek", "Week Day"}
    ),
    #"Merged Observed Historical Days" = Table.NestedJoin(
        #"Expanded Complete Historical Week",
        {"Facility", "Role", "Week No", "DayOfWeek"},
        ObservedHistoricalWeekDays,
        {"Facility", "Role", "Week No", "DayOfWeek"},
        "ObservedHistory",
        JoinKind.LeftOuter
    ),
    HistoricalMeasureColumns = {
        "HistoricalRosterRows", "HistoricalPositiveRosterRows", "HistoricalDistinctWorkers",
        "HistoricalRowsMissingEmployeeCode", "HistoricalRosterHours",
        "HistoricalAMDistinctWorkers", "HistoricalPMDistinctWorkers",
        "HistoricalNSDistinctWorkers", "HistoricalAMRosterHours",
        "HistoricalPMRosterHours", "HistoricalNSRosterHours"
    },
    #"Expanded Observed Historical Days" = Table.ExpandTableColumn(
        #"Merged Observed Historical Days",
        "ObservedHistory",
        HistoricalMeasureColumns,
        HistoricalMeasureColumns
    ),
    #"Replaced Missing Historical Days With Zero" = Table.ReplaceValue(
        #"Expanded Observed Historical Days",
        null,
        0,
        Replacer.ReplaceValue,
        HistoricalMeasureColumns
    ),
    HistoricalDailyAverages = Table.Group(
        #"Replaced Missing Historical Days With Zero",
        {
            "Facility", "MinuteCategory", "Role", "QFR Category",
            "Direct Care %", "DayOfWeek", "Week Day"
        },
        {
            {"HistoricalWeeksInDenominator", each Table.RowCount(_), Int64.Type},
            {
                "HistoricalWeeksWithRoster",
                each Table.RowCount(Table.SelectRows(_, each [HistoricalPositiveRosterRows] > 0)),
                Int64.Type
            },
            {"HistoricalTotalRosterRowsAcrossWeeks", each List.Sum([HistoricalRosterRows]), Int64.Type},
            {
                "HistoricalTotalPositiveRosterRowsAcrossWeeks",
                each List.Sum([HistoricalPositiveRosterRows]),
                Int64.Type
            },
            {
                "HistoricalNonPositiveOrNullRosterRowsAcrossWeeks",
                each List.Sum([HistoricalRosterRows]) - List.Sum([HistoricalPositiveRosterRows]),
                Int64.Type
            },
            {
                "HistoricalTotalWorkerOccurrencesAcrossWeeks",
                each List.Sum([HistoricalDistinctWorkers]),
                Int64.Type
            },
            {
                "HistoricalRowsMissingEmployeeCodeAcrossWeeks",
                each List.Sum([HistoricalRowsMissingEmployeeCode]),
                Int64.Type
            },
            {"HistoricalAverageDailyRosterRows", each List.Average([HistoricalRosterRows]), type number},
            {
                "HistoricalAverageDailyPositiveRosterRows",
                each List.Average([HistoricalPositiveRosterRows]),
                type number
            },
            {"HistoricalAverageDailyDistinctWorkers", each List.Average([HistoricalDistinctWorkers]), type number},
            {"HistoricalMinimumDailyDistinctWorkers", each List.Min([HistoricalDistinctWorkers]), Int64.Type},
            {"HistoricalMaximumDailyDistinctWorkers", each List.Max([HistoricalDistinctWorkers]), Int64.Type},
            {"HistoricalAverageDailyRosterHours", each List.Average([HistoricalRosterHours]), type number},
            {"HistoricalAverageAMDistinctWorkers", each List.Average([HistoricalAMDistinctWorkers]), type number},
            {"HistoricalAveragePMDistinctWorkers", each List.Average([HistoricalPMDistinctWorkers]), type number},
            {"HistoricalAverageNSDistinctWorkers", each List.Average([HistoricalNSDistinctWorkers]), type number},
            {"HistoricalAverageAMRosterHours", each List.Average([HistoricalAMRosterHours]), type number},
            {"HistoricalAveragePMRosterHours", each List.Average([HistoricalPMRosterHours]), type number},
            {"HistoricalAverageNSRosterHours", each List.Average([HistoricalNSRosterHours]), type number}
        }
    ),
    #"Merged Historical Coverage" = Table.NestedJoin(
        HistoricalDailyAverages,
        {"Facility"},
        FacilityCoverageSummary,
        {"Facility"},
        "HistoricalCoverage",
        JoinKind.LeftOuter
    ),
    #"Expanded Historical Coverage" = Table.ExpandTableColumn(
        #"Merged Historical Coverage",
        "HistoricalCoverage",
        {
            "HistoricalFacilityWeeks", "HistoricalIncompleteFacilityWeeks",
            "HistoricalMinimumFacilityDaysPresent"
        },
        {
            "HistoricalFacilityWeeks", "HistoricalIncompleteFacilityWeeks",
            "HistoricalMinimumFacilityDaysPresent"
        }
    ),
    Allocation = AllocationSource,
    TargetDailyAllocation = Table.Group(
        Allocation,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        {
            {
                "AllocatorHistoricalDailyRosterHours",
                each List.Sum([HistoricalRosterHours]),
                type number
            },
            {
                "AllocatorHistoricalAMRosterHours",
                each List.Sum(Table.SelectRows(_, each [Shift] = "AM")[HistoricalRosterHours]),
                type nullable number
            },
            {
                "AllocatorHistoricalPMRosterHours",
                each List.Sum(Table.SelectRows(_, each [Shift] = "PM")[HistoricalRosterHours]),
                type nullable number
            },
            {
                "AllocatorHistoricalNSRosterHours",
                each List.Sum(Table.SelectRows(_, each [Shift] = "NS")[HistoricalRosterHours]),
                type nullable number
            },
            {"RoleDailyTargetMinutes", each List.Max([RoleDailyTargetMinutes]), type number},
            {"TargetDailyProductiveMinutes", each List.Sum([WeekdayShiftTargetMinutes]), type number},
            {"TargetDailyRosterMinutes", each List.Sum([WeekdayShiftRosterMinutes]), type number},
            {"TargetDailyFTEShiftTotal", each List.Sum([FTE]), type number},
            {
                "TargetAMFTE",
                each List.Sum(Table.SelectRows(_, each [Shift] = "AM")[FTE]),
                type nullable number
            },
            {
                "TargetPMFTE",
                each List.Sum(Table.SelectRows(_, each [Shift] = "PM")[FTE]),
                type nullable number
            },
            {
                "TargetNSFTE",
                each List.Sum(Table.SelectRows(_, each [Shift] = "NS")[FTE]),
                type nullable number
            }
        }
    ),
    #"Merged Target Daily Allocation" = Table.NestedJoin(
        #"Expanded Historical Coverage",
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        TargetDailyAllocation,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        "TargetDailyAllocation",
        JoinKind.LeftOuter
    ),
    TargetMeasureColumns = {
        "AllocatorHistoricalDailyRosterHours", "AllocatorHistoricalAMRosterHours",
        "AllocatorHistoricalPMRosterHours", "AllocatorHistoricalNSRosterHours",
        "RoleDailyTargetMinutes", "TargetDailyProductiveMinutes",
        "TargetDailyRosterMinutes", "TargetDailyFTEShiftTotal",
        "TargetAMFTE", "TargetPMFTE", "TargetNSFTE"
    },
    #"Expanded Target Daily Allocation" = Table.ExpandTableColumn(
        #"Merged Target Daily Allocation",
        "TargetDailyAllocation",
        TargetMeasureColumns,
        TargetMeasureColumns
    ),
    #"Replaced Missing Target Allocation With Zero" = Table.ReplaceValue(
        #"Expanded Target Daily Allocation",
        null,
        0,
        Replacer.ReplaceValue,
        TargetMeasureColumns
    ),
    // Historical roster FTE and target FTE now use the same 7.6-hour unit.
    // Historical distinct workers remain visible as headcount context only.
    #"Added Historical Daily Roster FTE" = Table.AddColumn(
        #"Replaced Missing Target Allocation With Zero",
        "HistoricalAverageDailyRosterFTE",
        each [HistoricalAverageDailyRosterHours] / 7.6,
        type number
    ),
    #"Added Historical Shift Roster FTE" = Table.AddColumn(
        #"Added Historical Daily Roster FTE",
        "HistoricalShiftRosterFTE",
        each [
            AM = [HistoricalAverageAMRosterHours] / 7.6,
            PM = [HistoricalAveragePMRosterHours] / 7.6,
            NS = [HistoricalAverageNSRosterHours] / 7.6
        ],
        type [AM = number, PM = number, NS = number]
    ),
    #"Expanded Historical Shift Roster FTE" = Table.ExpandRecordColumn(
        #"Added Historical Shift Roster FTE",
        "HistoricalShiftRosterFTE",
        {"AM", "PM", "NS"},
        {"HistoricalAverageAMRosterFTE", "HistoricalAveragePMRosterFTE", "HistoricalAverageNSRosterFTE"}
    ),
    #"Added Historical Daily Productive Minutes" = Table.AddColumn(
        #"Expanded Historical Shift Roster FTE",
        "EstimatedHistoricalAverageDailyProductiveMinutesAtCurrentDirectCarePercent",
        each [HistoricalAverageDailyRosterHours] * 60 * [#"Direct Care %"],
        type number
    ),
    #"Added Allocator Historical Daily Roster FTE" = Table.AddColumn(
        #"Added Historical Daily Productive Minutes",
        "AllocatorHistoricalDailyRosterFTE",
        each [AllocatorHistoricalDailyRosterHours] / 7.6,
        type number
    ),
    #"Added Allocator Historical Shift Roster FTE" = Table.AddColumn(
        #"Added Allocator Historical Daily Roster FTE",
        "AllocatorHistoricalShiftRosterFTE",
        each [
            AM = [AllocatorHistoricalAMRosterHours] / 7.6,
            PM = [AllocatorHistoricalPMRosterHours] / 7.6,
            NS = [AllocatorHistoricalNSRosterHours] / 7.6
        ],
        type [AM = number, PM = number, NS = number]
    ),
    #"Expanded Allocator Historical Shift Roster FTE" = Table.ExpandRecordColumn(
        #"Added Allocator Historical Shift Roster FTE",
        "AllocatorHistoricalShiftRosterFTE",
        {"AM", "PM", "NS"},
        {
            "AllocatorHistoricalAMRosterFTE", "AllocatorHistoricalPMRosterFTE",
            "AllocatorHistoricalNSRosterFTE"
        }
    ),
    #"Added Historical Baseline Variance" = Table.AddColumn(
        #"Expanded Allocator Historical Shift Roster FTE",
        "AllWeeksVsAllocatorHistoricalRosterFTE",
        each [HistoricalAverageDailyRosterFTE] - [AllocatorHistoricalDailyRosterFTE],
        type number
    ),
    #"Added Target Versus Allocator Historical FTE" = Table.AddColumn(
        #"Added Historical Baseline Variance",
        "TargetVsAllocatorHistoricalRosterFTE",
        each [TargetDailyFTEShiftTotal] - [AllocatorHistoricalDailyRosterFTE],
        type number
    ),
    #"Added Target As Percent Of Allocator Historical FTE" = Table.AddColumn(
        #"Added Target Versus Allocator Historical FTE",
        "TargetAsPercentOfAllocatorHistoricalRosterFTE",
        each
            if [AllocatorHistoricalDailyRosterFTE] = 0 then
                null
            else
                [TargetDailyFTEShiftTotal] / [AllocatorHistoricalDailyRosterFTE],
        Percentage.Type
    ),
    #"Added Target Versus Allocator Shift FTE" = Table.AddColumn(
        #"Added Target As Percent Of Allocator Historical FTE",
        "TargetVsAllocatorShiftFTE",
        each [
            AM = [TargetAMFTE] - [AllocatorHistoricalAMRosterFTE],
            PM = [TargetPMFTE] - [AllocatorHistoricalPMRosterFTE],
            NS = [TargetNSFTE] - [AllocatorHistoricalNSRosterFTE]
        ],
        type [AM = number, PM = number, NS = number]
    ),
    #"Expanded Target Versus Allocator Shift FTE" = Table.ExpandRecordColumn(
        #"Added Target Versus Allocator Shift FTE",
        "TargetVsAllocatorShiftFTE",
        {"AM", "PM", "NS"},
        {
            "TargetVsAllocatorAMRosterFTE", "TargetVsAllocatorPMRosterFTE",
            "TargetVsAllocatorNSRosterFTE"
        }
    ),
    #"Added Target Versus All Weeks Historical FTE" = Table.AddColumn(
        #"Expanded Target Versus Allocator Shift FTE",
        "TargetVsAllWeeksHistoricalRosterFTE",
        each [TargetDailyFTEShiftTotal] - [HistoricalAverageDailyRosterFTE],
        type number
    ),
    #"Added Target As Percent Of All Weeks Historical FTE" = Table.AddColumn(
        #"Added Target Versus All Weeks Historical FTE",
        "TargetAsPercentOfAllWeeksHistoricalRosterFTE",
        each
            if [HistoricalAverageDailyRosterFTE] = 0 then
                null
            else
                [TargetDailyFTEShiftTotal] / [HistoricalAverageDailyRosterFTE],
        Percentage.Type
    ),
    #"Added FTE Formula Variance" = Table.AddColumn(
        #"Added Target As Percent Of All Weeks Historical FTE",
        "InternalFTEArithmeticVarianceMinutes",
        each ([TargetDailyFTEShiftTotal] * 456) - [TargetDailyRosterMinutes],
        type number
    ),
    #"Added FTE Formula Status" = Table.AddColumn(
        #"Added FTE Formula Variance",
        "InternalFTEArithmeticStatus",
        each
            if Number.Abs([InternalFTEArithmeticVarianceMinutes]) <= Tolerance then
                "PASS"
            else
                "ERROR",
        type text
    ),
    #"Added Direct Care Formula Variance" = Table.AddColumn(
        #"Added FTE Formula Status",
        "InternalDirectCareReconstructionVarianceMinutes",
        each
            ([TargetDailyFTEShiftTotal] * 456 * [#"Direct Care %"]) -
            [TargetDailyProductiveMinutes],
        type number
    ),
    #"Added Direct Care Formula Status" = Table.AddColumn(
        #"Added Direct Care Formula Variance",
        "InternalDirectCareReconstructionStatus",
        each
            if Number.Abs([InternalDirectCareReconstructionVarianceMinutes]) <= Tolerance then
                "PASS"
            else
                "ERROR",
        type text
    ),
    #"Added Role Day Target Variance" = Table.AddColumn(
        #"Added Direct Care Formula Status",
        "RoleDayTargetVarianceMinutes",
        each [TargetDailyProductiveMinutes] - [RoleDailyTargetMinutes],
        type number
    ),
    #"Added Role Day Target Status" = Table.AddColumn(
        #"Added Role Day Target Variance",
        "RoleDayTargetStatus",
        each
            if Number.Abs([RoleDayTargetVarianceMinutes]) <= Tolerance then
                "PASS"
            else
                "ERROR",
        type text
    ),
    #"Added Headcount Data Quality Status" = Table.AddColumn(
        #"Added Role Day Target Status",
        "HeadcountDataQualityStatus",
        each
            if [HistoricalRowsMissingEmployeeCodeAcrossWeeks] = 0 then
                "PASS"
            else
                "WARNING",
        type text
    ),
    #"Added Roster Row Data Quality Status" = Table.AddColumn(
        #"Added Headcount Data Quality Status",
        "RosterRowDataQualityStatus",
        each
            if [HistoricalNonPositiveOrNullRosterRowsAcrossWeeks] = 0 then
                "PASS"
            else
                "REVIEW",
        type text
    ),
    #"Added Historical Coverage Review Status" = Table.AddColumn(
        #"Added Roster Row Data Quality Status",
        "HistoricalCoverageReviewStatus",
        each
            if [HistoricalIncompleteFacilityWeeks] = 0 then
                "PASS"
            else
                "REVIEW",
        type text
    ),
    #"Added Comparison Note" = Table.AddColumn(
        #"Added Historical Coverage Review Status",
        "ComparisonNote",
        each
            "HistoricalAverageDailyDistinctWorkers counts people across the full day. " &
            "Compare TargetDailyFTEShiftTotal with AllocatorHistoricalDailyRosterFTE; " &
            "TargetAMFTE, TargetPMFTE, and TargetNSFTE are individual 7.6-hour shift equivalents, not simultaneous minimum headcount. " &
            "Each target uses this role/weekday/shift's share of whole-period productive hours; weekdays can have different totals. " &
            "REVIEW coverage means at least one facility/week has fewer than seven weekdays with eligible MinuteWorker rows.",
        type text
    ),
    #"Selected Historical Comparison Columns" = Table.SelectColumns(
        #"Added Comparison Note",
        {
            "Facility", "MinuteCategory", "Role", "QFR Category", "Direct Care %",
            "DayOfWeek", "Week Day", "HistoricalWeeksInDenominator", "HistoricalWeeksWithRoster",
            "HistoricalFacilityWeeks", "HistoricalIncompleteFacilityWeeks",
            "HistoricalMinimumFacilityDaysPresent", "HistoricalTotalRosterRowsAcrossWeeks",
            "HistoricalTotalPositiveRosterRowsAcrossWeeks",
            "HistoricalNonPositiveOrNullRosterRowsAcrossWeeks",
            "HistoricalTotalWorkerOccurrencesAcrossWeeks",
            "HistoricalRowsMissingEmployeeCodeAcrossWeeks", "HistoricalAverageDailyRosterRows",
            "HistoricalAverageDailyPositiveRosterRows",
            "HistoricalAverageDailyDistinctWorkers", "HistoricalMinimumDailyDistinctWorkers",
            "HistoricalMaximumDailyDistinctWorkers",
            "HistoricalAverageAMDistinctWorkers", "HistoricalAveragePMDistinctWorkers",
            "HistoricalAverageNSDistinctWorkers", "HistoricalAverageDailyRosterHours",
            "HistoricalAverageDailyRosterFTE", "HistoricalAverageAMRosterFTE",
            "HistoricalAveragePMRosterFTE", "HistoricalAverageNSRosterFTE",
            "EstimatedHistoricalAverageDailyProductiveMinutesAtCurrentDirectCarePercent",
            "AllocatorHistoricalDailyRosterHours", "AllocatorHistoricalAMRosterHours",
            "AllocatorHistoricalPMRosterHours", "AllocatorHistoricalNSRosterHours",
            "AllocatorHistoricalDailyRosterFTE", "AllWeeksVsAllocatorHistoricalRosterFTE",
            "AllocatorHistoricalAMRosterFTE", "AllocatorHistoricalPMRosterFTE",
            "AllocatorHistoricalNSRosterFTE",
            "RoleDailyTargetMinutes",
            "TargetDailyProductiveMinutes", "TargetDailyRosterMinutes", "TargetAMFTE",
            "TargetPMFTE", "TargetNSFTE", "TargetDailyFTEShiftTotal",
            "TargetVsAllocatorHistoricalRosterFTE", "TargetAsPercentOfAllocatorHistoricalRosterFTE",
            "TargetVsAllocatorAMRosterFTE", "TargetVsAllocatorPMRosterFTE",
            "TargetVsAllocatorNSRosterFTE",
            "TargetVsAllWeeksHistoricalRosterFTE", "TargetAsPercentOfAllWeeksHistoricalRosterFTE",
            "InternalFTEArithmeticVarianceMinutes", "InternalFTEArithmeticStatus",
            "InternalDirectCareReconstructionVarianceMinutes",
            "InternalDirectCareReconstructionStatus", "RoleDayTargetVarianceMinutes",
            "RoleDayTargetStatus", "HeadcountDataQualityStatus",
            "RosterRowDataQualityStatus", "HistoricalCoverageReviewStatus", "ComparisonNote"
        }
    ),
    #"Sorted Historical Day Check" = Table.Sort(
        #"Selected Historical Comparison Columns",
        {
            {"Facility", Order.Ascending},
            {"MinuteCategory", Order.Ascending},
            {"Role", Order.Ascending},
            {"DayOfWeek", Order.Ascending}
        }
    )
in
    #"Sorted Historical Day Check";

// Query: MW Input Check
// Purpose: Validates configured roles, Direct Care percentages, and RN/ALL target inputs.
// Inputs: MW MinuteWorkers Prepare and MW TargetMinutes Prepare.
// Output: Error records for invalid required inputs.
shared #"MW Input Check" =
let
    // A completely empty target table must fail before its facility columns
    // disappear from downstream grouping.
    TargetPresenceChecks = #"MW Check Table"(
        if Table.IsEmpty(#"INPUT TargetMinutes") then
            {[
                Check = "Fortnight target inputs are present",
                Severity = "Error",
                Actual = 0,
                Expected = 1,
                Message = "TargetMinutes must contain RN and ALL productive-care hours per fortnight."
            ]}
        else
            {}
    ),
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
                "RNFortnightTargetHours", "ALLFortnightTargetHours", "UnexpectedTypeCount"
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
                    if [RNNullCount] <> 0 or [RNFortnightTargetHours] = null or [RNFortnightTargetHours] < 0 then [
                        Check = "RN target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNFortnightTargetHours],
                        Expected = 0,
                        Message = "RN productive-care target hours per fortnight must be numeric and non-negative."
                    ] else null,
                    if [ALLNullCount] <> 0 or [ALLFortnightTargetHours] = null or [ALLFortnightTargetHours] < 0 then [
                        Check = "ALL target value",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = null,
                        Role = null,
                        Actual = [ALLFortnightTargetHours],
                        Expected = 0,
                        Message = "ALL productive-care target hours per fortnight must be numeric and non-negative."
                    ] else null,
                    if
                        [RNFortnightTargetHours] <> null and
                        [ALLFortnightTargetHours] <> null and
                        [RNFortnightTargetHours] > [ALLFortnightTargetHours]
                    then [
                        Check = "RN target not greater than ALL",
                        Severity = "Error",
                        Facility = [Facility],
                        MinuteCategory = "RN",
                        Role = null,
                        Actual = [RNFortnightTargetHours],
                        Expected = [ALLFortnightTargetHours],
                        Message = "RN fortnight target hours cannot exceed ALL fortnight target hours."
                    ] else null
                })
            )
        )
    ),
    Result = Table.Combine({TargetPresenceChecks, RoleChecks, DirectCareChecks, TargetChecks})
in
    Result;

// Query: MW Distribution Check
// Purpose: Validates historical source rows, complete periods, and conditional and whole-period distribution denominators.
// Inputs: Master Prepare, LocRoleWeekDaysHours, MW Historical DayShift, and MW TargetMinutes Prepare.
// Output: Pass, warning, and error records for historical allocation readiness.
shared #"MW Distribution Check" =
let
    Tolerance = 0.000001,
    HistoricalDayShift = Table.Buffer(#"MW Historical DayShift"),
    RoleDistribution = Table.Buffer(
        Table.Distinct(
            Table.SelectColumns(
                HistoricalDayShift,
                {
                    "Facility", "MinuteCategory", "Role", "RoleKey", "QFR Category",
                    "Direct Care %", "DayOfWeek", "Week Day",
                    "RoleDayHistoricalRosterHours",
                    "RoleDayHistoricalProductiveHours",
                    "CategoryDayHistoricalProductiveHours",
                    "RoleDayHistoryDistribution%"
                }
            )
        )
    ),
    TargetCategories = Table.Buffer(#"MW TargetMinutes Prepare"),
    HistoricalInput = Table.Buffer(
        Table.SelectColumns(
            LocRoleWeekDaysHours,
            {
                "Location", "Week No", "Role", "DayOfWeek", "Week Day",
                "Shift", "ShiftIndex", "Hours"
            }
        )
    ),
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

    // Invalid source keys and negative/null aggregated hours must fail before
    // they can create negative or incomplete allocation shares.
    // Check raw hours as well: aggregation can otherwise conceal a negative
    // adjustment or ignore a null beside valid hours in the same cell.
    InvalidRawRosterRows = Table.SelectRows(#"Master Prepare", each [Roster Hours] = null or [Roster Hours] < 0),
    RawRosterHoursChecks = #"MW Check Table"(List.Transform(Table.ToRecords(InvalidRawRosterRows), each [
        Check = "Historical raw roster hours are valid", Severity = "Error",
        Facility = [Location], MinuteCategory = null, Role = [Role],
        Actual = [Roster Hours], Expected = 0,
        Message = "Each eligible original roster row must have non-null, non-negative hours before grouping."
    ])),
    // Verify that the completed graph table preserves source totals and row
    // counts. Zero-fill must not duplicate observed cells or lose source rows.
    HistoricalGrid = Table.Buffer(#"MW Historical WeekDayShift"),
    HistoricalGridCellCounts = Table.Group(HistoricalGrid,
        {"Facility", "Role", "Week No", "DayOfWeek", "Shift"},
        {{"CellCount", each Table.RowCount(_), Int64.Type}}),
    HistoricalGridKeyChecks = #"MW Check Table"(List.Transform(
        Table.ToRecords(Table.SelectRows(HistoricalGridCellCounts, each [CellCount] <> 1)), each [
            Check = "Historical graph table has unique cells", Severity = "Error",
            Facility = [Facility], MinuteCategory = null, Role = [Role], Actual = [CellCount], Expected = 1,
            Message = "Expected exactly one historical row per facility/role/original week/weekday/shift."
        ])),
    RawPeriodTotals = Table.Group(#"Master Prepare", {"Location", "Week No"}, {
        {"SourceHours", each List.Sum([Roster Hours]), type nullable number},
        {"SourceRows", each Table.RowCount(_), Int64.Type}
    }),
    GridPeriodTotals = Table.Group(HistoricalGrid, {"Facility", "Week No"}, {
        {"GridHours", each List.Sum([HistoricalRosterHours]), type nullable number},
        {"GridSourceRows", each List.Sum([SourceRowCount]), Int64.Type}
    }),
    ComparedPeriodTotals = Table.ExpandTableColumn(
        Table.NestedJoin(RawPeriodTotals, {"Location", "Week No"}, GridPeriodTotals,
            {"Facility", "Week No"}, "Grid", JoinKind.LeftOuter),
        "Grid", {"GridHours", "GridSourceRows"}, {"GridHours", "GridSourceRows"}),
    HistoricalGridTotalChecks = #"MW Check Table"(List.Transform(Table.ToRecords(ComparedPeriodTotals), each [
        Check = "Historical graph table preserves source hours and rows",
        Severity = if [GridHours] <> null and [SourceHours] <> null and
            Number.Abs([GridHours] - [SourceHours]) <= Tolerance and [GridSourceRows] = [SourceRows]
            then "Pass" else "Error",
        Facility = [Location], MinuteCategory = null, Role = null,
        Actual = [GridHours], Expected = [SourceHours],
        Message = "Historical graph hours and contributing row counts must match the unaveraged source for week " &
            (if [Week No] = null then "<null>" else Text.From([Week No], "en-AU")) & "."
    ])),
    InvalidHistoricalInput = Table.SelectRows(
        HistoricalInput,
        each
            [Week No] = null or
            [DayOfWeek] = null or [DayOfWeek] < 1 or [DayOfWeek] > 7 or
            not List.Contains(
                {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"},
                [#"Week Day"]
            ) or
            [Shift] = null or
            not List.Contains({"AM", "PM", "NS"}, [Shift]) or
            [ShiftIndex] = null or
            [Hours] = null or [Hours] < 0
    ),
    HistoricalInputChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(InvalidHistoricalInput),
            each [
                Check = "Historical allocation source row is valid",
                Severity = "Error",
                Facility = [Location],
                MinuteCategory = null,
                Role = [Role],
                Actual = [Hours],
                Expected = 0,
                Message =
                    "History requires a non-null Week No, a recognised weekday/shift, " &
                    "a ShiftIndex, and non-negative roster hours. Week=" &
                    (if [Week No] = null then "<null>" else Text.From([Week No], "en-AU")) &
                    ", day=" &
                    (if [#"Week Day"] = null then "<null>" else [#"Week Day"]) &
                    ", shift=" &
                    (if [Shift] = null then "<null>" else [Shift]) & "."
            ]
        )
    ),
    TargetFacilityList = Table.Distinct(
        Table.SelectColumns(TargetCategories, {"Facility"})
    ),
    HistoricalFacilityPeriods = Table.Distinct(
        Table.SelectColumns(
            Table.SelectRows(HistoricalInput, each [Week No] <> null),
            {"Location", "Week No"}
        )
    ),
    TargetHistoricalPeriods = Table.RemoveColumns(
        Table.NestedJoin(
            HistoricalFacilityPeriods,
            {"Location"},
            TargetFacilityList,
            {"Facility"},
            "TargetFacility",
            JoinKind.Inner
        ),
        {"TargetFacility"}
    ),
    #"Added Expected Historical Weekdays" = Table.AddColumn(
        TargetHistoricalPeriods,
        "Weekdays",
        each Weekdays,
        type table [DayOfWeek = Int64.Type, #"Week Day" = text]
    ),
    #"Expanded Expected Historical Weekdays" = Table.ExpandTableColumn(
        #"Added Expected Historical Weekdays",
        "Weekdays",
        {"DayOfWeek", "Week Day"},
        {"DayOfWeek", "Week Day"}
    ),
    ObservedHistoricalWeekdays = Table.Distinct(
        Table.SelectColumns(
            Table.SelectRows(
                HistoricalInput,
                each
                    [Week No] <> null and
                    [DayOfWeek] <> null and
                    [DayOfWeek] >= 1 and [DayOfWeek] <= 7
            ),
            {"Location", "Week No", "DayOfWeek"}
        )
    ),
    MissingHistoricalWeekdays = Table.NestedJoin(
        #"Expanded Expected Historical Weekdays",
        {"Location", "Week No", "DayOfWeek"},
        ObservedHistoricalWeekdays,
        {"Location", "Week No", "DayOfWeek"},
        "ObservedDay",
        JoinKind.LeftAnti
    ),
    HistoricalCoverageChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(MissingHistoricalWeekdays),
            each [
                Check = "Historical facility period has every weekday",
                Severity = "Error",
                Facility = [Location],
                MinuteCategory = null,
                Role = null,
                Actual = 0,
                Expected = 1,
                Message =
                    "Week " & Text.From([Week No], "en-AU") & " has no eligible " &
                    [#"Week Day"] &
                    " MinuteWorker history; complete periods are required before corresponding weekdays are averaged."
            ]
        )
    ),

    BuildDistributionRangeChecks =
        (Rows as table, ColumnName as text, CheckName as text) as table =>
            let
                InvalidRows = Table.SelectRows(
                    Rows,
                    each
                        let
                            DistributionValue = Record.Field(_, ColumnName)
                        in
                            DistributionValue = null or
                            Number.IsNaN(DistributionValue) or
                            DistributionValue < -Tolerance or
                            DistributionValue > 1 + Tolerance
                ),
                Result = #"MW Check Table"(
                    List.Transform(
                        Table.ToRecords(InvalidRows),
                        each [
                            Check = CheckName,
                            Severity = "Error",
                            Facility = [Facility],
                            MinuteCategory = [MinuteCategory],
                            Role = [Role],
                            Actual = Record.Field(_, ColumnName),
                            Expected = 1,
                            Message =
                                ColumnName & " must be between 0% and 100% on " &
                                [#"Week Day"] & "."
                        ]
                    )
                )
            in
                Result,
    RoleDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.Distinct(
            Table.SelectColumns(
                RoleDistribution,
                {
                    "Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day",
                    "RoleDayHistoryDistribution%"
                }
            )
        ),
        "RoleDayHistoryDistribution%",
        "Role weekday history share is within range"
    ),
    ShiftDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.SelectColumns(
            HistoricalDayShift,
            {
                "Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day",
                "RoleDayShiftDistribution%"
            }
        ),
        "RoleDayShiftDistribution%",
        "Role weekday shift share is within range"
    ),
    CategoryDistributionRangeChecks = BuildDistributionRangeChecks(
        Table.SelectColumns(
            HistoricalDayShift,
            {
                "Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day",
                "CategoryDayRoleShiftDistribution%"
            }
        ),
        "CategoryDayRoleShiftDistribution%",
        "Category weekday role/shift share is within range"
    ),
    WholePeriodDistributionRangeChecks = BuildDistributionRangeChecks(
        HistoricalDayShift,
        "CategoryWeekRoleDayShiftDistribution%",
        "Category whole-period role/day/shift share is within range"
    ),
    WholePeriodDistributionTotals = Table.Group(HistoricalDayShift, {"Facility", "MinuteCategory"}, {
        {"Actual", each List.Sum([#"CategoryWeekRoleDayShiftDistribution%"]), type nullable number}
    }),
    WholePeriodDistributionChecks = #"MW Check Table"(List.Transform(Table.ToRecords(WholePeriodDistributionTotals), each [
        Check = "Category whole-period distribution totals 100%",
        Severity = if [Actual] <> null and Number.Abs([Actual] - 1) <= Tolerance then "Pass" else "Error",
        Facility = [Facility], MinuteCategory = [MinuteCategory], Role = null,
        Actual = [Actual], Expected = 1,
        Message = "CategoryWeekRoleDayShiftDistribution% must total 100% across all roles, weekdays and shifts in the representative week."
    ])),

    RoleDistributionTotals = Table.Group(
        RoleDistribution,
        {"Facility", "MinuteCategory", "DayOfWeek", "Week Day"},
        {{"Actual", each List.Sum([#"RoleDayHistoryDistribution%"]), type nullable number}}
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
                Message =
                    "RoleDayHistoryDistribution% must total 100% for " &
                    [#"Week Day"] & " independently."
            ]
        )
    ),

    DayShiftDistributionTotals = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day"},
        {{"Actual", each List.Sum([#"RoleDayShiftDistribution%"]), type nullable number}}
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
                Message =
                    "RoleDayShiftDistribution% must total 100% for role " &
                    [Role] & " on " & [#"Week Day"] & " independently."
            ]
        )
    ),

    CategoryDayDistributionTotals = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory", "DayOfWeek", "Week Day"},
        {
            {
                "Actual",
                each List.Sum([#"CategoryDayRoleShiftDistribution%"]),
                type nullable number
            }
        }
    ),
    CategoryDayDistributionChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryDayDistributionTotals),
            each [
                Check = "Category weekday role/shift distribution totals 100%",
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
                Message =
                    "CategoryDayRoleShiftDistribution% must total 100% for " &
                    [#"Week Day"] & " independently."
            ]
        )
    ),

    CategoryHistory = Table.Group(
        HistoricalDayShift,
        {"Facility", "MinuteCategory"},
        {
            {
                "CategoryWeeklyHistoricalProductiveHours",
                each List.Sum([HistoricalProductiveHours]),
                type nullable number
            }
        }
    ),
    TargetCategoryGrain = Table.Distinct(
        Table.SelectColumns(
            TargetCategories,
            {
                "Facility", "MinuteCategory", "CategoryDailyTargetMinutes",
                "CategoryTargetMinutes"
            }
        )
    ),
    // A category may legitimately have zero weight on a particular weekday.
    // Only absence of history for the entire positive-target category is fatal;
    // missing source weekdays at facility/week grain are checked separately.
    TargetHistoryJoin = Table.ExpandTableColumn(
        Table.NestedJoin(
            TargetCategoryGrain,
            {"Facility", "MinuteCategory"},
            CategoryHistory,
            {"Facility", "MinuteCategory"},
            "CategoryHistory",
            JoinKind.LeftOuter
        ),
        "CategoryHistory",
        {"CategoryWeeklyHistoricalProductiveHours"},
        {"CategoryWeeklyHistoricalProductiveHours"}
    ),
    MissingHistory = Table.SelectRows(
        TargetHistoryJoin,
        each
            [CategoryTargetMinutes] > 0 and
            (
                [CategoryWeeklyHistoricalProductiveHours] = null or
                [CategoryWeeklyHistoricalProductiveHours] <= 0
            )
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
                    if [CategoryWeeklyHistoricalProductiveHours] = null then
                        0
                    else
                        [CategoryWeeklyHistoricalProductiveHours],
                Expected = 1,
                Message =
                    "Facility " & [Facility] & ", category " & [MinuteCategory] &
                    " has a positive fortnight target of " &
                    Text.From([CategoryTargetMinutes], "en-AU") &
                    " minutes but no eligible productive MinuteWorker hours across the historical period."
            ]
        )
    ),

    MasterRoleKeys = Table.Distinct(
        Table.SelectColumns(
            Table.AddColumn(
                HistoricalInput,
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
            HistoricalInputChecks,
            RawRosterHoursChecks,
            HistoricalGridKeyChecks,
            HistoricalGridTotalChecks,
            HistoricalCoverageChecks,
            RoleDistributionRangeChecks,
            ShiftDistributionRangeChecks,
            CategoryDistributionRangeChecks,
            WholePeriodDistributionRangeChecks,
            WholePeriodDistributionChecks,
            RoleDistributionChecks,
            DayShiftDistributionChecks,
            CategoryDayDistributionChecks,
            MissingHistoryChecks,
            UnmatchedRoleChecks,
            TargetFacilityChecks,
            HistoricalFacilityChecks
        }
    )
in
    Result;

// Query: MW Daily Allocation Check
// Purpose: Blocks publication unless category/day and role/day allocations match whole-period weights and fortnight totals reconcile.
// Inputs: MW TargetMinutes Prepare, MW Category Weekday Targets, MW Role Targets, and MW DayShift Allocation.
// Output: Standard MinuteWorker check rows for weekday target and two-stage allocation invariants.
shared #"MW Daily Allocation Check" =
let
    Tolerance = 0.000001,
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
    // One narrow buffered allocation feeds both category/day and role/day
    // checks, avoiding two independent evaluations of the history chain.
    Allocation = Table.Buffer(
        Table.SelectColumns(
            #"MW DayShift Allocation",
            {
                "Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day",
                "CategoryDailyTargetMinutes", "CategoryTargetMinutes",
                "WeekdayShiftTargetMinutes", "TwoStageAllocationVarianceMinutes"
            }
        )
    ),
    CategoryTargets = Table.Distinct(
        Table.SelectColumns(
            #"MW TargetMinutes Prepare",
            {
                "Facility", "MinuteCategory", "CategoryDailyTargetMinutes",
                "CategoryTargetMinutes"
            }
        )
    ),
    #"Added Complete Category Week" = Table.AddColumn(
        CategoryTargets,
        "Weekdays",
        each Weekdays,
        type table [DayOfWeek = Int64.Type, #"Week Day" = text]
    ),
    #"Expanded Complete Category Week" = Table.ExpandTableColumn(
        #"Added Complete Category Week",
        "Weekdays",
        {"DayOfWeek", "Week Day"},
        {"DayOfWeek", "Week Day"}
    ),
    #"Expanded Weighted Category Targets" = Table.ExpandTableColumn(
        Table.NestedJoin(#"Expanded Complete Category Week", {"Facility", "MinuteCategory", "DayOfWeek"},
            #"MW Category Weekday Targets", {"Facility", "MinuteCategory", "DayOfWeek"}, "WeightedTarget", JoinKind.LeftOuter),
        "WeightedTarget", {"CategoryWeekdayTargetMinutes"}, {"CategoryWeekdayTargetMinutes"}
    ),
    AllocatedCategoryDays = Table.Group(
        Allocation,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        {
            {
                "AllocatedDayProductiveMinutes",
                each List.Sum([WeekdayShiftTargetMinutes]),
                type number
            }
        }
    ),
    #"Merged Category Day Allocation" = Table.NestedJoin(
        #"Expanded Weighted Category Targets",
        {"Facility", "MinuteCategory", "DayOfWeek"},
        AllocatedCategoryDays,
        {"Facility", "MinuteCategory", "DayOfWeek"},
        "CategoryDayAllocation",
        JoinKind.LeftOuter
    ),
    #"Expanded Category Day Allocation" = Table.ExpandTableColumn(
        #"Merged Category Day Allocation",
        "CategoryDayAllocation",
        {"AllocatedDayProductiveMinutes"},
        {"AllocatedDayProductiveMinutes"}
    ),
    #"Replaced Missing Category Day Allocation" = Table.ReplaceValue(
        #"Expanded Category Day Allocation",
        null,
        0,
        Replacer.ReplaceValue,
        {"AllocatedDayProductiveMinutes"}
    ),
    #"Added Category Day Variance" = Table.AddColumn(
        #"Replaced Missing Category Day Allocation",
        "DailyVarianceMinutes",
        each [AllocatedDayProductiveMinutes] - [CategoryWeekdayTargetMinutes],
        type number
    ),
    CategoryPeriodTotals = Table.Group(
        #"Added Category Day Variance",
        {"Facility", "MinuteCategory"},
        {
            {
                "ReconstructedFortnightProductiveMinutes",
                each List.Sum([AllocatedDayProductiveMinutes]) * 2,
                type number
            }
        }
    ),
    #"Merged Category Period Totals" = Table.NestedJoin(
        #"Added Category Day Variance",
        {"Facility", "MinuteCategory"},
        CategoryPeriodTotals,
        {"Facility", "MinuteCategory"},
        "CategoryPeriod",
        JoinKind.LeftOuter
    ),
    #"Expanded Category Period Totals" = Table.ExpandTableColumn(
        #"Merged Category Period Totals",
        "CategoryPeriod",
        {"ReconstructedFortnightProductiveMinutes"},
        {"ReconstructedFortnightProductiveMinutes"}
    ),
    CategoryDaily = Table.Buffer(
        Table.AddColumn(
            #"Expanded Category Period Totals",
            "FortnightVarianceMinutes",
            each [ReconstructedFortnightProductiveMinutes] - [CategoryTargetMinutes],
            type number
        )
    ),
    CategoryChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryDaily),
            each [
                Check = "Category weekday allocation reconciles",
                Severity =
                    if
                        [DailyVarianceMinutes] <> null and
                        Number.Abs([DailyVarianceMinutes]) <= Tolerance
                    then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [AllocatedDayProductiveMinutes],
                Expected = [CategoryWeekdayTargetMinutes],
                Message =
                    [#"Week Day"] &
                    " must receive its historical share of the representative-week category budget."
            ]
        )
    ),
    CategoryPeriodGrain = Table.Distinct(
        Table.SelectColumns(
            CategoryDaily,
            {
                "Facility", "MinuteCategory", "CategoryTargetMinutes",
                "ReconstructedFortnightProductiveMinutes", "FortnightVarianceMinutes"
            }
        )
    ),
    CategoryPeriodChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(CategoryPeriodGrain),
            each [
                Check = "Category fortnight allocation reconciles",
                Severity =
                    if
                        [FortnightVarianceMinutes] <> null and
                        Number.Abs([FortnightVarianceMinutes]) <= Tolerance
                    then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = null,
                Actual = [ReconstructedFortnightProductiveMinutes],
                Expected = [CategoryTargetMinutes],
                Message =
                    "Twice the representative week must reconstruct the 14-day category target."
            ]
        )
    ),
    AllocatedRoleDays = Table.Group(
        Allocation,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        {
            {
                "AllocatedRoleDayProductiveMinutes",
                each List.Sum([WeekdayShiftTargetMinutes]),
                type number
            },
            {
                "MaximumAbsoluteTwoStageVarianceMinutes",
                each
                    let
                        Variances = List.RemoveNulls([TwoStageAllocationVarianceMinutes])
                    in
                        if List.IsEmpty(Variances) then
                            null
                        else
                            List.Max(List.Transform(Variances, Number.Abs)),
                type nullable number
            }
        }
    ),
    RoleTargets = Table.Buffer(
        Table.SelectColumns(
            #"MW Role Targets",
            {
                "Facility", "MinuteCategory", "Role", "DayOfWeek", "Week Day",
                "RoleDailyTargetMinutes"
            }
        )
    ),
    #"Merged Expected Role Days" = Table.NestedJoin(
        RoleTargets,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        AllocatedRoleDays,
        {"Facility", "MinuteCategory", "Role", "DayOfWeek"},
        "AllocatedRoleDay",
        JoinKind.LeftOuter
    ),
    #"Expanded Expected Role Days" = Table.ExpandTableColumn(
        #"Merged Expected Role Days",
        "AllocatedRoleDay",
        {
            "AllocatedRoleDayProductiveMinutes",
            "MaximumAbsoluteTwoStageVarianceMinutes"
        },
        {
            "AllocatedRoleDayProductiveMinutes",
            "MaximumAbsoluteTwoStageVarianceMinutes"
        }
    ),
    RoleDaily = Table.Buffer(
        Table.ReplaceValue(
            #"Expanded Expected Role Days",
            null,
            0,
            Replacer.ReplaceValue,
            {
                "AllocatedRoleDayProductiveMinutes",
                "MaximumAbsoluteTwoStageVarianceMinutes"
            }
        )
    ),
    RoleChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(RoleDaily),
            each [
                Check = "Role weekday allocation reconciles",
                Severity =
                    if
                        [RoleDailyTargetMinutes] <> null and
                        Number.Abs(
                            [AllocatedRoleDayProductiveMinutes] - [RoleDailyTargetMinutes]
                        ) <= Tolerance
                    then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = [Role],
                Actual = [AllocatedRoleDayProductiveMinutes],
                Expected = [RoleDailyTargetMinutes],
                Message =
                    [Role] & " shift allocations must total its weekday-specific target on " &
                    [#"Week Day"] & "."
            ]
        )
    ),
    TwoStageChecks = #"MW Check Table"(
        List.Transform(
            Table.ToRecords(RoleDaily),
            each [
                Check = "Weekday allocation formulas agree",
                Severity =
                    if
                        [MaximumAbsoluteTwoStageVarianceMinutes] <> null and
                        [MaximumAbsoluteTwoStageVarianceMinutes] <= Tolerance
                    then
                        "Pass"
                    else
                        "Error",
                Facility = [Facility],
                MinuteCategory = [MinuteCategory],
                Role = [Role],
                Actual = [MaximumAbsoluteTwoStageVarianceMinutes],
                Expected = 0,
                Message =
                    "Direct category/day allocation and role/day then shift allocation must agree on " &
                    [#"Week Day"] & "."
            ]
        )
    ),
    Result = Table.Combine(
        {CategoryChecks, CategoryPeriodChecks, RoleChecks, TwoStageChecks}
    )
in
    Result;

// Query: MW PreAllocation Check
// Purpose: Combines input and historical-distribution checks that can run before allocation.
// Inputs: MW Input Check and MW Distribution Check.
// Output: Validation rows; Severity Error prevents the allocation checks from being evaluated.
shared #"MW PreAllocation Check" =
let
    Result = Table.Combine({#"MW Input Check", #"MW Distribution Check"})
in
    Result;

// Query: MW Fortnight Hours Check
// Purpose: Independently reconciles allocated FTE to the original RN and ALL hours-per-fortnight inputs.
// Inputs: INPUT TargetMinutes and MW DayShift Allocation; evaluate after MW PreAllocation Check passes.
// Output: RN and ALL checks per facility with Actual and Expected in productive-care hours per fortnight.
// Notes: Reads raw input hours directly, avoiding the minute conversion used to produce allocation targets.
shared #"MW Fortnight Hours Check" =
let
    // Match the existing 0.000001-minute tolerance, expressed in hours.
    ToleranceHours = 0.000001 / 60,
    RawTargets = Table.Buffer(#"INPUT TargetMinutes"),
    FacilityColumns = List.RemoveItems(Table.ColumnNames(RawTargets), {"MinuteType"}),
    NormalisedTargets = Table.TransformColumns(
        RawTargets,
        {{"MinuteType", each if _ = null then null else Text.Upper(Text.Trim(_)), type nullable text}}
    ),
    Allocation = Table.Buffer(
        Table.SelectColumns(
            #"MW DayShift Allocation",
            {"Facility", "MinuteCategory", "Direct Care %", "FTE"}
        )
    ),
    CheckRecords = List.Combine(
        List.Transform(
            FacilityColumns,
            (FacilityColumn as text) as list =>
                List.Transform(
                    {"RN", "ALL"},
                    (TargetType as text) as record =>
                        let
                            FacilityKey = Text.Upper(Text.Trim(FacilityColumn)),
                            InputValues = Table.Column(
                                Table.SelectRows(NormalisedTargets, each [MinuteType] = TargetType),
                                FacilityColumn
                            ),
                            ExpectedHours =
                                if List.Count(InputValues) = 1 then InputValues{0} else null,
                            EligibleAllocation = Table.SelectRows(
                                Allocation,
                                each [Facility] = FacilityKey and
                                    (TargetType = "ALL" or [MinuteCategory] = "RN")
                            ),
                            // A representative week repeated twice, at 7.6 roster
                            // hours per shift equivalent, reconstructs fortnight hours.
                            ReconstructedHours = List.Transform(
                                Table.ToRecords(EligibleAllocation),
                                each [FTE] * 7.6 * [#"Direct Care %"] * 2
                            ),
                            HasInvalidAllocation = List.NonNullCount(ReconstructedHours) <> List.Count(ReconstructedHours),
                            ActualHours =
                                if List.IsEmpty(ReconstructedHours) then 0 else List.Sum(ReconstructedHours),
                            Passed =
                                if ExpectedHours = null or HasInvalidAllocation then false
                                else Number.Abs(ActualHours - ExpectedHours) <= ToleranceHours
                        in
                            [
                                Check = "Allocated FTE reconciles to original fortnight hours",
                                Severity = if Passed then "Pass" else "Error",
                                Facility = FacilityKey,
                                MinuteCategory = if TargetType = "RN" then "RN" else null,
                                Role = null,
                                Actual = ActualHours,
                                Expected = ExpectedHours,
                                Message = TargetType &
                                    ": sum of representative-week FTE x 7.6 x Direct Care % x 2 must equal original TargetMinutes hours per fortnight."
                            ]
                )
        )
    ),
    Result = #"MW Check Table"(CheckRecords)
in
    Result;

// Query: MW Publication Check
// Purpose: Combines pre-allocation and post-allocation checks for a complete publication gate.
// Inputs: MW PreAllocation Check, MW Daily Allocation Check, and MW Fortnight Hours Check.
// Output: Standard validation rows; any Severity Error blocks TABLE and MATRIX.
shared #"MW Publication Check" =
let
    PreAllocationChecks = Table.Buffer(#"MW PreAllocation Check"),
    FatalPreAllocation = Table.SelectRows(
        PreAllocationChecks,
        each [Severity] = "Error"
    ),
    Result =
        if Table.RowCount(FatalPreAllocation) > 0 then
            PreAllocationChecks
        else
            Table.Combine({PreAllocationChecks, #"MW Daily Allocation Check", #"MW Fortnight Hours Check"})
in
    Result;

// Query: MinuteWorkersFTE_TABLE
// Purpose: Publishes the validated representative-week MinuteWorker FTE allocation.
// Inputs: MW Publication Check and MW DayShift Allocation.
// Output: One row per facility, role, weekday, and shift with unrounded 7.6-hour shift equivalents in FTE.
shared MinuteWorkersFTE_TABLE =
let
    RaiseValidationError = (ErrorTitle as text, ValidationRows as table) as any =>
        let
            // Include failing keys so Excel identifies the exact input,
            // historical denominator, or weekday allocation without opening Details.
            AddedErrorContext = Table.AddColumn(
                ValidationRows,
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
            ErrorMessage = Text.Combine(
                List.Distinct(AddedErrorContext[ErrorContext]),
                "#(lf)"
            )
        in
            error Error.Record(ErrorTitle, ErrorMessage, ValidationRows),
    // The publication query evaluates allocation checks only after input and
    // history checks pass, and includes reconciliation to original input hours.
    PublicationValidation = Table.Buffer(#"MW Publication Check"),
    FatalValidation = Table.SelectRows(
        PublicationValidation,
        each [Severity] = "Error"
    ),
    Result =
        if Table.RowCount(FatalValidation) > 0 then
            RaiseValidationError(
                "MinuteWorkersFTE publication validation failed",
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
    Source = Table.Buffer(
        Table.SelectColumns(
            MinuteWorkersFTE_TABLE,
            {
                "Facility", "MinuteCategory", "Role", "QFR Category",
                "Direct Care %", "DayOfWeek", "ShiftIndex", "WeekdayShift", "FTE"
            }
        )
    ),
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
// Inputs: MW PreAllocation Check, MW Daily Allocation Check, MW Fortnight Hours Check, MW DayShift Allocation, and MW TargetMinutes Prepare.
// Output: Validation rows reconciling RN/OTHERS/ALL minutes plus independent RN/ALL original fortnight hours.
shared MinuteWorkersFTE_CHECK =
let
    Tolerance = 0.000001,
    PreAllocationChecks = Table.Buffer(#"MW PreAllocation Check"),
    FatalValidation = Table.SelectRows(
        PreAllocationChecks,
        each [Severity] = "Error"
    ),
    DailyAllocationChecks = #"MW Daily Allocation Check",

    // Reconstruct fortnight productive minutes from the representative-week FTE.
    AllocationWithProductiveMinutes = Table.Buffer(
        Table.AddColumn(
            Table.SelectColumns(
                #"MW DayShift Allocation",
                {
                    "Facility", "MinuteCategory", "CategoryTargetMinutes",
                    "Direct Care %", "FTE"
                }
            ),
            "ReconciledProductiveMinutes",
            each [FTE] * 456 * [#"Direct Care %"] * 2,
            type number
        )
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
            Table.Combine(
                {
                    PreAllocationChecks,
                    DailyAllocationChecks,
                    #"MW Fortnight Hours Check",
                    CategoryChecks,
                    FacilityChecks
                }
            )
in
    Result;
