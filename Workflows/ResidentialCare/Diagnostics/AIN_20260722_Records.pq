// Query: AIN_20260722_Records
// Purpose: Show original source fields and counted durations for BD AIN records overlapping 22 July 2026.
// Inputs: ReconciledWorkers_Prepare, AvailabilityRecords and IMPORT Combined Availabilities.
// Output: Worker/source-record rows, including a placeholder for workers with no dated record.
// Notes: Keeps true, false and invalid dated records visible; only valid false intervals contribute excluded hours.
let
    TestDate = #date(2026, 7, 22),
    DayStart = DateTime.From(TestDate),
    DayEnd = DayStart + #duration(1, 0, 0, 0),
    Workers = Table.SelectRows(ReconciledWorkers_Prepare,
        each [Role] = "AIN" and [#"Facility-Abbrev"] = "BD"),
    WorkerIssues = Table.RenameColumns(Workers, {{"Issue", "EligibilityIssue"}}),

    // Include records that began earlier and continue onto this date. Midnight end is exclusive.
    DatedRecords = Table.SelectRows(AvailabilityRecords,
        each [Start] <> null and [End] <> null and [Start] < DayEnd and [End] > DayStart),
    MatchedOriginals = Table.NestedJoin(DatedRecords, {"SourceRecord", "Payroll Code", "Facility-Abbrev"},
        #"IMPORT Combined Availabilities", {"SourceRecord", "Payroll Code", "Facility-Abbrev"},
        "OriginalRows", JoinKind.LeftOuter),
    OriginalCounts = Table.AddColumn(MatchedOriginals, "OriginalMatchCount",
        each Table.RowCount([OriginalRows]), Int64.Type),
    // Extract one original record without multiplying rows if a join unexpectedly has duplicates.
    OriginalFields = Table.AddColumn(OriginalCounts, "Original",
        each if [OriginalMatchCount] = 1 then [OriginalRows]{0} else
            [#"Date From" = null, #"Date To" = null, #"Availability or Leave Reason" = null], type record),
    RawColumns = Table.ExpandRecordColumn(OriginalFields, "Original",
        {"Date From", "Date To", "Availability or Leave Reason"},
        {"OriginalDateFrom", "OriginalDateTo", "OriginalReason"}),
    MatchedWorkers = Table.NestedJoin(WorkerIssues, {"EmployeeID", "Facility-Abbrev"},
        RawColumns, {"Payroll Code", "Facility-Abbrev"}, "DatedRecords", JoinKind.LeftOuter),
    ExpandedRecords = Table.ExpandTableColumn(MatchedWorkers, "DatedRecords",
        {"SourceRecord", "Start", "End", "Available", "Issue", "OriginalMatchCount",
         "OriginalDateFrom", "OriginalDateTo", "OriginalReason"},
        {"SourceRecord", "Start", "End", "Available", "RecordIssue", "OriginalMatchCount",
         "OriginalDateFrom", "OriginalDateTo", "OriginalReason"}),

    // Preserve the complete recorded duration beside the part that falls on 22 July.
    OriginalHours = Table.AddColumn(ExpandedRecords, "RecordedDurationHours",
        each if [SourceRecord] = null then null else Duration.TotalHours([End] - [Start]), type nullable number),
    ClippedStart = Table.AddColumn(OriginalHours, "StartOn22July",
        each if [SourceRecord] = null then null else List.Max({[Start], DayStart}), type nullable datetime),
    ClippedEnd = Table.AddColumn(ClippedStart, "EndOn22July",
        each if [SourceRecord] = null then null else List.Min({[End], DayEnd}), type nullable datetime),
    ClippedHours = Table.AddColumn(ClippedEnd, "RecordHoursOn22July",
        each if [SourceRecord] = null then null
            else Duration.TotalHours([EndOn22July] - [StartOn22July]), type nullable number),
    ExclusionFlag = Table.AddColumn(ClippedHours, "CountsAsExclusion",
        each [SourceRecord] <> null and [RecordIssue] = null and [Available] = false, type logical),
    CountedHours = Table.AddColumn(ExclusionFlag, "ExcludedRecordHours",
        each if [CountsAsExclusion] then [RecordHoursOn22July] else 0, type number),
    SourceCheck = Table.AddColumn(CountedHours, "SourceCheck",
        each if [SourceRecord] = null then "No record overlapping this date"
            else if [OriginalMatchCount] <> 1 then "FAIL: original source match is not unique"
            else if [RecordIssue] <> null then "Invalid dated record: " & [RecordIssue]
            else "OK", type text),
    Readable = Table.ReorderColumns(SourceCheck,
        {"EmployeeID", "BaseName", "SourceRecord", "OriginalReason", "Available",
         "OriginalDateFrom", "OriginalDateTo", "Start", "End", "RecordedDurationHours",
         "StartOn22July", "EndOn22July", "RecordHoursOn22July", "CountsAsExclusion",
         "ExcludedRecordHours", "SourceCheck", "EligibilityIssue"})
in
    Table.Sort(Readable, {{"BaseName", Order.Ascending}, {"EmployeeID", Order.Ascending}, {"Start", Order.Ascending}})
