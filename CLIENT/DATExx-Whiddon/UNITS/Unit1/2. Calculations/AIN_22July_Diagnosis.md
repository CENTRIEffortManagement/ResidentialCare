# AIN diagnosis: 22 July 2026

The previous morning-only diagnosis is superseded by [AIN_22July_HoursTest.md](AIN_22July_HoursTest.md), which shows every eligible BD AIN in AM, PM and NIGHT, including zero availability.

Read SourceRows for the original reason and time. UNAVAILRecordHoursOnDate is shown separately and contributes nothing to SumOfLeaveRecordHours or DistinctLeaveHoursOnDate. The distinct leave value alone determines FullDayLeaveBlocked.

Compare ExpectedEffectiveShiftHrs with ActualEffectiveShiftHrs and read AvailabilityDecision. Both missing and duplicate calculation rows fail rather than being treated as zero. A PASS establishes agreement with the revised rule, not approval of an unknown leave code or evidence that a workbook has been synchronized.

Resolve AvailabilityReasons_DIAGNOSTICS and AvailabilityRecords_Invalid separately. Unplaceable source timestamps cannot be assigned to this date's test. Workers excluded by reconciliation remain in WorkerAvailability_DIAGNOSTICS rather than the eligible shift grid.
