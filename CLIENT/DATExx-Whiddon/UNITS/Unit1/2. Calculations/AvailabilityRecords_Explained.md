# AvailabilityRecords: separate reason categories

Updated 19 September 2026. The authoritative source is [Capacity-ShiftAvailability.xlsx_PowerQuery.m](Capacity-ShiftAvailability.xlsx_PowerQuery.m). The four-worker refresh matched the independent day/shift matrix on all 336 rows. The sample filters have since been removed from the source; the full BD result is under review after a user-run Unit1 refresh.

`IMPORT Availability Leave Source` reads the BD extraction. `AvailabilityRecords` preserves SourceRecord, OriginalReason, parsed Start/End, worker/facility keys, RecordKind and Issue.

| RecordKind | Classification | Effect |
| --- | --- | --- |
| UNAVAIL | Case-sensitive substring UNAVAIL, checked first | Subtract only overlapping shift time; never daily leave hours |
| AVAIL | Otherwise case-sensitive substring AVAIL | Restrict its calendar date to offered windows |
| Leave | Exact description in AvailabilityLeaveReasons, ignoring letter case and surrounding spaces | Count distinct daily leave; subtract shift overlap below the threshold |
| Unknown | Missing or unmatched reason | Review in AvailabilityReasons_DIAGNOSTICS; eligible-worker failures block publication |

The explicit leave catalogue contains 14 recognised absence descriptions. It includes the original four examples plus the ten descriptions confirmed by the 197-record diagnosis: workers compensation (Not Worked Sec 40, Sec 37 unfit, Sec 36 unfit), personal sick leave (paid/unpaid), personal carer leave (paid/unpaid), Long Service Lve, Compassionate Lve Paid and Personal Emergency Leave (PEL). These contribute to the recognised daily absence total; UNAVAIL does not. See AvailabilityLeaveReasons in the source for exact labels. Extend that table when another genuine absence description is confirmed. Unknown descriptions never silently become available or leave.

The compatibility field Available is true for AVAIL, false for UNAVAIL/Leave and null for Unknown. **Do not use Available=false to calculate daily leave totals.** Use RecordKind=Leave.

Invalid IDs, unknown reasons, missing/invalid timestamps or end<=start produce an Issue. AvailabilityRecords_Invalid restricts required input failures to eligible workers plus records without a usable ID. Worker eligibility remains the responsibility of ReconciledWorkers_Prepare and ReconciledWorkers_Eligible.

For inspectable records and independent shift comparisons, use [AIN_22July_HoursTest.md](AIN_22July_HoursTest.md).
