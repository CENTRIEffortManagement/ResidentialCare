# Unit1 availability version repair plan

## Implementation status, 19 September 2026

The user reports that they ran Unit1 after the sample filters were removed and is checking the results. The generated actual-run note for `20260919-165249-04a88aab` reports **Completed** technical refresh, explicitly without business reconciliation. The full-population Availability rows, AIN lineage and Capacity outputs have not yet been independently reconciled here. Keep the four-worker CSV as the earlier bounded-test evidence.

The Availability workbook was synchronized, refreshed and saved at 16:32. Read-only extraction confirms that its embedded M matches the adjacent authoritative `.m`. Its 336 four-worker shift measurements and 108 positive `ResDayShift` rows match the independent matrix with zero differences; all 12 saved publication checks pass. The row-by-row explanation is in `outputs/four-worker-availability/availability-walkthrough.md`, with actual-versus-expected rows in `day_shift_matrix_reconciled.csv`. StaffListMaster and AIN A.1/A.2/B remain to be refreshed in order before judging Capacity.

Following that review, the user authorized removal of the sample filters. The authoritative Availability `.m` now has no four-ID filters in its reconciliation, contract or availability imports. At the last source gate before this edit, the saved workbook still held the filtered test version. The later Unit1 run's embedded M and full BD outputs have not been checked in this note. The original bounded-test steps below are retained as the historical plan.

The user confirmed that the adjacent `.m` source takes precedence over the workbook's embedded M. Both exact workbook/source pairs were snapshotted read-only in `outputs/four-worker-availability/`. Availability's pair matched at the source gate; AIN B's pair did not. The existing `.m` files were edited without replacing B's `.m` from its embedded workbook definition.

The Availability `.m` has the revised leave-only daily rule and a single canonical path resolver. Its four-worker filters were retained through the bounded test and then removed as described above. The B `.m` checks positive A.2 C# against positive A.1 original availability before publishing C###. `day_shift_matrix_before.csv` preserves the old result; `day_shift_matrix_reconciled.csv` shows the 336-row revised-rule expectation alongside the saved filtered test result. The rule document was corrected.

Both workbooks expose `FilePathUrl` as a named cell. An intermediate Excel refresh displayed `FilePathUrl must have one row and a FilePath column` because the edited Availability resolver accepted only a table. The Availability and B `.m` resolvers now accept Excel's `Column1` named-cell shape and normalize it to `FilePath` before applying the same mapping rules. No path table needs to be installed in these workbooks. The four-worker Availability runtime check has passed; same-run downstream Capacity and the full BD population remain open.

## How to judge the repair

The current disagreement is between saved results from different times. Do not call it a continuing A.1/A.2/B defect until those workbooks have been refreshed in order from the same Availability result. Once the separately authorized source sync is complete, refresh `Capacity-ShiftAvailability` with the four-worker filter, then `StaffListMaster`, AIN A.1, A.2 and B in that order. A.1 reads both Availability and StaffListMaster, so StaffListMaster belongs in the chain. Refresh `Capacity` afterwards if the question is whether its published output agrees too. Follow the runner's declared prerequisites and record each saved input/output version.

Expected outcome: A.1 should receive the new positive `ResDayShift` rows, A.2 should cap those same rows, and B's positive C# values should have matching positive A.1 originals. This is a hypothesis until the ordered run and row-level comparison pass. A.1/A.2/B alone would remove their stale-version disagreement, but if Availability still contains its old zero-hour result, that run cannot test the revised availability rule or establish positive capacity for these workers. The B check should pass on a consistent chain; if it fails, inspect its Role/Resource/Period diagnostic rows. A successful refresh receipt alone is insufficient: compare the four worker IDs across `ResDayShift`, A.1, A.2, B and, when refreshed, `Capacity`.

## Decision and scope

Repair the Power Query definition in the existing `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx`. Preserve that workbook, its tables, connections, unrelated queries, current contract work, and its exact input-file choices. Then add a consistency check to the existing Unit1/AIN `CapacityDistrib(B)-shifts.xlsx`. Do not replace either workbook with a backup or roll an entire old M section over current code. AIN A.1, A.2, StaffListMaster, and Capacity are refresh consumers; they are not business-logic edit targets in this repair unless validation finds a separate defect.

The proposed availability behavior is the version previously embedded in the workbook: `AVAIL` restricts only the calendar dates it overlaps; `UNAVAIL` subtracts only its actual shift overlap; recognised leave alone contributes to the full-day leave threshold; unknown reasons fail validation. This is described in `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AvailabilityRecords_Explained.md`. `docs/DATExx-Whiddon-Capacity-Shift-Availability-Rules.md` currently says `UNAVAIL` contributes to the full-day threshold. Resolve that contradiction in the same change so the implemented rule and one current explanation agree.

The evidence is in `outputs/four-worker-availability/README.md` and the 18 September batch `CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20260918-093026-9249c1e7/`. That complete 60-job run succeeded. Today's saved workbooks are from different later states: the current filtered Availability output has zero `ResDayShift` rows for the four workers, A.1 has no original availability for them, while A.2 retains positive C# and B/Capacity retain positive C###. The revised availability code was embedded through 17 September and had reverted in the workbook by 19 September.

## 1. Establish exact source of truth

1. Record Git status, timestamps and hashes for the two exact target workbooks and their M sources. Keep existing user edits intact. Confirm both workbooks are closed before any source extraction or synchronization.
2. Re-extract Power Query from the exact current `Capacity-ShiftAvailability.xlsx` and verify its definitions against the adjacent `Capacity-ShiftAvailability.xlsx_PowerQuery.m`. This pair matched during the read-only diagnosis; recheck at implementation time because the workbook has been saved during this conversation.
3. Re-extract the exact current AIN `CapacityDistrib(B)-shifts.xlsx` before editing its source. Its embedded M differs from the adjacent `.m` in several import and contract definitions, even though the relevant C### calculation matches. Record the difference, then edit the authoritative `.m` without replacing it from the workbook.
4. Record the approved work files for this repair. The repository normally keeps product M source under `Workflows/`, as configured in `pq.project.json`. The user's `.m` precedence makes the exact adjacent `.m` files the narrow authoritative exceptions for these existing workbook-linked edits. Do not edit a backup, copy, or generated same-stem file.

## 2. Repair the existing availability query

1. Merge only the revised availability classification, daily AVAIL baseline, leave-only full-day conversion, and their validation checks from the last verified revised implementation into the current source. Retain the September contract and identity work, Settings imports, public query names and output columns, and unrelated calculations.
2. Keep the four-worker local filters for the first bounded repair test. Correct the Worker's Report sample filter so it includes 19822 once alongside 19835, 24562, and 25376. Preserve these filters in the existing workbook through the four-worker review. Remove them only as a separate, reviewed step before full-population validation and a full Unit1 run; do not leave four hard-coded IDs in the final production query.
3. Audit every external import in the edited source against the ResidentialCare `FilePathUrl` → `CentriSyncPaths` resolver convention. Accept the existing one-cell `FilePathUrl` name as `Column1`, normalize it to `FilePath`, and preserve the selected Whiddon client, Unit1, and source files.
4. Put the revised logic in named preparation, daily-rule, shift-segment, output, and check queries with current `// Query:` and `// Purpose:` headers and comments on the threshold, interval boundaries, and row grain. Keep the existing `Availability-StaffList` and `ResDayShift` interfaces stable.

## 3. Stop mixed-version Capacity publication

1. In the existing AIN B source, add a check that every positive imported A.2 C# `Role`/`Resource`/`Period` has exactly one matching positive original A.1 availability row for the same key. Report missing and duplicate keys with Role, Resource and Period; fail `C###TABLE B` publication when the invariant fails. Keep the existing C##/C### arithmetic and contract caps unchanged.
2. Use today's saved mismatch as a historical negative case: Resources 10, 12, 26 and 27 have positive A.2 C# without A.1 originals. Do not judge the repaired chain from that mixed snapshot. The check must fail on an inconsistent fixture and pass after a consistent ordered rerun. Keep it connection-only when workbook support allows.
3. Because B is a workbook-linked M edit, apply the same external-import and path-resolver review to its exact approved `.m` source. Preserve the source/workbook version gap as evidence until separately authorized synchronization.

## 4. Validate source changes before workbook synchronization

1. Run the installed M parser and source checks. Review every edited query's headers, dependencies, comments, step names, final `in`, output interface, and external file access. Verify the approved resolver and report any unverified `FilePathUrl` prerequisites.
2. Test interval cases: `UNAVAIL` before/after a shift, touching endpoints, overlapping and duplicate records, partial and full-day recognised leave, unknown reason, AVAIL on one date but not the next, and overnight shifts. A complete shift should receive the configured 7.6 effective hours; a partial shift should use the greater of zero and the smaller of residual clock hours and the allowance after distinct shift absence hours.
3. Preserve the current `outputs/four-worker-availability/day_shift_matrix.csv` as a clearly named before snapshot. Update its independent calculation script for the revised AVAIL/UNAVAIL/leave rules, then create an expected four-worker matrix with source-record evidence. On 20 July, worker 24562's PM and worker 25376's AM and PM are specific cases that should no longer disappear solely because non-overlapping `UNAVAIL` time passed 7.6 hours. Worker 19835's recorded full-day `UNAVAIL` should still exclude that day.
4. After an authorized sync and **four-worker Availability refresh**, regenerate the main `day_shift_matrix.csv` from the exact saved workbook. Include expected versus actual available and effective hours, the `ResDayShift` publication result, and an explicit match/failure column for every worker/date/shift. Keep the before snapshot. Then refresh StaffListMaster, AIN A.1, A.2 and B before judging downstream capacity.
5. Once the four-worker matrix is reviewed, remove the local filters in a separate source edit and validate the complete BD population, published-name uniqueness, contracts, invalid-record gates, unique worker/shift rows, and positive-hours publication. Compare counts and representative workers with the last coherent full-run evidence. Do not treat a syntax pass or cached worksheet values as a runtime pass.

## 5. Apply to the existing workbooks in separate opt-in steps

1. **Synchronization:** only after an explicit instruction naming each exact workbook and approved M source, confirm the workbook is closed, validate the source, sync with the approved Power Query mechanism, and re-extract to compare exact query definitions. Synchronization does not refresh.
2. **Four-worker ordered test:** only after an explicit refresh instruction, refresh the exact Availability workbook with its four-worker filter and update the matrix as described above. Then refresh StaffListMaster, AIN A.1, A.2 and B in order before judging the downstream result. Keep the four-worker filter throughout this bounded test.
3. **Full-population refresh:** after the separate filter removal, approved source sync, and an explicit refresh instruction, run the named Unit1 chain in dependency order: availability inputs and `Capacity-ShiftAvailability`, `StaffListMaster`, AIN A.1, A.2, B, then `Capacity`. Include any runner-declared prerequisites; use the repository runner instructions and receipts. A full suite is an option when the intended scope is the whole workflow.
4. Compare the run's input stamps and output rows at each stage. Confirm the four IDs' `ResDayShift` rows, A.1 original availability, A.2 C#, B C###, and Capacity rows reconcile at worker/date/shift or Resource/Period grain. Require all availability and B checks to pass. Record the run ID and exact saved files.

## Completion conditions

- The existing Availability workbook contains the approved revised definitions, and re-extraction matches the approved M source.
- The four-worker filter is retained through the first test, and the refreshed `day_shift_matrix.csv` matches independently checked AVAIL, UNAVAIL and leave intervals. The old-rule CSV remains available as a before snapshot.
- After the separate filter-removal step, the full BD output passes identity, reason, segment and publication checks.
- The existing AIN B workbook rejects a positive C# row without matching original A.1 availability, and its embedded M matches its approved source after sync.
- One ordered refresh receipt shows matching input versions through Capacity, with no positive Capacity row arising solely from an older saved A.2 result.
- The rule documents agree with the installed code. The four-ID filters remain during the bounded test and are absent from production imports only after the separately reviewed full-population step.
