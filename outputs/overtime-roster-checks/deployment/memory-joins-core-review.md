# OvertimeRosterChecks core: memory and join review

The first memory amendment did not resolve the full ExcessSegments refresh: its retry was stopped unsaved after renewed private-memory growth. Five bounded stage probes then completed. The current amendment changes only the ExcessSegments leaf: an eleven-column shift projection and buffers of three-column fortnight status and seven-column period parts. Business fixtures passed and the approved source was synchronized and re-extracted successfully. The third full-refresh attempt has started; its outcome remains pending.

Reviewed 5 November 2026, Australia/Sydney. Current source and line references below refer to `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq`, SHA-256 `23F8DA4322DB02598D9581B743FE16C412A52407412AB98DDF7C5F00335AF882`. The original memory-pressure baseline used source `31B28974A1B2D111A367DE2D16954E2361A84A759B7EF1243C64BAA55CA4E66C`; the stopped first-amendment retry and five probes used `AB3868C13B052B0D1925CAF54592EFBB68598F6B98A2A0732AC93FDC9E3142A8`. This audit read source and saved evidence only; the audit itself did not operate a workbook, refresh or process.

## Observed runtime versus source evidence

Saved process snapshots establish the memory pressure directly:

| Sydney time, 5 November | Four Mashup containers, total logged PrivateMB | Mashup working sets | Excel PrivateMB / WorkingSetMB |
|---|---:|---|---|
| 08:21:15 | 22,200.0 | 252.6–256.0 MB each | 1,568.0 / 1,639.1 |
| 08:22:19, before owned-attempt stop | 25,241.2 | 252.0–256.0 MB each | 1,568.2 / 1,639.4 |

Evidence: [memory during the first refresh](20261104-202652/memory-during-first-refresh.json) and [memory before the agent stop](20261104-202652/memory-before-agent-stop.json). The four containers added 3,041.2 logged MB in approximately 64 seconds. The snapshots also record sustained CPU accumulation; the supervising agent observed active disk I/O and no blocking dialog. These are not an allocation trace. Private memory and working set are different measures; a small working set does not establish a small committed-memory footprint or a configured memory limit.

The refresh log identifies the first loaded query as `ExcessSegments/OvertimeRosterChecks_ExcessSegments`. No error or blocking dialog was reported before the supervisor requested stopping the owned attempt. The source review cannot assign an individual container to a specific query branch or prove which step retained the memory.

The first-amendment AB3868 retry also failed to complete ExcessSegments before the supervisor stopped it unsaved. Its [saved snapshot before the second stop](20261104-202652/memory-before-second-agent-stop.json), at 08:32:01 Sydney, records four Mashup private-memory values of 1,906.4, 22,792.8, 2,037.5 and 2,061.6 logged MB: 28,798.3 MB total. Excel private memory was 1,993.7 MB. Thus the earlier buffers did not resolve the observed pressure; this is not merely a theoretical risk.

The subsequent bounded read-only probes forced every scalar output column at each selected stage. Each receipt records closed-unsaved operation and unchanged workbook and canonical-source hashes. These are independent stage evaluations, not timings to add together or equivalent end-to-end refresh measurements.

| AB3868 stage probe | Rows | Scalar columns | Elapsed seconds | Saved evidence |
|---|---:|---:|---:|---|
| Enrichment immediately before physical grouping | 53,118 | 48 | 12.96 | [Enrichment receipt](20261104-202652/stage-Enrichment.json) |
| Physical shifts through BufferedMatchShifts | 17,639 | 52 | 12.24 | [Physical receipt](20261104-202652/stage-Physical.json) |
| Complete ShiftPrepare including chronology/checks | 17,639 | 79 | 41.04 | [Chronology receipt](20261104-202652/stage-Chronology.json) |
| Complete FortnightParts and counters | 17,809 | 25 | 65.26 | [Parts receipt](20261104-202652/stage-Parts.json) |
| Complete FortnightResults | 15,162 | 13 | 66.40 | [Fortnight receipt](20261104-202652/stage-Fortnight.json) |

The physical and chronology receipts contain 14,535 Unknown-identity and 3,104 Verified-identity shifts. Parts contain 14,673 Unknown-identity and 3,136 Verified-identity portions. Those counts describe current source outcomes; resolving the unknown identities was outside this performance amendment. The measured stages do not show a flat join explosion through physical shifts, and neither SelectedSource expansion nor chronology alone reproduces the full refresh failure. Repeated evaluation or retention when these branches are composed in ExcessSegments remains the leading hypothesis, not an allocation-trace conclusion.

## Grains, units and cardinality

Let R be prepared published input rows, N be canonical physical shifts after duplicate resolution, S be net-work segments, P be shift/fortnight parts, and F be employee/fortnight rows. The AB3868 probes measured R = 53,118, N = 17,639, P = 17,809 and F = 15,162; S and complete ExcessSegments output counts remain unmeasured here.

| Stage | Grain and units | Expected row behavior |
|---|---|---|
| `RosterPrepare`, lines 297–331 | One supplied published-duty candidate; dates/local datetimes, minutes for breaks, hours for net/gross quantities | Explicit Published=false and Non Attended=true records are excluded. Other uncertain publication/duty facts remain blockers downstream. All three configured exports contribute. |
| `WorkerPrepare`, lines 336–358 | One row per worker employee ID, with a count of distinct current facts | Identical worker facts collapse; conflicting facts produce one ambiguous lookup row, not several expanded matches. |
| `ShiftPrepare`, lines 363–529 | One physical shift identified by employer, worker, facility, original boundaries and source role | Physical grouping at 460–466 collapses identical copies. Conflicts remain Unknown. No display-name join is used. |
| `WorkSegments`, lines 534–557 | Net-work positions in hours; clock bounds only when meal/clock facts establish them | At most two segments per shift; no guessed meal proration. |
| `FortnightParts`, lines 563–618 | One shift/fortnight portion; counters and excess in hours | At most two positive portions for the encoded sub-24-hour shift span and configured 14-day period. Unknown anchor produces no period allocation. |
| `FortnightResults`, lines 623–660 | One employee/common roster-fortnight; hours and whole local calendar days | Grouped employee/period keys are unique. Invalid duty remains a completeness blocker. |
| `ExcessSegments`, lines 665–711 | One disjoint excess interval in net-work positions, with all applicable daily/fortnight reasons | Constant-sized local cut lists, including meal boundaries. Daily and fortnight overlap is not summed twice. |
| `ShiftResults`, lines 716 onward | One SourceShiftID with scalar summaries | Period/segment matches remain nested and are summarized; no expanded period/segment Cartesian product. |

The `PayrollFortnightStart/End` names are retained interfaces. Their anchor is now the earliest Settings permutation date; the names do not establish payroll alignment.

## Join audit

All source-level joins are keyed. None uses Name as employee identity. A nested one-to-many result is not itself a multiplied output: multiplication occurs only where that nested result is expanded.

| Join and lines | Keys / right-hand uniqueness | Cardinality assessment |
|---|---|---|
| Role, 404–405 | Normalized source role → role definitions | Right may contain duplicates, but the nested table is not expanded. Exactly one match is required to select Role; otherwise it is null. R rows remain R. |
| Facility, 407–409 | Normalized source location → facility definitions | Same scalar-selection behavior. Duplicate definitions cannot directly multiply physical rows. |
| Worker, 411–413 | WorkerEmployeeID → grouped worker lookup | WorkerPrepare guarantees one lookup row per ID; expanded facts do not multiply R. This does not independently prove employee-ID uniqueness outside the configured Whiddon source scope. |
| Allocation alias, 414–419 | EmployerKey + AllocationCode → grouped alias count | One aggregate row per alias key. Multiple worker IDs make identity Unknown rather than selecting an alias arbitrarily. |
| Match count, 467–473 | Encoded physical handoff key → grouped key count | Right is unique by construction. Both consumers now use BufferedMatchShifts. |
| Shift work, 570 | SourceShiftID → 0–2 work segments | Nested only; later portion construction uses that shift's small segment table. No full work-segment scan per shift remains. |
| Employee duty, 636–643 | EmployeeKey → all supplied duty for that employee | Nested only. Period filtering happens within the employee subset and retains uncertain rows by SourceDate. |
| Fortnight status, 671–675 and 720–722 | EmployeeKey + period start → grouped fortnight results | Right is one row per employee/period. Expanding status fields preserves P. Excess now buffers only the two keys and scalar status before joining. |
| Shift parts, 675–684; ShiftResults 722 | SourceShiftID → 0–2 period parts | Excess buffers seven scalar part columns, then intentionally expands period portions, with one whole-shift fallback when net positions cannot be split. ShiftResults keeps the same relation nested. |
| Excess work, 686 | SourceShiftID → 0–2 work segments | Nested only. Clock lookup scans at most that shift's segments. |
| Work timing / excess summary, 740–743 | SourceShiftID → timing aggregate / excess segments | Timing aggregate is unique; excess remains nested and is summed once. ShiftResults remains N. |

These cardinality protections are verified in the source; stage counts above supply additional bounded runtime evidence. Actual reference-list contents and every production join's before/after count were not audited here.

## Applied memory amendment

| Verified current change | Named steps and current lines | Intended effect and measured limitation |
|---|---|---|
| Buffer the three-column role and facility maps | RoleKeyRows/RoleKeys, 385–387; FacilityKeyRows/FacilityKeys, 392–393 | Reuse scalar lookup data while nested matches are counted and selected; no full SharePoint navigation buffer. |
| Buffer the already narrowed roster preparation | SourceRows, 395 | Coverage, aliases and main preparation at 396, 403 and 414 reuse this local buffer. |
| Remove nested lookup/context columns before physical grouping | RemovedNestedInputs, 457; PhysicalGroups/PhysicalColumns, 460–465 | Two nested lookup tables and seven obsolete helper columns no longer survive into SelectedSource records. |
| Materialize scalar physical match rows before their two consumers | BufferedMatchShifts, 469; MatchCounts/MatchJoin, 470–471 | Avoid replaying the same physical preparation for grouping and the left side of the match-count join. |
| Use the post-cleanup current table for chronology | EmployeeGroups/ChronologicalShifts, 507–508 | Moving cleanup did not accidentally bypass coverage or change the grain. |
| Narrow the ExcessSegments shift input to eleven scalar fields | ShiftColumns, 667–668 | Per-part records no longer carry unused provenance, chronology and other check fields. No additional shift buffer was added. |
| Buffer three-column fortnight status before the parts/status join | FNStatusColumns/BufferedFNStatus, 671–674 | Materialize the two join keys and status before reuse, without forcing thirteen-column FNResults output. |
| Buffer seven-column period parts before nested per-shift access | PartsColumns/BufferedParts/ShiftPartJoin, 675–677 | Repeated RowCount, null-count and position reads operate on scalar materialized parts rather than retaining/replaying the compound branch. |

The query count remains 33. There are no new sources, loaded cache tables, date filters, population reductions or business-rule changes. The [first-amendment formula-update receipt](20261104-202652/Core-update-20261104-212415-a92b4653-receipt.json) records only ShiftPrepare updated. The [current leaf-update receipt](20261104-202652/Core-update-20261104-214439-14ec5b86-receipt.json) records only ExcessSegments updated, all three loaded-output bindings preserved, save/close completed, and exact re-extraction matched. Both formula-only operations explicitly record no refresh. SelectedSource grouping, chronology, and both Table.Combine calls remain unchanged.

The eleven retained shift fields are SourceShiftID, EmployeeKey, QualifyingNetHours, DailyHoursLimit, EligibilityStatus, MealTreatmentStatus, MatchStatus, DutyStatus, CoverageStatus, OverlapStatus and DailyStatus. Exact clock bounds continue to come from the existing seven-column WorkSegments table. The output schema, join keys, null/Unknown behavior and interval trigger union are unchanged.

## Remaining source-level risks and runtime hypotheses

1. **ExcessSegments has several dependency paths to the same preparation**, at 668–686. FortnightParts and FortnightResults also reach ShiftPrepare/WorkSegments. The new leaf buffers cut two compound scalar branches before nested reuse. Buffers remain local to evaluations; they do not persist across independent query evaluations or loaded outputs. The number of physical reads and evaluations is not measured.
2. **Some upstream leaf inputs remain wide.** WorkSegments at 536 and FortnightParts at 565 still start from ShiftPrepare before projecting records. ExcessSegments now projects its shift input explicitly. Further projections remain unapplied possibilities; the completed standalone probes do not establish a need to change those upstream inputs.
3. **The period-filtered employee duty table is still streamed.** Day counting at 646–651 scans Duty for each of 14 dates; the local filter at 637–643 can be re-enumerated. This is per employee, not a whole-organisation scan per day. A local buffer of that small filtered table is a possible follow-up, not a measured hotspot.
4. **Buffers trade retained materialized values for fewer lazy contexts and repeated passes.** The current amendment buffers small-width scalar tables at the measured tens-of-thousands row scale, but the first amendment's stopped retry shows buffers alone do not guarantee improvement. The current change does not yet prove that private memory will fall. Actual stage counts are recorded above; reference-list contents remain outside this audit.
5. **Independent source/parameter retrieval remains possible.** Shared named imports and scalar parameter queries are not a persistent cache. No direct File/SharePoint call is placed inside the per-shift business loop, but source/provider evaluation behavior still requires runtime evidence.

Other existing bounded buffers remain the two-column path mapping, grouped scalar worker facts, three-column shift-period settings, seven-column work segments in two consuming evaluations, and eight-column duty facts. Chronology and counters use buffered lists and List.Generate; they do not repeatedly sum growing prefixes or copy a growing accumulator at every worked-hour increment. No whole workbook or SharePoint navigation was buffered.

## Verification and pending live result

The current 23F8 source passed a fresh [57 core native cases](../native-tests/20261104-214425-d7b0a44e/summary.json), with zero external data sources. The unchanged Shifts/consumer source previously passed [29 native cases](../shifts-native-tests/20261104-212345-d2d0c4cb/summary.json). These test business output values and interfaces; they do not validate production scale or live memory use.

The current third full-refresh attempt reached the first ExcessSegments refresh at 08:45:49 Sydney time in [Core-refresh.log](20261104-202652/Core-refresh.log); the saved status was RUNNING at review time. Await its completed output, errors and memory snapshots. Record elapsed time per output and private memory as well as working set. The stopped AB3868 retry did not resolve pressure; no production speed or memory improvement is claimed for the new 23F8 amendment before its live result.
