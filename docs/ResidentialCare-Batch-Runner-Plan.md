# ResidentialCare expanded Date-level batch runner

## Implementation and rollout status

The separate manual entry point is `CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd`.
The existing runner remains the production fallback until the expanded runner
passes its separately approved controlled refresh trial. This implementation does
not start a refresh or install a watcher.

**Run any independent, ready batches in parallel, with a default and hard maximum
of twelve active batches across the entire Date-level run.** This is not limited
to C2. Every batch refreshes its selected files sequentially, with one workbook
active per batch. A lower `-MaxParallelBatches` value is supported; one is serial.

The ceiling was raised from seven to twelve on 21 September 2026. The latest
completed production run, `20260921-154942-fd516fc7`, reached nine concurrent
batches, completed successfully in approximately three minutes and peaked at
61% total memory with approximately 38% background use. Twelve remains available
for future plans with enough simultaneously ready batches, but has not yet been
reached in a live refresh.

The executable catalogue and operational settings are located beside the new
runner. Project-relative locations are configured in `pq.project.json`.
The catalogue references the unchanged legacy Unit and organisation manifests
for existing workbook paths.

## Operator controls

Run these examples from the repository root, using PowerShell 7:

```powershell
# Primary action: configured default Units (currently BD, TE and JH-RY only).
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -RunAll

# No Excel: inspect the resolved plan, or validate file access and locks.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -ShowPlan
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -RunAll -ValidateSelectionOnly

# Explicit Unit-only run; IncludeOrg applies only to organisation consumer Units.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -RunAll -Units Unit1

# Explicit legacy scope, including organisation work.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -RunAll -Units Unit1,Unit2 -IncludeOrg

# Non-contiguous batches, or named/numbered capacity role batches.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Units Unit1,Unit2 -Batches D1,A3 -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Units Unit1 -Batches C2.1 -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Units Unit2 -Roles RN -RefreshSelected

# Add required producer files instead of using their saved outputs.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Units Unit1 -Batches A2 -IncludeDependencies -RefreshSelected

# Exact file identity; same filenames in different Units must be qualified.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Workbooks Unit1/DemandMaster -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Workbooks "UNITS/Unit1/2. Calculations/Effort.xlsx" -RefreshSelected

# Preserve existing global numbering, or explicitly request Unit-local numbering.
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -StartAtSequence 24 -EndAtSequence 26 -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -Units Unit2 -StartAtSequence 14 -EndAtSequence 16 -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -StartAtWorkbook "2. Calculations/Cost/Cost..xlsx" -RefreshSelected

./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -ShowStatus
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -ResumeRun <RunId> -RefreshSelected
./CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd -StopMode AfterBatch
```

Opening the command without arguments presents a menu; option 1 runs the
configured default Units. The current setting selects BD, TE and JH-RY and
excludes Unit1, Unit2 and organisation batches. Option 3 allows an explicit
Unit selection.
The menu leaves successful, blocked and failed results visible until the operator
presses a key to close the window. This includes validation and status results.
Option 0 (Exit) closes immediately. Run exclusion is released before waiting, and
the original exit code is preserved. Command-line calls with arguments never pause.
The separate live-approval gate has been removed at the user's request; Run options now proceed
through normal validation. Dependency and operational safety checks remain active.
Selecting option 2 displays the batch pick list **before** asking for batch IDs.
Each choice is formatted as `ID - Title [file1.xlsx, file2.xlsx]`, omitting the
title when none is available and preserving file execution order. Unit choices
are listed once; organisation choices appear separately. The list includes the
C2 all-enabled-roles choice and dynamically generated C2.1 through C2.N choices,
with role names and filenames. The same list-before-prompt rule applies when an
agent asks the user to choose batches in conversation. The list is generated
from current configuration, not copied from a potentially outdated diagram.
Organisation batch IDs use the letter `O`, as in `O1,O2,O5`. The runner also
accepts `01` through `05` with a leading zero as aliases for `O1` through `O5`
and prints the interpretation before resolving the plan. A batch selection
containing only organisation IDs runs those batches once and does not ask for
Units. A mixed Unit and organisation batch selection asks for Units; that answer
limits only the Unit jobs, not the organisation workbooks or their configured
consumer Units. Unknown IDs still stop before refresh.

Every Unit-selection prompt also first lists the currently discovered Unit IDs
in numeric order, followed by `All - all listed Units`. This applies to option 3
and the Unit questions after Unit batch or role selection. Accept comma-separated
IDs or `All`. State that blank means all Units for batch/role qualification;
blank cancels option 3. Agents must show the same available IDs before asking
the user to select Units in conversation. Unit IDs come from current folder
discovery.

`-RunAll` authorises execution. Other selections require `-RefreshSelected`;
without it they resolve and validate only. `-ShowPlan` is read-only catalogue
preview. `-ValidateSelectionOnly` validates paths, access and file users but
never opens Excel, even alongside an execution switch.

Selection modes (all, batches, exact files, start-at-file, range) are mutually
exclusive. Roles may select or qualify C2 batches. Comma-separated values work
through the CMD launcher. Empty interactive selections fail closed.

Legacy sequence IDs keep their manifest meanings; new input workbooks use
batch/file selection rather than renumbering existing IDs. Unit-local ranges
require exactly one Unit. Disjoint selections never expand into a min/max range.
Dependency expansion is explicit and may add producer work outside the initial
Unit selection; the complete resolved plan is printed before execution.

## Catalogue and dependencies

Unit batches retain D1, A1, S1, C1, D2, A2, A3, C2 and U4:

- D1 includes Demand-MasterRoster Manual Read.xlsx before 2-DemandExtract.xlsx.
- C1 starts with Worker Reconciliation.xlsx, then Capacity-ShiftAvailability.xlsx
  and StaffListMaster.xlsx.
- C2.1 through C2.N expand from enabled roles in the saved client profile.
  Each contains the three role workbooks in the profile's order. Stable internal
  identity uses the Unit and role-folder name; display numbering is frozen for
  a run. Test folders and unlisted roles are not auto-included.
- Organisation batches O1 through O5 run once, not once per selected Unit.

The current two-Unit, three-role profile resolves to **60 files in 27 batches**:
26 files per Unit and eight organisation files.

Dependencies are file-level edges, not blanket stage barriers. The draft
execution catalogue includes Allocation input -> Settings Data -> demand
transformation, and Intervals.xlsx -> A2's Shifts.xlsx. Shifts need not wait for
later D2 files. An early organisation file can run as soon as its actual Unit
inputs are ready; O3 and O4 are independent once their own inputs are ready.

Selecting Units does not change organisation workbooks' internal filters.
The batch settings declare their configured consumer Units; those dependencies
are included even when only a smaller Unit selection is refreshed.

### Readiness and concurrency

Maintain one ready queue across branches and Units. Admit a batch whenever its
next selected file has satisfied prerequisites, accessible inputs and a free
slot. A waiting batch consumes no slot. Refill slots as work finishes. Ready ties
use numeric Unit order, catalogue batch order and saved role-profile order.

Multiple readers can share stable inputs. No producer may overwrite an active
reader's input. An approved additional-input inventory also participates in
dependency matching and read/write exclusion. Jobs with unverified write side
effects can be declared exclusive and run alone.

Unselected upstream files may be reused after access/lock checks, explicitly
labelled **existing input - not refreshed this run**. Their presence is not proof
of freshness. A selected prerequisite that fails cannot fall back to an old saved
output. Unavailable files block their jobs and descendants; independent work
continues. Validation-only reports unavailable files with a nonzero exit.

## Explicit run authorisation and retained protections

The separate approval gate has been removed at the user's request. A Run menu
selection, `-RunAll`, or a scoped `-RefreshSelected` can start a refresh after the
normal file and dependency checks. No approval flag, evidence field, approval
fingerprint or organisation-confirmation flag is required. Neither preview nor
validation-only executes a refresh, even when combined with a run switch.

The historical `runner/ResidentialCare-BatchApproval.psd1` filename and
`pq.project.json` key are retained for compatibility. That file now contains
operational settings only:

- `OrganisationUnits`: the configured Units consumed by organisation workbooks.
- `AdditionalInputs`: extra read paths by stable job ID. Repository-relative paths
  and environment roots are supported, for example
  `@{ Environment = 'PUBLIC'; Path = 'Public Scripts/CentriSyncPaths.xlsx' }`.
- `ExclusiveJobs`: jobs that must run alone because of additional write effects.

Removing the approval gate does not certify that workbook sources were inspected.
Keep the dependency catalogue and declared inputs aligned with the actual workbook
workflow; workbook/source changes still follow the project's source-of-truth and
opt-in synchronization rules. No M source or workbook logic is changed here.

Read/write exclusion, selected-producer failure blocking, file-user checks,
timeouts, verified saves and the twelve-batch maximum remain enforced. Configuration,
profile, manifest and relevant engine changes still invalidate the identity used
for resume; this fingerprint is no longer a live-approval signature.

The existing coloured staging Mermaid and dynamic C2 layout remain untouched.
They explain batches, not the runtime dependency authority.

## Progress, stopping and recovery

Run records live below `RunLogs/BatchRefresh/<RunId>/`: immutable `plan.json`,
atomic `state.json`, and separate per-job/per-attempt logs, receipts and stop
files. Status includes batch counts and pending/running/completed/failed/blocked/
stopped file states. Technical success is separate from business reconciliation,
which is reported as not evaluated.

The run window now prints each file starting and finishing, and displays the
path to a readable `coordinator.log` in that run's directory. The same events
are appended there. Use menu option 8 or `-ShowStatus` for a current summary;
per-job `refresh.log` files contain detailed refresh/readiness heartbeats.
Validation-only explicitly says Excel was not opened; execution instead announces
that the selected refresh is starting.

At the end of each execution attempt, the window displays `Run duration` in
hours:minutes:seconds, including setup, refresh, reporting and cleanup after
run exclusion is acquired. It excludes selection/validation and the final key
wait. Resumed runs time the current attempt only. Failed and stopped attempts
also show the duration; preview, validation and status requests do not.

Coordinator and worker JSON publication retries transient Windows access/sharing
denials for up to 30 seconds per write. It preserves atomic replacement: it never
truncates the published file or changes permissions to defeat a reader. The
coordinator writes on state changes and at a 15-second heartbeat, rather than
every 100-ms polling cycle. A running intent is still saved before every worker
launch. A stop arriving while publication waits is rechecked before dispatch.

An application holding a status file indefinitely can still prevent publication.
Prefer the runner's status command or the readable log over keeping raw JSON
open in a viewer that prevents replacement. Temporary denial no longer immediately
stops the run; persistent denial fails safely and names the exact affected path.
No workbook lock, timeout or owned-process protection has been removed.

Coordinator exceptions are recorded in `coordinator-error-<id>.json`, including
the original message and code location, and in the live log. During shutdown,
confirmed save/close completions retain their completion metadata; interrupted
unsaved files remain stopped and uncertain saves require inspection. A coordinator
failure returns exit 1 rather than being disguised as an operator stop.
If final state publication also fails, `state-recovery-<id>.json` captures the
terminal state separately, and prepared JSON replacements are retained as `.tmp`
evidence. `state.json` may then be stale. Status warns and resume is blocked;
inspect the records and original workbooks before creating a targeted new run.
Do not replace the ledger with a recovery file or automatically retry an uncertain
save. A failure to copy a worker's console log does not discard its valid receipt.

The new runner uses the existing supervised Excel refresh worker for every
file, including organisation files. Hidden mode and a 30-minute per-file timeout
are defaults. Optional visibility and timeout controls remain available.
There is no blanket termination of user Excel processes.

A shared Date-level gate excludes legacy and parallel-test refreshes. Nested
legacy stages may share their existing run, while the new coordinator is unique.
Batch worker leases retain exclusion if the coordinator dies; workers request a
safe supervised stop when their registered parent disappears.

- **AfterCurrent:** finish currently running files; start nothing further.
- **AfterBatch:** finish admitted unfinished batches; admit no new batch. If a
  dependency would require starting another batch, stop and report the unmet
  dependency instead of overriding the stop.
- **Now:** ask each active supervisor to stop and clean up its owned processes.

Resume keeps the frozen selection and role mapping. It checks recorded input/
output size and modification-time metadata, invalidating affected downstream
completion when inputs change. Metadata is a change indicator, not a business
validation. Interrupted/running or uncertain-save jobs require inspection;
verify the original workbook and create a new targeted run rather than editing
the ledger to force success. No automatic refresh retries occur.

Exit 0 means all selected files completed; 1 means blocked/failed work; 3 means
operator stop; 124 preserves a workbook timeout when the run was not stopped.

## Verification and rollout

`pwsh -NoProfile -File scripts/test-batch-refresh-runner.ps1` exercises selection,
dynamic identities, twelve-batch admission, cross-branch/Unit scheduling, serial
batch files, file-level release, waiting slots, read/write conflicts, partial
failure, stop/resume and exclusion with synthetic non-workbook files.

Also run the existing Excel-safety and parallel-role mock suites and the legacy
smoke checks. Validate the complete new plan without Excel. Resolve reported
locks/missing inputs before the approved controlled refresh trial. Keep the
legacy entry point until rollout is accepted.

`pwsh -NoProfile -File scripts/test-batch-runner-recovery.ps1` reproduces Windows
access-denied replacement failures using disposable JSON files and real file
handles. It checks transient retry, bounded permanent failure, preservation of
the published JSON and recovery evidence, original-error reporting, accurate
completion during shutdown, stop during publication, heartbeat throttling and
uncertain-save protection. All workers and output files are synthetic; no Excel
workbook is opened.

`pwsh -NoProfile -File scripts/test-batch-runner-launcher.ps1` checks the actual
CMD launcher in a disposable configuration with mock dispatch and no Excel workers.
It verifies that C1/C2 Unit1 reach mock dispatch without approval metadata,
successful and failed menu requests leave results visible until a key press,
run exclusion is released before waiting, scripted calls never pause, explicit
Exit closes normally, and validation-only never dispatches.

Automatic watching, scheduled refresh, workbook business-logic edits and
M-to-Excel synchronisation are outside this release.

### Implementation verification — 2026-09-16

- Expanded runner suite: 113 assertions passed, including a real PowerShell
  adapter process using a synthetic-data worker; no Excel started.
- Existing Excel-safety suite and seven-role mock harness passed.
- Expanded CMD preview (two Units, A1 and A3): exactly eight files/four batches;
  default maximum seven. Expanded status and Unit1 A1 validation passed.
- Legacy status and global sequence 1 validation passed.
- Full expanded validation resolved 60 files/27 batches, but reported an Excel
  lock on Unit1's DemandIntervals.xlsx and its use by Demand. No lock was removed.
- The pre-existing RN CMD smoke check returned exit 10 because its referenced
  bin/runner/Invoke-RNRefresh.ps1 is missing. That unrelated shortcut was not repaired.
- The initial implementation shipped with live approval disabled. That separate
  gate has since been removed at the user's request; operational safety checks
  remain. No real refresh, workbook save, M sync or watcher installation was
  performed while making this change.

### Coordinator interruption repair — 2026-09-16

- Reproduced the reported `Move` access denial (Windows code 5) with a held
  status-file handle and with a temporarily read-only synthetic JSON target.
  The old writer did not retry `UnauthorizedAccessException`; its IO retry
  budget was only 200 ms. The replacement successfully waits for either
  temporary denial to clear.
- Fixed shutdown reconciliation that previously marked every active file
  stopped, even when a worker had already returned a successful save/close.
- Latest historical Unit1 run `20260916-224740-a2435452` remains unchanged:
  11 completed, ShiftAverage stopped before saving, 14 pending. Its original
  coordinator exception was not persisted, so the exact trigger and holding
  application for that run cannot be established from the surviving records.
  The earlier full-run console explicitly reported the JSON `Move` denial.
- Batch regression: 133 assertions passed. Recovery regression: 55 checks
  passed. Menu/launcher regression: 68 checks passed. Existing supervised-worker
  safety tests and all 21 files in the seven-lane mock test passed.
- Complete expanded plan validation passed: 60 files, 27 batches, using the
  saved role profile. Legacy status and sequence 1 validation passed. The
  pre-existing RN shortcut still references a missing script and was not changed.
- No live Excel refresh, workbook content inspection, workbook save, M sync,
  watcher or Mermaid change was performed for this repair. A fresh controlled
  refresh remains the live verification step. The engine fingerprint changed:
  older run records must not be force-resumed under the revised implementation.

### Git coexistence investigation and strategy — 2026-09-17

The completed run `20260917-161919-19c329e1` saved five files, failed two and
blocked 53 descendants. Unit2 AllocationInput refreshed but the pre-save
external-user check detected Git for Windows (PID 57664) and refused to call
Excel Save. This was a runner policy rejection, not evidence of a failed Excel
save or proof that Git held an incompatible write lock. The Git process had
exited before investigation; its originating application was not identified.

Repository history establishes a behaviour change: the original linear worker
in `47e11a2` (8 September) called Excel Save directly. `464b57f` (13 September)
added the external-file-user guard. The hardened linear workers and expanded
batch worker share that guard now. Parallelism may increase overlap with readers,
but it is not necessary for this failure. Metadata-only index inspection found
95 client Excel files tracked, including the failed Unit2 allocation workbook.
The separate Unit1 DemandMaster failure was an empty workbook identity during
opening, not a demonstrated Git conflict.

Initial hardening added waits for external users before opening/saving (up to
60 seconds per worker check, within the overall supervisor timeout). Pre-dispatch
batch checks allow 15 seconds; validation-only remains an immediate check. That
first change did not exempt Git. The approved refinement below now permits only
verified read-only Git commands. No external process is killed and no tracking
configuration is changed. Stop-now remains responsive. Changed target metadata
during the worker's access wait blocks saving instead of overwriting an intervening edit.
The worker waits up to 15 seconds for Excel's exact workbook identity; a wrong
path or read-only workbook fails closed. Persistent missing identity remains a
failure requiring investigation, not permission to choose another workbook.
These guards are not a complete repository-isolation strategy.

### Approved Git-aware refinement — implemented 2026-09-17

The shared guard now classifies each detected external user. A Git exception
requires the same PID and creation time as the Windows file-user identity, an
exact installed executable path resolved from `PATH` (or its standard Windows
backend), and an explicitly allowed command/option combination. A process name
such as `git.exe` or display label `Git for Windows` alone is never sufficient.
Unverifiable processes, unrecognised installations and unknown options fail
closed. The installation trust is based on configured executable paths, not
signature attestation or a sandbox against malicious repository configuration.

The allowlist in `CLIENT/DATExx-Whiddon/runner/src/RefreshGitAccess.ps1` supports narrow forms of
`status`, `ls-files`, `cat-file`, `hash-object` and the `diff` family. Diff commands
must include both `--no-ext-diff` and `--no-textconv`; `hash-object` must include
`--no-filters` and cannot write objects. Only limited harmless global options are
accepted. Aliases, arbitrary configuration overrides, write commands such as
`add`, `checkout`, `restore`, `reset`, `merge`, `pull` or `commit`, and unknown
options are not exempted. Read-only here refers to workbook/working-tree content:
Git status may still refresh its own index metadata.

Permitted readers bypass only the runner's external-user policy rejection.
Actual OS sharing/access failures, Excel lock files, internal runner exclusion,
dependency read/write exclusion, exact workbook identity, timeout and verified
save/close protections remain in force. This does not detect or prevent every
external write operation: do not switch branches, restore or otherwise replace
workbooks during a refresh. No arbitrary Git process is suspended or terminated.

Each observation is logged immediately as `FILE USER` with target, PID, creation
time, executable, redacted command, parent application when available, and the
allow/block reason. Worker evidence is in each attempt's `refresh.log`. A policy
rejection says explicitly that the check did not attempt Save. Actual Save
exceptions retain their phase and original message in the worker state and batch
result, rather than being described merely as an unexplained exit code.

The coordinator freezes declared input size/modification-time stamps at dispatch
and sends `input-snapshot.json` to the exact supervised workbook job. The worker
captures the target stamp before opening, checks it after opening and before
saving, and checks the declared inputs before opening, before saving and after
saving. A changed/missing input is not accepted as a successful refresh. If an
input changes during Save, the uncertain outcome requires inspection; no
automatic retry occurs. Target checks apply to both worker entry points; legacy
calls without an input snapshot do not acquire a new implicit input inventory.
Metadata checks cannot detect changes that preserve both size and timestamp,
and are not atomic isolation against unrelated writers.

The engine fingerprint includes the classifier. Start a fresh run after this
change; do not force-resume older records with a different engine fingerprint.

Verification of the refinement, without Excel:

- `scripts/test-git-aware-refresh.ps1 -LiveProcessChecks`: 58 checks passed,
  including real installed Git command/process identification and 18 synthetic
  file jobs (nine serial, nine with maximum seven active batches) alongside real
  Git reads. File-user sightings are injected for deterministic scheduler tests;
  this is not proof of a live Excel save with a Restart Manager-reported reader.
- Denied commands, unknown/missing identities, PID reuse, real OS sharing denial,
  stop during waits, and changed/missing target/input metadata were exercised.
- Batch suite: 149 assertions; recovery: 55; launcher: 68; file-access/identity:
  26. Existing Excel-safety cases and all 21 files of the seven-role mock passed.
- Full expanded validation passed: 60 files, 27 batches, saved role profile.
  Legacy status and sequence 1 validation passed. The pre-existing RN shortcut
  still references a missing script and was not repaired by this change.
- No real workbook opened/refreshed/saved, no M sync, Git tracking change or
  Mermaid modification. A controlled live refresh remains untested. The separate
  DemandMaster opening/identity failure is not established as resolved.

### Optional operating strategy (not enabled automatically)

1. **Near term: a quiet Git window when needed.** Complete deliberate repository
   operations before refresh. Verified allowlisted readers can coexist; pause
   other background repository scans in the tools that launch them where
   supported, and avoid staging/committing, switching branches,
   checkout/restore, pulls or merges in the working tree while a refresh runs.
   Keep code-only inspection separate from live workbook scans. Re-enable checks
   after all workers close. Git is invoked by tools; it is not one background
   service the runner can universally suspend. A runner lease or an index lock
   cannot make unrelated Git readers cooperate. Do not suspend/kill arbitrary Git
   processes or create a fake index lock.
2. **Durable isolation: an explicit execution root outside the Git worktree.**
   Keep versioned workbook sources/templates, M sources, scripts and configuration
   in the repository. Prepare an approved complete Date-level dependency tree
   in a separate runtime root, preserving Unit/role identity and relative paths.
   Validate that workbook imports and path resolution target that runtime tree,
   not the repository originals, before enabling refresh there. This needs a
   deliberate runner/path design and approval; do not silently run temporary or
   same-stem copies. Continue to allow up to twelve independent batches.
3. **Controlled publication.** After refresh/save/close verification, publish
   only the explicitly selected successful output set, with recovery records,
   input provenance and checks against changes to destination originals. Report
   partial runs accurately. Repository publication still needs an access-safe
   window; a publication lock must not discard the successfully refreshed runtime
   output. Staging/committing workbook outputs is a separate deliberate action.
4. **Verification.** Compare serial and parallel runs with deliberate background
   Git reads in a disposable, approved test environment. Test actual competing
   writers, stop during waits, changed inputs, failed publication and incomplete
   saves. Require a separately approved live trial before claiming the original
   workbook failures resolved.

Do not use `assume-unchanged`/`skip-worktree` to hide refresh changes. Adding a
tracked workbook to `.gitignore` does not stop Git tracking it. No file relocation,
untracking, repository-wide Git setting, automatic checkpoint or M-to-workbook
synchronization is authorised merely by this proposal.

References: [Git background status refresh](https://git-scm.com/docs/git-status#_background_refresh),
[Git ignore scope](https://git-scm.com/docs/gitignore),
[Windows resource-user enumeration](https://learn.microsoft.com/en-us/windows/win32/api/restartmanager/nf-restartmanager-rmgetlist).

Synthetic verification: `scripts/test-runner-file-access.ps1` passes 26 checks
covering transient/persistent file use, cancellation, unavailable inspection,
target changes during a wait, delayed/missing/wrong identity and read-only opens.
No Excel workbook was opened or refreshed during these tests.
