# ResidentialCare runner agent instructions

Use these instructions when the user asks to validate, run, monitor, resume, or
stop the refresh workflow under `CLIENT/DATExx-Whiddon`.

The executable sequence is defined by the runner `.psd1` manifests. The maps in
`docs/ResidentialCare-Runner-Sequences.md` explain those manifests but do not
replace them.

## Expanded batch runner (separate rollout)

For requests explicitly naming the expanded/new/batch runner, use
`CLIENT/DATExx-Whiddon/Run-BatchRefreshRunner.cmd` and follow
`docs/ResidentialCare-Batch-Runner-Plan.md`. The legacy entry-point mapping below
remains the fallback until rollout is accepted.

The expanded runner's primary action is `-RunAll`. Without a Unit selector, it
uses `RunAllUnits` and `RunAllIncludeOrg` from the batch settings file. The current
default runs BD, TE and JH-RY only; Unit1, Unit2 and organisation jobs require an
explicit selection. The runner executes **up to twelve ready batches concurrently**
across the selected branches; files within each batch remain sequential.
`-MaxParallelBatches` accepts 1 through 12, default 12.
Batch IDs and the saved role profile are the manual-control interface.

Whenever an option leads to asking the user to pick batches, first display all
available batch choices. This applies both to the runner menu and to an agent
asking for a batch selection in conversation. Each choice must show its batch
ID, title when available, and its files in square brackets, comma-separated and
in execution order. For example:

```text
D1 - Demand transformation [Demand-MasterRoster Manual Read.xlsx, 2-DemandExtract.xlsx]
A2 - Role-allocation basis [Shifts.xlsx, AllocationByShiftAverage.xlsx]
```

Generate the list from the current catalogue, manifests and enabled role profile;
do not ask the user to remember IDs or refer them to a separate diagram. List
shared Unit batch choices once, with organisation batches in a separate section.
Show C2 as the all-enabled-roles choice and each actual C2.1 through C2.N choice
with its role name and three filenames. Then ask for one or more comma-separated
batch IDs. Listing choices must not open Excel or authorise a refresh.
Organisation batch IDs start with the letter `O`. The runner accepts `01` through
`05` as aliases for `O1` through `O5` and prints that interpretation. After an
organisation-only batch pick, do not ask for Units; those batches run once with
their configured consumer Units. For mixed Unit and organisation picks, ask for
Units to qualify only the Unit jobs. Unknown IDs remain invalid.

Whenever asking for Units, first list the exact available Unit IDs from current
Unit-folder discovery, in numeric order (for example, Unit1, Unit2, Unit10).
This applies to all Unit prompts, including those following batch or role picks,
and to an agent asking in conversation. Display `All - all listed Units`, then
ask for comma-separated Unit IDs or `All`. Explain the blank-input behaviour:
blank means all Units when qualifying batches/roles, but cancels the dedicated
Run selected Units option. Never require the user to remember Unit IDs or invent
friendly names. Listing Units does not open Excel or alter the selection scope.

Use `-ShowPlan` for a catalogue preview and `-ValidateSelectionOnly` for file-access
validation. Neither opens Excel. Other selections execute only with
`-RefreshSelected`. Unit-scoped all runs exclude organisation work unless
`-IncludeOrg` is explicit. Organisation filters inside workbooks are not changed.
Legacy global/local range numbering is preserved.

The user has removed the separate live-approval gate. An explicit Run menu
selection, `-RunAll`, or a scoped `-RefreshSelected` now proceeds through normal
validation to refresh; do not require an Approved flag, approval evidence,
configuration sign-off or OrganisationScopeConfirmed flag. Do not invent review
evidence. The historical `ResidentialCare-BatchApproval.psd1` filename now holds
only operational settings: default run scope, organisation consumer Units,
additional inputs and exclusive jobs. The configuration fingerprint still
protects resume consistency.

Dependency readiness, file access/lock checks, conflicting-run exclusion,
twelve-batch maximum, timeout and verified-save protections remain enforced.
Removing the gate does not start a refresh by itself, authorise M synchronization
or permit automatic scheduled refresh. Preview and validation never start Excel.

The no-argument CMD menu keeps its window open after successful, blocked or failed
requests until a key is pressed. Successful validation and status results also
stay visible. Option 0 (Exit) closes immediately. Calls with command-line arguments
never pause and preserve their exit code. Run exclusion is released before waiting
for a key. Menu option 7 still allows validation without opening Excel.

For the expanded runner, inaccessible files block their jobs and descendants,
while unrelated ready work continues; a nonzero final result reports incomplete
work. Invalid selectors and graph/configuration errors block execution globally.
Do not bypass a failed selected producer with its old
saved output. Unselected saved inputs must be reported as not refreshed this run.

Use the same entry point with `-ShowStatus`, `-ResumeRun <RunId> -RefreshSelected`,
or `-StopMode AfterCurrent|AfterBatch|Now`. Status identifies interrupted owners;
do not force-resume an uncertain save. Use its separate `RunLogs/BatchRefresh`
records, not the legacy current-status file, when reporting the new runner.

New runs print their readable `coordinator.log` path and each file's start/result.
Use that log, `-ShowStatus`, and the per-job `refresh.log` for monitoring. JSON
publication retries transient Windows access/sharing denials for up to 30 seconds;
state is written on changes and a 15-second heartbeat. Do not hold raw JSON open
in a viewer that prevents replacement, bypass workbook locks, or treat the
validation-only message as evidence that a subsequently requested run never started.

Inspect `coordinator-error-*.json` for the original coordinator exception.
Coordinator failures return exit 1; they are not operator stops. Shutdown must
preserve workers with confirmed save/close receipts as completed, including
their output stamp and finish/exit records. Uncertain saves still need inspection.
If a `state-recovery-*.json` file exists, final state publication failed and
`state.json` may be stale. Resume is blocked: inspect the recovery evidence and
original workbooks, then obtain an explicit targeted-run instruction. Never edit
the ledger to force success or automatically adopt a prepared `.tmp` replacement.
Old records lacking coordinator diagnostics cannot establish which application
caused a status-file denial; do not attribute one to the viewer or Git without
evidence. Engine changes invalidate older run fingerprints.

External-file-user detection is not proof that Excel attempted or failed a save.
The hardened worker waits up to 60 seconds for external users before opening and
saving, with stop-now and the overall deadline enforced. A narrow exception allows
verified installed Git processes running explicitly allowlisted read-only
working-tree commands. PID and creation time, executable path and command options
must all be verified. Git write commands, unknown commands/options, other external
applications and unavailable process information still wait or block. This does
not bypass OS sharing/access denials, Excel lock files or verified-save checks.
Never kill or suspend another application's process to clear a check.

Each observed external user produces a `FILE USER` diagnostic with PID, creation
time, executable, redacted command, parent application when available, target and
classification. Worker diagnostics are in that attempt's `refresh.log`. A policy
rejection explicitly says no save was attempted by that check; an actual Excel
save error is reported separately. The original worker error also appears in
batch completion records.

Workers capture target size/modification time before opening and recheck after
opening and immediately before saving. Expanded batches also freeze their
declared input metadata at dispatch and check it before opening, before saving
and after saving. Changed or missing files fail the job; an input change found
after Save makes the outcome require inspection, not automatic retry. Legacy
workers have the target check; input checks require an explicit input snapshot.
These are metadata checks, not workbook-content comparison or atomic protection
against arbitrary external writers. Keep workbook-changing Git operations out of
an active run.

Batch pre-dispatch allows 15 seconds; validation-only remains read-only and
immediate. Exact workbook identity may settle for up to 15 seconds, but a wrong
path, read-only workbook or permanently missing identity still blocks execution.
The allowlist and optional longer-term Git/execution separation are documented in
the batch runner plan. No Git suspension or external runtime root is configured
automatically. Do not untrack workbooks, move them, alter repository-wide settings
or substitute runtime copies implicitly.

After changes, run `scripts/test-batch-refresh-runner.ps1` and
`scripts/test-batch-runner-recovery.ps1` without Excel, the
existing safety/mock suites, then the legacy smoke checks and new validation.
For file-access or workbook-identity guards, also run
`scripts/test-runner-file-access.ps1` using synthetic data only.
For Git classification or refresh-wide metadata checks, run
`scripts/test-git-aware-refresh.ps1`. Its optional `-LiveProcessChecks` exercises
real Windows Git process identification and real Git reads of disposable data
alongside serial/twelve-batch simulated runs. File-user sightings are injected in
that test; it is not a live Excel save or end-to-end Restart Manager trial.
For menu/launcher changes, also run `scripts/test-batch-runner-launcher.ps1`;
it verifies explicit C1/C2 Unit1 dispatch, validation-only behaviour, successful and
failed menu hold-open handling, and run-exclusion release before waiting, in a
disposable configuration with mock dispatch and Excel disabled.
Report which checks passed and any access/approval blockers. Keep approved
Mermaid styling and role expansion unchanged.

## Natural-language command contract

Interpret the user's concise requests as follows:

| User request | Scope and action |
|---|---|
| `status` | Show the top-level current status and latest log. Do not refresh. |
| `validate all` | Validate the complete global sequence. Do not refresh. |
| `run all` | Validate, then refresh the complete global sequence. |
| `run units` | Validate, then refresh all discovered `UnitN` folders. |
| `run Unit1` | Validate, then refresh Unit1's complete 24-workbook sequence. |
| `run org` | Validate, then refresh the eight organisation workbooks. |
| `run RN` | Validate, then refresh the three Unit1 RN workbooks. |
| `run roles` | Validate, then refresh the nine Unit1 role workbooks. |
| `run calculations` | Validate, then refresh the 22-entry Unit1 calculation sequence. |
| `run 12 to 18` | Validate, then refresh global sequence 12 through 18 inclusive. |
| `run Unit1 12 to 18` | Validate, then refresh Unit1 sequence 12 through 18 inclusive. |
| `run calculations 12 to 18` | Validate, then refresh calculation-stage sequence 12 through 18 inclusive. |
| `run from <workbook>` | Use the top-level global sequence and begin at the named workbook. |
| `check roles` | Compare current role folders with the saved client role profile. Do not refresh or update files. |
| `stop after current` | Write an `AfterCurrent` stop request. |
| `stop now` | Write a `Now` stop request. |

Unqualified sequence numbers always mean the top-level global sequence. Use
the explicitly named stage's local numbering when the user names `Unit1`,
`RN`, `roles`, or `calculations`.

If a workbook name is ambiguous across folders, stop and ask for its relative
path. Do not guess.

## Saved client roles and optional checking

The persistent role list for this client is:

`CLIENT/DATExx-Whiddon/runner/ResidentialCare-ClientRunProfile.psd1`

The order of enabled entries in `Roles` is the approved role execution order.
Normal runs read this file and do not scan calculation folders.

Before any refresh request, ask one short question unless the user has already
answered it in the request:

> Use the saved client roles, or check the role folders first?

Interpret `use saved roles`, `no role check`, or an equivalent answer as
permission to proceed using the persistent list without scanning. Interpret
`check roles` as a read-only comparison before validation. The role-check
choice does not itself authorise a refresh.

For a role check:

1. Discover each `UnitN` folder under `CLIENT/DATExx-Whiddon/UNITS`.
2. Inspect its direct subfolders under `2. Calculations`.
3. Treat a subfolder containing any filename from `RoleWorkbookOrder` as a
   candidate role folder.
4. Compare candidate folders with enabled and disabled entries in the saved
   profile, case-insensitively.
5. Report unlisted folders, configured folders that are absent, disabled
   folders that are present, and role folders missing any required workbook.
6. Compare the enabled profile roles and their order with every applicable
   runner sequence manifest.
7. Do not change the profile, manifests, folder names, or workbook names. If a
   difference is found, stop and ask for explicit approval of the proposed
   profile and manifest update.

Use a role check when onboarding a new client or unit, changing role names,
enabling or disabling roles, or changing which role workbooks participate in
the run. The user can include the decision in a concise command, for example:
`run all, use saved roles` or `check roles, then run Unit1`.

## Execution rules

1. Treat `run` as explicit permission to refresh only the named scope. It also
   authorises the required validation immediately before that refresh.
2. Treat `validate`, `check`, and `preview` as validation-only instructions.
   They never authorise opening or refreshing Excel.
3. Resolve whether to use the saved roles or perform a role check, following
   the saved-client-role rules above.
4. Resolve the request to one entry point and an exact inclusive selection.
   State that interpretation before execution when the request could reasonably
   refer to more than one numbering scope.
5. Confirm `pwsh` is available, the client role profile and manifest load,
   enabled roles agree with the applicable sequence manifest, selected paths exist, no
   selected workbook has an Excel `~$` lock file, and no conflicting runner is
   active.
6. Run the selected scope with `-ValidateSelectionOnly` first. If validation
   fails, stop. Do not skip, rename, copy, or substitute a workbook to bypass
   validation.
7. After successful validation, continue directly to refresh only when the
   user's current request used `run` or otherwise explicitly requested refresh.
8. Run hidden by default. Use `-VisibleOverride true` only when the user asks
   for a visible run.
9. Keep the default 30-minute per-workbook timeout unless the user specifies a
   different timeout.
10. Do not synchronize `.m` code into a workbook as part of a refresh request.
   M-code synchronization remains separately opt-in.
11. Monitor the active process and status/log files. Do not start an overlapping
    runner.

## Entry-point mapping

- Global/all/org/range/from-workbook:
  `CLIENT/DATExx-Whiddon/Run-RefreshRunner.cmd`
- All discovered units:
  `CLIENT/DATExx-Whiddon/UNITS/Run-AllUnitsRefresh.cmd`
- Unit1 complete:
  `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-UnitRefresh.Stage4-Unit1Only.cmd`
- Unit1 calculations:
  `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-CalculationRefresh.Stage3-UnitCalculations.cmd`
- Unit1 roles:
  `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-RoleRefresh.Stage2-AllRoles.cmd`
- Unit1 RN:
  `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-RNRefresh.Stage1-RNOnly.cmd`

## Completion report

Always report:

- interpreted scope and inclusive sequence range;
- validation result;
- whether Excel was opened and whether it was hidden or visible;
- each completed workbook or the count completed;
- the exact failed or stopped workbook, when applicable;
- exit code and latest log/status path;
- whether any workbook was refreshed, saved, or left open.

## Three brief smoke tests

These tests verify agent control without opening or refreshing Excel. Run them
in order after runner installation or changes.

### Test 1 — status path

```cmd
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -ShowStatus
```

Expected: the runner starts under PowerShell 7, prints current status or reports
that no status/log exists, and exits without Excel.

### Test 2 — narrow manifest validation

```cmd
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-RNRefresh.Stage1-RNOnly.cmd -ValidateSelectionOnly -RefreshAll
```

Expected: all three RN workbooks validate and Excel is not started.

### Test 3 — single global sequence validation

```cmd
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -ValidateSelectionOnly -StartAtSequence 1 -EndAtSequence 1
```

Expected: global sequence 001 validates and Excel is not started.

After these three smoke tests pass, resolve all known missing manifest paths and
run `validate all` before any complete production refresh.
