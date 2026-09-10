# ResidentialCare runner agent instructions

Use these instructions when the user asks to validate, run, monitor, resume, or
stop the refresh workflow under `CLIENT/DATExx-Whiddon`.

The executable sequence is defined by the runner `.psd1` manifests. The maps in
`docs/ResidentialCare-Runner-Sequences.md` explain those manifests but do not
replace them.

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
