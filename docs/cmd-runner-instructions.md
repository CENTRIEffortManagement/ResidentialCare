# ResidentialCare CMD runner instructions

These instructions apply to the local Whiddon workflow under
`CLIENT/DATExx-Whiddon`.

The `.cmd` files are thin Windows wrappers. They call the matching PowerShell
runner with PowerShell 7 (`pwsh`), pass through the command-line arguments, and
return the runner's exit code. Workbook order comes from the matching `.psd1`
manifest; it is not defined in the `.cmd` file.

## Current readiness

Do not start a refresh until all preflight items pass.

- The runner package and six `.cmd` entry points are present locally.
- The PowerShell scripts and `.psd1` manifests parse successfully.
- PowerShell 7 (`pwsh`) is not currently available on this PC's `PATH`.
- The Unit 1 manifest currently references two workbook paths that do not
  exist locally under those exact names:
  - `UNITS/Unit1/1. Input/2-DemandExtracted-Master-.xlsx`
  - `UNITS/Unit1/2. Calculations/Settings Data.xlsx`
- The local folder contains possible alternatives, but they must not be
  substituted automatically:
  - `2-DemandExtracted-Master-.Christ.xlsx`
  - `Demand-MasterRoster Manual Read.xlsx`
  - `Settings DataW.xlsx`
  - `Settings DataOld.xlsx`

The missing-path issue must be resolved by confirming the intended workbook
names or approving a manifest change. Do not rename, copy, or replace a
workbook merely to make validation pass.

## Main entry points

Run commands from the repository root.

Top-level flow: all discovered units, followed by the organisation workbooks:

```cmd
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd
```

All-units flow only:

```cmd
CLIENT\DATExx-Whiddon\UNITS\Run-AllUnitsRefresh.cmd
```

Unit 1 stage flows:

```cmd
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-RNRefresh.Stage1-RNOnly.cmd
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-RoleRefresh.Stage2-AllRoles.cmd
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-CalculationRefresh.Stage3-UnitCalculations.cmd
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-UnitRefresh.Stage4-Unit1Only.cmd
```

Running a wrapper without arguments opens its interactive menu.

## Safe Codex operating sequence

1. Confirm the intended manifest and workbook scope.
2. Confirm every selected workbook exists at its exact manifest path.
3. Confirm PowerShell 7 is available with `pwsh --version`.
4. Confirm the selected workbooks are closed and no Excel temporary lock files
   (`~$*.xlsx`) are present.
5. Run validation only. This checks paths and selection without refreshing the
   workbooks:

   ```cmd
   CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -ValidateSelectionOnly -RefreshAll
   ```

6. Review the validation result with the user. A successful validation does
   not authorize refresh.
7. After explicit approval to refresh, run the agreed full flow, stage, or
   sequence range.
8. Monitor `current-status.txt` and the latest log. Report the exact workbook
   where any failure occurs.
9. Confirm the final exit code, completed workbook range, log path, and whether
   Excel was closed normally.

## Common options

```text
-RefreshAll
-ValidateSelectionOnly
-StartAtSequence <number>
-EndAtSequence <number>
-StartAtWorkbook "<name or relative path>"
-ShowStatus
-StopMode AfterCurrent
-StopMode Now
-VisibleOverride true
-VisibleOverride false
-TimeoutMinutesOverride <minutes>
```

Examples:

```cmd
REM Validate the complete top-level flow
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -ValidateSelectionOnly -RefreshAll

REM Validate Unit 1 only
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-UnitRefresh.Stage4-Unit1Only.cmd -ValidateSelectionOnly -RefreshAll

REM Run Unit 1 calculations from Capacity.xlsx onward
CLIENT\DATExx-Whiddon\UNITS\Unit1\bin\Run-CalculationRefresh.Stage3-UnitCalculations.cmd -StartAtWorkbook "Capacity.xlsx"

REM Show top-level status
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -ShowStatus

REM Request a graceful stop after the current workbook
CLIENT\DATExx-Whiddon\Run-RefreshRunner.cmd -StopMode AfterCurrent
```

## Logs and safety behaviour

The runners maintain logs and status beside the workflow being run:

- `CLIENT/DATExx-Whiddon/RunLogs`
- `CLIENT/DATExx-Whiddon/UNITS/RunLogs`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/RunLogs`

Common files include `current-status.txt`, `latest-log.txt`,
`stop-request.txt`, and run-lock files. The runners default to hidden Excel,
check selected paths before starting Excel, check for workbook lock files, and
avoid terminating pre-existing Excel processes in safe mode. A nonzero exit
code means validation or refresh failed.

## Sequence sources of truth

- `CLIENT/DATExx-Whiddon/runner/ResidentialCare-UnitWorkbookSequence.psd1`
- `CLIENT/DATExx-Whiddon/runner/ResidentialCare-OrgWorkbookSequence.psd1`
- `CLIENT/DATExx-Whiddon/UNITS/runner/AllUnits-WorkbookSequence.psd1`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/RN-WorkbookSequence.Stage1-RNOnly.psd1`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Role-WorkbookSequence.Stage2-AllRoles.psd1`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Calculation-WorkbookSequence.Stage3-UnitCalculations.psd1`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Unit-WorkbookSequence.Stage4-Unit1Only.psd1`

Edit a sequence manifest only after the approved workbook order or exact path
has changed. The Mermaid diagrams document the manifests but do not control
execution.

