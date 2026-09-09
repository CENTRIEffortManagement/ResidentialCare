# ResidentialCare runner handover recap

## Local workflow

The active local workflow is `CLIENT/DATExx-Whiddon`. The runner package was
copied from the Production ResidentialCare `Whiddon/DATExx` tree and verified
file-for-file after copying.

The available entry points are:

- `CLIENT/DATExx-Whiddon/Run-RefreshRunner.cmd`
- `CLIENT/DATExx-Whiddon/UNITS/Run-AllUnitsRefresh.cmd`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-RNRefresh.Stage1-RNOnly.cmd`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-RoleRefresh.Stage2-AllRoles.cmd`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-CalculationRefresh.Stage3-UnitCalculations.cmd`
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/bin/Run-UnitRefresh.Stage4-Unit1Only.cmd`

## Documentation

- `docs/cmd-runner-instructions.md` — operating and safety instructions.
- `docs/ResidentialCare-Runner-Sequences.md` — readable workbook orders and
  explanation of each runner level.
- `docs/ResidentialCare-Runner-Overview.mmd` — Codex/preflight/top-level flow.
- `docs/ResidentialCare-Unit1-Sequence.mmd` — complete 26-workbook Unit 1
  execution order.
- `docs/ResidentialCare-Org-Sequence.mmd` — eight-workbook organisation order.

## Current readiness status

- Runner scripts and manifests parse successfully.
- PowerShell 7.6.5 is installed and available as `pwsh`.
- Unit 1 validation is expected to stop because the manifest requires
  `2-DemandExtracted-Master-.xlsx` and `Settings Data.xlsx`, which are not
  present under those exact local names.
- Only `Unit1` is currently discoverable under `CLIENT/DATExx-Whiddon/UNITS`.
- No Excel workbook was refreshed as part of copying, documenting, or checking
  the runner package.
- A top-level validation-only run on 9 September 2026 enumerated the complete
  34-workbook flow and stopped at global sequence 002 because
  `2-DemandExtracted-Master-.xlsx` was missing. Excel was not opened.

## Next controlled decision

Confirm which existing workbook corresponds to each missing manifest entry, or
confirm that Production should supply the exact missing files. After the paths
are resolved, Codex can rerun validation-only, report the selected sequence,
and wait for explicit approval before starting a real Excel refresh.
