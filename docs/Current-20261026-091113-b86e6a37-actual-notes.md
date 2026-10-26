# Actual run: 20261026-091113-b86e6a37

Status: PartialOrFailed. Elapsed: 00:02:33. Peak recorded concurrency: 2.
Union of observed active intervals: 152.182 seconds. Gaps including resume pauses: 1 seconds.

Technical refresh only; business reconciliation not evaluated.
Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.

**OUTPUTS NOT VERIFIED/UPDATED: 3 of 7.**

Metadata assessment only; later manual saves are not runner receipts. See output-status.md for every output.
- BD/ShiftAverage: NotUpdated; Not updated by this run (Blocked). Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
- BD/Allocation: NotUpdated; Not updated by this run (Blocked). Waiting for BD/ShiftAverage The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
- BD/Distribution: NotUpdated; Not updated by this run (Blocked). Waiting for BD/Allocation The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx

## Untimed or incomplete work

- BD/ShiftAverage: Blocked; Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
- BD/Allocation: Blocked; Waiting for BD/ShiftAverage
- BD/Distribution: Blocked; Waiting for BD/Allocation
