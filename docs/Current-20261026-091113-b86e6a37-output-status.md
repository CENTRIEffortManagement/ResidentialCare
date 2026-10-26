# Output status: 20261026-091113-b86e6a37

Historical run status: PartialOrFailed. Checked UTC: 2026-10-25T22:13:47.0300559Z.

**OUTPUTS NOT VERIFIED/UPDATED: 3 of 7.**

Filesystem metadata only. This checker did not open, refresh, save or change a workbook. A later manual save is not a runner completion receipt. Business results were not validated.

- BD/Intervals: **Verified**; UNITS/BD/2. Calculations/Intervals.xlsx; last saved UTC: 2026-10-25T22:11:25.9222431Z. Confirmed run output and all recorded declared input metadata still match.
- BD/DemandIntervals: **Verified**; UNITS/BD/2. Calculations/DemandIntervals.xlsx; last saved UTC: 2026-10-25T22:12:47.6798363Z. Confirmed run output and all recorded declared input metadata still match.
- BD/Demand: **Verified**; UNITS/BD/2. Calculations/Demand.xlsx; last saved UTC: 2026-10-25T22:13:03.6214434Z. Confirmed run output and all recorded declared input metadata still match.
- BD/Shifts: **Verified**; UNITS/BD/2. Calculations/Shifts.xlsx; last saved UTC: 2026-10-25T22:13:43.6831507Z. Confirmed run output and all recorded declared input metadata still match.
- BD/ShiftAverage: **NotUpdated**; UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx; last saved UTC: 2026-10-25T02:54:27.5490264Z. Not updated by this run (Blocked). Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
- BD/Allocation: **NotUpdated**; UNITS/BD/2. Calculations/Allocation.xlsx; last saved UTC: 2026-10-25T02:54:58.5337697Z. Not updated by this run (Blocked). Waiting for BD/ShiftAverage The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
- BD/Distribution: **NotUpdated**; UNITS/BD/2. Calculations/Shift-StaffDistribution.xlsx; last saved UTC: 2026-10-25T02:55:26.8054252Z. Not updated by this run (Blocked). Waiting for BD/Allocation The output predates a current declared input.
  Root cause: BD/ShiftAverage: Excel lock file found: C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\2. Calculations\~$AllocationByShiftAverage.xlsx
