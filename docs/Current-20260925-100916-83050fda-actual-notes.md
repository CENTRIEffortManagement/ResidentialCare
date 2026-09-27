# Actual run: 20260925-100916-83050fda

Status: PartialOrFailed. Elapsed: 00:34:26. Peak recorded concurrency: 2.
Union of observed active intervals: 2065.839 seconds. Gaps including resume pauses: 1 seconds.

Technical refresh only; business reconciliation not evaluated.
Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.

## Untimed or incomplete work

- BD/DemandIntervals: Blocked; Prerequisite BD/Intervals did not complete.
- BD/Demand: Blocked; Prerequisite BD/DemandIntervals did not complete.
- BD/Shifts: Blocked; Prerequisite BD/Intervals did not complete.
- BD/ShiftAverage: Blocked; Prerequisite BD/Shifts did not complete.
- BD/Allocation: Blocked; Prerequisite BD/ShiftAverage did not complete.
- BD/Distribution: Blocked; Prerequisite BD/Allocation did not complete.
- BD/MultiRole: Blocked; Prerequisite BD/Distribution did not complete.
- BD/Role:AIN/1: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:AIN/2: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:AIN/3: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/1: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/2: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/3: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:RN/1: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:RN/2: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:RN/3: Blocked; Prerequisite BD/Demand did not complete.
- BD/Capacity: Blocked; Prerequisite BD/Role:AIN/3 did not complete.
- BD/Effort: Blocked; Prerequisite BD/Capacity did not complete.
