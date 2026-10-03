# Actual run: 20261002-185510-a07eff9f

Status: PartialOrFailed. Elapsed: 00:10:49. Peak recorded concurrency: 6.
Union of observed active intervals: 646.807 seconds. Gaps including resume pauses: 3 seconds.

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
- BD/Role:RN/1: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:RN/2: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:RN/3: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/1: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/2: Blocked; Prerequisite BD/Demand did not complete.
- BD/Role:EN/3: Blocked; Prerequisite BD/Demand did not complete.
- BD/Capacity: Blocked; Prerequisite BD/Role:AIN/3 did not complete.
- BD/Effort: Blocked; Prerequisite BD/Capacity did not complete.
- TE/Demand: Blocked; Prerequisite TE/DemandIntervals did not complete.
- TE/Role:AIN/1: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:AIN/2: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:AIN/3: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:RN/1: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:RN/2: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:RN/3: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:EN/1: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:EN/2: Blocked; Prerequisite TE/Demand did not complete.
- TE/Role:EN/3: Blocked; Prerequisite TE/Demand did not complete.
- TE/Capacity: Blocked; Prerequisite TE/Role:AIN/3 did not complete.
- TE/Effort: Blocked; Prerequisite TE/Capacity did not complete.
- JH-RY/Demand: Blocked; Prerequisite JH-RY/DemandIntervals did not complete.
- JH-RY/Role:AIN/1: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:AIN/2: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:AIN/3: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:RN/1: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:RN/2: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:RN/3: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:EN/1: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:EN/2: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:EN/3: Blocked; Prerequisite JH-RY/Demand did not complete.
- JH-RY/Capacity: Blocked; Prerequisite JH-RY/Role:AIN/3 did not complete.
- JH-RY/Effort: Blocked; Prerequisite JH-RY/Capacity did not complete.
- Org/EffortAll: Blocked; Prerequisite TE/Effort did not complete.
- Org/Outcomes: Blocked; Prerequisite Org/EffortAll did not complete.
- Org/Inefficiencies: Blocked; Prerequisite Org/Outcomes did not complete.
- Org/Cost: Blocked; Prerequisite Org/Inefficiencies did not complete.
- Org/History: Blocked; Prerequisite Org/Outcomes did not complete.
- Org/HistoryRead: Blocked; Prerequisite Org/History did not complete.
- Org/Reporting: Blocked; Prerequisite Org/Cost did not complete.
