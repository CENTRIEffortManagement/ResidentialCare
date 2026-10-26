# Forecast: Current

Estimated elapsed: unknown - incomplete timing history. Peak scheduled workers: 7.

Historical coordinator timings include process overhead. Input sizes and concurrency may change actual durations.

Original 17m 23s chart retained separately as an unverified illustrative baseline.

## Untimed work

- BD/ShiftAverage — predecessors: BD/Shifts, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Allocation — predecessors: BD/ShiftAverage, BD/AllocationInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Distribution — predecessors: BD/Allocation, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:CS/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:CS/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:CS/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:CS/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:CS/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Capacity — predecessors: BD/Role:AIN/3, BD/Role:EN/3, BD/Role:RN/3, BD/Role:CS/3, BD/Demand, BD/Staff, BD/Availability, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Effort — predecessors: BD/Capacity, BD/Demand, BD/Allocation, BD/Distribution, BD/Staff, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Distribution — predecessors: TE/Allocation, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:AIN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:AIN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:RN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:RN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:EN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:EN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:CS/1 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:CS/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:CS/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:CS/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:CS/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Capacity — predecessors: TE/Role:AIN/3, TE/Role:EN/3, TE/Role:RN/3, TE/Role:CS/3, TE/Demand, TE/Staff, TE/Availability, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Effort — predecessors: TE/Capacity, TE/Demand, TE/Allocation, TE/Distribution, TE/Staff, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Demand — predecessors: JH-RY/DemandIntervals, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Allocation — predecessors: JH-RY/ShiftAverage, JH-RY/AllocationInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Distribution — predecessors: JH-RY/Allocation, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:CS/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:CS/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:CS/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:CS/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:CS/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Capacity — predecessors: JH-RY/Role:AIN/3, JH-RY/Role:EN/3, JH-RY/Role:RN/3, JH-RY/Role:CS/3, JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Effort — predecessors: JH-RY/Capacity, JH-RY/Demand, JH-RY/Allocation, JH-RY/Distribution, JH-RY/Staff, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- Org/EffortAll — predecessors: Org/StaffAll, BD/Effort, TE/Effort, JH-RY/Effort; Unknown duration or start affected by unknown slot occupancy.
- Org/Outcomes — predecessors: Org/EffortAll, Org/StaffAll; Unknown duration or start affected by unknown slot occupancy.
- Org/Inefficiencies — predecessors: Org/Outcomes; Unknown duration or start affected by unknown slot occupancy.
- Org/Cost — predecessors: Org/Inefficiencies, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- Org/History — predecessors: Org/Outcomes; Unknown duration or start affected by unknown slot occupancy.
- Org/HistoryRead — predecessors: Org/History; Unknown duration or start affected by unknown slot occupancy.
- Org/Reporting — predecessors: Org/Cost, Org/HistoryRead, Org/Inefficiencies; Unknown duration or start affected by unknown slot occupancy.
- Org/LeaveBalanceStats — predecessors: Org/Reporting; Unknown duration or start affected by unknown slot occupancy.
- Org/LeaveBalanceReduction — predecessors: Org/LeaveBalanceStats; Unknown duration or start affected by unknown slot occupancy.
- Org/LeaveBalanceReporting — predecessors: Org/LeaveBalanceReduction; Unknown duration or start affected by unknown slot occupancy.
