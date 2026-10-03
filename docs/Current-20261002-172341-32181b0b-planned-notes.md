# Forecast: Current

Estimated elapsed: unknown - incomplete timing history. Peak scheduled workers: 3.

Historical coordinator timings include process overhead. Input sizes and concurrency may change actual durations.

Original 17m 23s chart retained separately as an unverified illustrative baseline.

## Untimed work

- Date/DemandMaster — predecessors: BD/Settings, TE/Settings, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/DemandInput — predecessors: Date/DemandMaster, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Availability — predecessors: BD/Workers, BD/AllocationInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Staff — predecessors: BD/Availability, BD/Workers, BD/AllocationInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Intervals — predecessors: BD/AllocationInput, BD/DemandInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/DemandIntervals — predecessors: BD/Intervals, BD/DemandInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Demand — predecessors: BD/DemandIntervals, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Shifts — predecessors: BD/Intervals, BD/AllocationInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/ShiftAverage — predecessors: BD/Shifts, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Allocation — predecessors: BD/ShiftAverage, BD/AllocationInput, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Distribution — predecessors: BD/Allocation, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/MultiRole — predecessors: BD/Distribution, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:AIN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:RN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/1 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/2 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- BD/Role:EN/3 — predecessors: BD/Demand, BD/Staff, BD/Availability, BD/ShiftAverage, BD/Settings, BD/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- BD/Capacity — predecessors: BD/Role:AIN/3, BD/Role:EN/3, BD/Role:RN/3, BD/Demand, BD/Staff, BD/Availability, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- BD/Effort — predecessors: BD/Capacity, BD/Demand, BD/Allocation, BD/Distribution, BD/MultiRole, BD/Staff, BD/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/DemandInput — predecessors: Date/DemandMaster, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Workers — predecessors: TE/AllocationInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Availability — predecessors: TE/Workers, TE/AllocationInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Staff — predecessors: TE/Availability, TE/Workers, TE/AllocationInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Intervals — predecessors: TE/AllocationInput, TE/DemandInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/DemandIntervals — predecessors: TE/Intervals, TE/DemandInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Demand — predecessors: TE/DemandIntervals, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Shifts — predecessors: TE/Intervals, TE/AllocationInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/ShiftAverage — predecessors: TE/Shifts, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Allocation — predecessors: TE/ShiftAverage, TE/AllocationInput, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Distribution — predecessors: TE/Allocation, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/MultiRole — predecessors: TE/Distribution, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:AIN/1 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:AIN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:AIN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:RN/1 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:RN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:RN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:EN/1 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:EN/2 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- TE/Role:EN/3 — predecessors: TE/Demand, TE/Staff, TE/Availability, TE/ShiftAverage, TE/Settings, TE/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- TE/Capacity — predecessors: TE/Role:AIN/3, TE/Role:EN/3, TE/Role:RN/3, TE/Demand, TE/Staff, TE/Availability, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- TE/Effort — predecessors: TE/Capacity, TE/Demand, TE/Allocation, TE/Distribution, TE/MultiRole, TE/Staff, TE/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/DemandInput — predecessors: Date/DemandMaster, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Availability — predecessors: JH-RY/Workers, JH-RY/AllocationInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Staff — predecessors: JH-RY/Availability, JH-RY/Workers, JH-RY/AllocationInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Intervals — predecessors: JH-RY/AllocationInput, JH-RY/DemandInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/DemandIntervals — predecessors: JH-RY/Intervals, JH-RY/DemandInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Demand — predecessors: JH-RY/DemandIntervals, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Shifts — predecessors: JH-RY/Intervals, JH-RY/AllocationInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/ShiftAverage — predecessors: JH-RY/Shifts, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Allocation — predecessors: JH-RY/ShiftAverage, JH-RY/AllocationInput, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Distribution — predecessors: JH-RY/Allocation, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/MultiRole — predecessors: JH-RY/Distribution, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:AIN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:AIN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:AIN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:RN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:RN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:RN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/1 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/2 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:EN/1; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Role:EN/3 — predecessors: JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/ShiftAverage, JH-RY/Settings, JH-RY/Role:EN/2; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Capacity — predecessors: JH-RY/Role:AIN/3, JH-RY/Role:EN/3, JH-RY/Role:RN/3, JH-RY/Demand, JH-RY/Staff, JH-RY/Availability, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- JH-RY/Effort — predecessors: JH-RY/Capacity, JH-RY/Demand, JH-RY/Allocation, JH-RY/Distribution, JH-RY/MultiRole, JH-RY/Staff, JH-RY/Settings; Unknown duration or start affected by unknown slot occupancy.
- Org/StaffAll — predecessors: BD/Staff, TE/Staff, JH-RY/Staff; Unknown duration or start affected by unknown slot occupancy.
- Org/EffortAll — predecessors: Org/StaffAll, BD/Effort, TE/Effort, JH-RY/Effort; Unknown duration or start affected by unknown slot occupancy.
- Org/Outcomes — predecessors: Org/EffortAll, Org/StaffAll; Unknown duration or start affected by unknown slot occupancy.
- Org/Inefficiencies — predecessors: Org/Outcomes; Unknown duration or start affected by unknown slot occupancy.
- Org/Cost — predecessors: Org/Inefficiencies, Unit1/Settings; Unknown duration or start affected by unknown slot occupancy.
- Org/History — predecessors: Org/Outcomes; Unknown duration or start affected by unknown slot occupancy.
- Org/HistoryRead — predecessors: Org/History; Unknown duration or start affected by unknown slot occupancy.
- Org/Reporting — predecessors: Org/Cost, Org/HistoryRead, Org/Inefficiencies; Unknown duration or start affected by unknown slot occupancy.
