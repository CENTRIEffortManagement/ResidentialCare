# Residential Care Refresh Order

## Recommended Sequence

1. Input workbooks
2. Settings Data.xlsx
3. Intervals.xlsx
4. DemandIntervals.xlsx
5. Demand.xlsx
6. Capacity-ShiftAvailability.xlsx
7. StaffListMaster.xlsx
8. Capacity.xlsx
9. Shifts.xlsx
10. AllocationByShiftAverage.xlsx
11. Allocation.xlsx
12. Effort.xlsx
13. MutliRoleCheck.xlsx
14. Shift-StaffDistribution.xlsx

## Flowchart

```mermaid
flowchart TD
    Input["1. Input workbooks<br/>1-AllocationExtracted.xlsx<br/>2-DemandExtract.xlsx<br/>StaffList Availability.xlsx<br/>Demand-MasterRoster Manual Read.xlsx"]
    RoleCapacity["Role capacity distribution files<br/>Role folders / CapacityDistrib*.xlsx"]

    Settings["Settings Data.xlsx"]
    Intervals["Intervals.xlsx"]
    DemandIntervals["DemandIntervals.xlsx"]
    Demand["Demand.xlsx"]
    CapacityShiftAvailability["Capacity-ShiftAvailability.xlsx"]
    StaffListMaster["StaffListMaster.xlsx"]
    Capacity["Capacity.xlsx"]
    Shifts["Shifts.xlsx"]
    AllocationByShiftAverage["AllocationByShiftAverage.xlsx"]
    Allocation["Allocation.xlsx"]
    Effort["Effort.xlsx"]
    MultiRoleCheck["MutliRoleCheck.xlsx"]
    ShiftStaffDistribution["Shift-StaffDistribution.xlsx"]

    Input --> Settings
    Input --> Intervals
    Settings --> Intervals

    Input --> DemandIntervals
    Settings --> DemandIntervals
    Intervals --> DemandIntervals

    Settings --> Demand
    Intervals --> Demand
    DemandIntervals --> Demand
    Input --> Demand

    Input --> CapacityShiftAvailability
    Settings --> CapacityShiftAvailability

    CapacityShiftAvailability --> StaffListMaster
    Input --> StaffListMaster

    Settings --> Capacity
    RoleCapacity --> Capacity

    Settings --> Shifts
    Intervals --> Shifts
    Input --> Shifts

    Settings --> AllocationByShiftAverage
    Shifts --> AllocationByShiftAverage

    Settings --> Allocation
    AllocationByShiftAverage --> Allocation
    Input --> Allocation

    Demand --> Effort
    Capacity --> Effort
    Allocation --> Effort
    Settings --> Effort
    StaffListMaster --> Effort

    CapacityShiftAvailability --> MultiRoleCheck
    Settings --> MultiRoleCheck
    Shifts --> MultiRoleCheck

    Shifts --> ShiftStaffDistribution
```

## Notes

- Refresh `Settings Data.xlsx` before dependent calculation workbooks because it provides dates, roles, shifts, meals, shift starts, and permutation dimensions.
- Refresh `Intervals.xlsx` before demand and shift calculations because it supplies the interval spine.
- Refresh `DemandIntervals.xlsx` before `Demand.xlsx`; demand rolls interval-level demand up to shift/date/role outputs.
- Refresh `Shifts.xlsx` before allocation averages and staff distribution outputs.
- Refresh `AllocationByShiftAverage.xlsx` before `Allocation.xlsx`; allocation depends on its role/resource shift summaries.
- Refresh `Effort.xlsx` after demand, capacity, allocation, settings, and staff master data are current.
- `MutliRoleCheck.xlsx` and `Shift-StaffDistribution.xlsx` are downstream diagnostic/reporting workbooks.
