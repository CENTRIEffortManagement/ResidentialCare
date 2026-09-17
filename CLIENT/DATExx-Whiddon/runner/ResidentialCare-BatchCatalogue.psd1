@{
    SchemaVersion = 1
    # Sequence references resolve paths from the unchanged legacy manifests.
    BatchOrder = @('A1', 'S1', 'D1', 'C1', 'D2', 'A2', 'A3', 'C2', 'U4', 'O1', 'O2', 'O3', 'O4', 'O5')
    BatchTitles = @{
        A1 = 'Allocation transformation'
        S1 = 'Shared controls'
        D1 = 'Demand transformation'
        C1 = 'Capacity inputs'
        D2 = 'Demand preparation'
        A2 = 'Role-allocation basis'
        A3 = 'Allocation outputs'
        C2 = 'Role capacity'
        U4 = 'Unit consolidation'
        O1 = 'Cross-unit assembly'
        O2 = 'Outcomes'
        O3 = 'Cost'
        O4 = 'Outcome history'
        O5 = 'Reporting release'
    }
    UnitJobs = @(
        @{ Id = 'AllocationInput'; Batch = 'A1'; Sequence = 1; Depends = @() }
        @{ Id = 'Settings'; Batch = 'S1'; Sequence = 3; Depends = @('AllocationInput') }
        @{ Id = 'DemandMaster'; Batch = 'D1'; Path = '1. Input/Demand-MasterRoster Manual Read.xlsx'; Depends = @('Settings') }
        @{ Id = 'DemandInput'; Batch = 'D1'; Sequence = 2; Depends = @('DemandMaster', 'Settings') }
        @{ Id = 'Workers'; Batch = 'C1'; Path = '1. Input/Worker Reconciliation.xlsx'; Depends = @('AllocationInput', 'Settings') }
        @{ Id = 'Availability'; Batch = 'C1'; Sequence = 7; Depends = @('Workers', 'AllocationInput', 'Settings'); Inputs = @('1. Input/whiddon_availability_leave_extraction.xlsx') }
        @{ Id = 'Staff'; Batch = 'C1'; Sequence = 8; Depends = @('Availability', 'Workers', 'AllocationInput', 'Settings') }
        @{ Id = 'Intervals'; Batch = 'D2'; Sequence = 4; Depends = @('AllocationInput', 'DemandInput', 'Settings') }
        @{ Id = 'DemandIntervals'; Batch = 'D2'; Sequence = 5; Depends = @('Intervals', 'DemandInput', 'Settings') }
        @{ Id = 'Demand'; Batch = 'D2'; Sequence = 6; Depends = @('DemandIntervals', 'Settings') }
        @{ Id = 'Shifts'; Batch = 'A2'; Sequence = 9; Depends = @('Intervals', 'AllocationInput', 'Settings') }
        @{ Id = 'ShiftAverage'; Batch = 'A2'; Sequence = 10; Depends = @('Shifts', 'Settings') }
        @{ Id = 'Allocation'; Batch = 'A3'; Sequence = 11; Depends = @('ShiftAverage', 'AllocationInput', 'Settings') }
        @{ Id = 'Distribution'; Batch = 'A3'; Sequence = 12; Depends = @('Allocation', 'Settings') }
        @{ Id = 'MultiRole'; Batch = 'A3'; Sequence = 13; Depends = @('Distribution', 'Settings') }
        @{ Id = 'Capacity'; Batch = 'U4'; Sequence = 23; Depends = @('RoleOutputs', 'Demand', 'Staff', 'Availability', 'Settings') }
        @{ Id = 'Effort'; Batch = 'U4'; Sequence = 24; Depends = @('Capacity', 'Demand', 'Allocation', 'Distribution', 'MultiRole', 'Staff', 'Settings') }
    )
    RoleDependencies = @('Demand', 'Staff', 'Availability', 'ShiftAverage', 'Settings')
    OrgJobs = @(
        @{ Id = 'StaffAll'; Batch = 'O1'; Sequence = 1; Depends = @(); UnitDepends = @('Staff') }
        @{ Id = 'EffortAll'; Batch = 'O1'; Sequence = 2; Depends = @('StaffAll'); UnitDepends = @('Effort'); Inputs = @('2. Calculations/Change/AllocationChange.xlsx') }
        @{ Id = 'Outcomes'; Batch = 'O2'; Sequence = 3; Depends = @('EffortAll', 'StaffAll'); UnitDepends = @() }
        @{ Id = 'Inefficiencies'; Batch = 'O2'; Sequence = 4; Depends = @('Outcomes'); UnitDepends = @() }
        @{ Id = 'Cost'; Batch = 'O3'; Sequence = 5; Depends = @('Inefficiencies'); UnitDepends = @(); Inputs = @('UNITS/Unit1/2. Calculations/Settings Data.xlsx') }
        @{ Id = 'History'; Batch = 'O4'; Sequence = 6; Depends = @('Outcomes'); UnitDepends = @() }
        @{ Id = 'HistoryRead'; Batch = 'O4'; Sequence = 7; Depends = @('History'); UnitDepends = @(); Inputs = @('2. Calculations/EOW/Grid Thresholds.xlsx') }
        @{ Id = 'Reporting'; Batch = 'O5'; Sequence = 8; Depends = @('Cost', 'HistoryRead', 'Inefficiencies'); UnitDepends = @() }
    )
}
