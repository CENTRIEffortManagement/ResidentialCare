@{
    Workflow = 'ResidentialCare All Units Refresh'
    Version = 1
    Workbooks = @(
        @{
            Sequence = 1
            Folder = '1. Input'
            FileName = '1-AllocationExtracted.xlsx'
        }
        @{
            Sequence = 2
            Folder = '1. Input'
            FileName = '2-DemandExtracted-Master-.xlsx'
        }
        @{
            Sequence = 3
            Folder = '1. Input'
            FileName = 'StaffList Availability.xlsx'
        }
        @{
            Sequence = 4
            Folder = '2. Calculations'
            FileName = 'Settings Data.xlsx'
        }
        @{
            Sequence = 5
            Folder = '1. Input'
            FileName = '2-DemandExtract.xlsx'
        }
        @{
            Sequence = 6
            Folder = '2. Calculations'
            FileName = 'Intervals.xlsx'
        }
        @{
            Sequence = 7
            Folder = '2. Calculations'
            FileName = 'DemandIntervals.xlsx'
        }
        @{
            Sequence = 8
            Folder = '2. Calculations'
            FileName = 'Demand.xlsx'
        }
        @{
            Sequence = 9
            Folder = '2. Calculations'
            FileName = 'Capacity-ShiftAvailability.xlsx'
        }
        @{
            Sequence = 10
            Folder = '2. Calculations'
            FileName = 'StaffListMaster.xlsx'
        }
        @{
            Sequence = 11
            Folder = '2. Calculations'
            FileName = 'Shifts.xlsx'
        }
        @{
            Sequence = 12
            Folder = '2. Calculations'
            FileName = 'AllocationByShiftAverage.xlsx'
        }
        @{
            Sequence = 13
            Folder = '2. Calculations'
            FileName = 'Allocation.xlsx'
        }
        @{
            Sequence = 14
            Folder = '2. Calculations'
            FileName = 'Shift-StaffDistribution.xlsx'
        }
        @{
            Sequence = 15
            Folder = '2. Calculations'
            FileName = 'MutliRoleCheck.xlsx'
        }
        @{
            Sequence = 16
            Folder = '2. Calculations/AIN'
            FileName = 'CapacityDistrib(A.1)-shifts.xlsx'
        }
        @{
            Sequence = 17
            Folder = '2. Calculations/AIN'
            FileName = 'CapacityDistrib(A.2)-shifts.xlsx'
        }
        @{
            Sequence = 18
            Folder = '2. Calculations/AIN'
            FileName = 'CapacityDistrib(B)-shifts.xlsx'
        }
        @{
            Sequence = 19
            Folder = '2. Calculations/AINC4'
            FileName = 'CapacityDistrib(A.1)-shifts.xlsx'
        }
        @{
            Sequence = 20
            Folder = '2. Calculations/AINC4'
            FileName = 'CapacityDistrib(A.2)-shifts.xlsx'
        }
        @{
            Sequence = 21
            Folder = '2. Calculations/AINC4'
            FileName = 'CapacityDistrib(B)-shifts.xlsx'
        }
        @{
            Sequence = 22
            Folder = '2. Calculations/RN'
            FileName = 'CapacityDistrib(A.1)-shifts.xlsx'
        }
        @{
            Sequence = 23
            Folder = '2. Calculations/RN'
            FileName = 'CapacityDistrib(A.2)-shifts.xlsx'
        }
        @{
            Sequence = 24
            Folder = '2. Calculations/RN'
            FileName = 'CapacityDistrib(B)-shifts.xlsx'
        }
        @{
            Sequence = 25
            Folder = '2. Calculations'
            FileName = 'Capacity.xlsx'
        }
        @{
            Sequence = 26
            Folder = '2. Calculations'
            FileName = 'Effort.xlsx'
        }
    )
}
