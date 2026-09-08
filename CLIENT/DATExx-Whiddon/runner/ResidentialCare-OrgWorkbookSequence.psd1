@{
    Workflow = 'ResidentialCare Date Refresh'
    Version = 1
    Workbooks = @(
        @{
            Sequence = 1
            Path = '2. Calculations/E-O-I/StafMasterList-All.xlsx'
        }
        @{
            Sequence = 2
            Path = '2. Calculations/E-O-I/Effort-All.xlsx'
        }
        @{
            Sequence = 3
            Path = '2. Calculations/E-O-I/EffortOutcomes.xlsx'
        }
        @{
            Sequence = 4
            Path = '2. Calculations/E-O-I/Inefficiencies.xlsx'
        }
        @{
            Sequence = 5
            Path = '2. Calculations/Cost/Cost..xlsx'
        }
        @{
            Sequence = 6
            Path = '2. Calculations/EOW/EffortOutcomeLogXY.xlsx'
        }
        @{
            Sequence = 7
            Path = '2. Calculations/EOW/2DRead.xlsx'
        }
        @{
            Sequence = 8
            Path = '2. Calculations/Tableau Connection.xlsx'
        }
    )
    RequiredSupportFiles = @(
        '2. Calculations/Change/AllocationChange.xlsx',
        '2. Calculations/EOW/Grid Thresholds.xlsx',
        'UNITS/Unit1/2. Calculations/Settings Data.xlsx'
    )
    ExcludedPaths = @(
        '2. Calculations/Map',
        '2. Calculations/Change/SS/ResidualAllocation.xlsx',
        '2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx'
    )
}
