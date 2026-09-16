@{
    SchemaVersion = 1
    Unit = 'Unit1'
    Roles = @{
        AIN = '2. Calculations/AIN'
        AINC4 = '2. Calculations/AINC4'
        RN = '2. Calculations/RN'
        TestRole1 = '2. Calculations/TestRole1'
        TestRole2 = '2. Calculations/TestRole2'
        TestRole3 = '2. Calculations/TestRole3'
        TestRole4 = '2. Calculations/TestRole4'
    }
    WorkbookOrder = @(
        'CapacityDistrib(A.1)-shifts.xlsx'
        'CapacityDistrib(A.2)-shifts.xlsx'
        'CapacityDistrib(B)-shifts.xlsx'
    )
    MaxConcurrency = 7
    TimeoutMinutes = 30
}
