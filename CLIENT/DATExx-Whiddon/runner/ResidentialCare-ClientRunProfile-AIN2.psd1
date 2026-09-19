@{
    SchemaVersion = 1
    Client = 'DATExx-Whiddon'
    LastRoleConfirmationDate = '2026-09-19'

    # Safety boundary for the temporary Unit1/AIN2 two-workbook prototype.
    # Selection code rejects RunAll, organisation workbooks, dependency
    # scheduling, legacy ranges, other Units and all non-role jobs.
    IsolatedRoleOnly = $true
    AllowedUnits = @('Unit1')

    Roles = @(
        @{
            Folder = 'AIN2'
            Enabled = $true
            WorkbookOrder = @(
                'CapacityDistrib(A.1)-shifts.xlsx'
                'CapacityDistrib(B)-shifts.xlsx'
            )
        }
    )
}
