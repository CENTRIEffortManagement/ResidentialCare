@{
    SchemaVersion = 1
    Client = 'DATExx-Whiddon'
    LastRoleConfirmationDate = '2026-09-19'

    # Prepared post-acceptance profile. It is intentionally not registered in
    # pq.project.json and cannot be selected by the runner before promotion.
    Roles = @(
        @{
            Folder = 'AIN'
            Enabled = $true
            WorkbookOrder = @(
                'CapacityDistrib(A.1)-shifts.xlsx'
                'CapacityDistrib(B)-shifts.xlsx'
            )
        }
        @{
            Folder = 'AINC4'
            Enabled = $true
            WorkbookOrder = @(
                'CapacityDistrib(A.1)-shifts.xlsx'
                'CapacityDistrib(A.2)-shifts.xlsx'
                'CapacityDistrib(B)-shifts.xlsx'
            )
        }
        @{
            Folder = 'RN'
            Enabled = $true
            WorkbookOrder = @(
                'CapacityDistrib(A.1)-shifts.xlsx'
                'CapacityDistrib(A.2)-shifts.xlsx'
                'CapacityDistrib(B)-shifts.xlsx'
            )
        }
    )
}
