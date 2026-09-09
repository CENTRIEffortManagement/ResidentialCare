@{
    SchemaVersion = 1
    Client = 'DATExx-Whiddon'
    LastRoleConfirmationDate = '2026-09-09'

    # The listed order is the approved role execution order. Normal runs read
    # this saved list and do not scan role folders.
    Roles = @(
        @{
            Folder = 'AIN'
            Enabled = $true
        }
        @{
            Folder = 'AINC4'
            Enabled = $true
        }
        @{
            Folder = 'RN'
            Enabled = $true
        }
    )

    # Every enabled role folder is expected to contain these workbooks in this
    # order. A role check reports incomplete folders but never repairs them.
    RoleWorkbookOrder = @(
        'CapacityDistrib(A.1)-shifts.xlsx'
        'CapacityDistrib(A.2)-shifts.xlsx'
        'CapacityDistrib(B)-shifts.xlsx'
    )
}
