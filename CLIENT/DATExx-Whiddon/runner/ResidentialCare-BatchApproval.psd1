@{
    # Operational settings only. The historical filename and project configuration
    # key are retained for compatibility; there is no separate live-approval gate.
    OrganisationUnits = @('Unit1', 'Unit2')
    # Default menu option 1 and unqualified -RunAll scope. Organisation work is
    # excluded until its workbook inputs are configured for these Units.
    RunAllUnits = @('BD', 'TE', 'JH-RY')
    RunAllIncludeOrg = $false
    # Temporary unit-only onboarding. Existing Unit1/Unit2 roles are unchanged;
    # these Units use EN and are not organisation consumers yet.
    AdditionalUnits = @(
        @{ Folder = 'BD'; Roles = @('AIN', 'EN', 'RN') }
        @{ Folder = 'TE'; Roles = @('AIN', 'EN', 'RN') }
        @{ Folder = 'JH-RY'; Roles = @('AIN', 'EN', 'RN') }
    )
    # Key: stable job ID printed by ShowPlan (e.g. Unit1/DemandMaster).
    # Values: additional read paths relative to the repository root. Environment-rooted
    # inputs use @{ Environment = 'PUBLIC'; Path = 'Public Scripts/CentriSyncPaths.xlsx' }.
    AdditionalInputs = @{}
    # Jobs whose non-workbook write side effects have not been ruled out run alone.
    ExclusiveJobs = @()
}
