@{
    # Operational settings only. The historical filename and project configuration
    # key are retained for compatibility; there is no separate live-approval gate.
    OrganisationUnits = @('BD', 'TE', 'JH-RY')
    # Default menu option 1 and unqualified -RunAll scope. Run the configured
    # Units and then the organisation release chain through Tableau Connection.
    RunAllUnits = @('BD', 'TE', 'JH-RY')
    RunAllIncludeOrg = $true
    # Current facility Units use EN. Existing Unit1/Unit2 roles are unchanged
    # and remain available for explicit selections.
    AdditionalUnits = @(
        @{ Folder = 'BD'; Roles = @('AIN', 'EN', 'RN') }
        @{ Folder = 'TE'; Roles = @('AIN', 'EN', 'RN') }
        @{ Folder = 'JH-RY'; Roles = @('AIN', 'EN', 'RN') }
    )
    # These Unit jobs now share one Date-level workbook, refreshed once per run.
    # Unit1 and Unit2 retain their existing Unit-level demand masters.
    SharedUnitJobs = @(
        @{ Id = 'DemandMaster'; Units = @('BD', 'TE', 'JH-RY'); Path = '1. Input/Demand-MasterRoster Manual Read.xlsx' }
    )
    # Key: stable job ID printed by ShowPlan (e.g. Unit1/DemandMaster).
    # Values: additional read paths relative to the repository root. Environment-rooted
    # inputs use @{ Environment = 'PUBLIC'; Path = 'Public Scripts/CentriSyncPaths.xlsx' }.
    AdditionalInputs = @{}
    # Jobs whose non-workbook write side effects have not been ruled out run alone.
    ExclusiveJobs = @()
}
