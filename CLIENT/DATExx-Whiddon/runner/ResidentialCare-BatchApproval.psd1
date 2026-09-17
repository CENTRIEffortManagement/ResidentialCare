@{
    # Operational settings only. The historical filename and project configuration
    # key are retained for compatibility; there is no separate live-approval gate.
    OrganisationUnits = @('Unit1', 'Unit2')
    # Key: stable job ID printed by ShowPlan (e.g. Unit1/DemandMaster).
    # Values: additional read paths relative to the repository root. Environment-rooted
    # inputs use @{ Environment = 'PUBLIC'; Path = 'Public Scripts/CentriSyncPaths.xlsx' }.
    AdditionalInputs = @{}
    # Jobs whose non-workbook write side effects have not been ruled out run alone.
    ExclusiveJobs = @()
}
