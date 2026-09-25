Set-StrictMode -Version Latest

function Read-ResidentialCareRunScope {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Client
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Run scope file is missing: $Path" }
    try { $scope = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable -ErrorAction Stop }
    catch { throw "Run scope is not valid JSON: $($_.Exception.Message)" }
    if ($scope -isnot [System.Collections.IDictionary]) { throw 'Run scope must be a JSON object.' }
    $expected = @('schemaVersion', 'client', 'enabledUnits', 'roles')
    foreach ($key in $expected) {
        if (-not $scope.Contains($key)) { throw "Run scope is missing $key." }
    }
    foreach ($key in $scope.Keys) {
        if ($key -notin $expected) { throw "Unknown run scope property: $key" }
    }
    if ($scope.schemaVersion -isnot [long] -and $scope.schemaVersion -isnot [int]) { throw 'Run scope schemaVersion must be an integer.' }
    if ($scope.schemaVersion -ne 1) { throw "Unsupported run scope schemaVersion: $($scope.schemaVersion)" }
    if ($scope.client -isnot [string] -or $scope.client -cne $Client) { throw "Run scope client must be $Client." }
    foreach ($field in @('enabledUnits', 'roles')) {
        $values = $scope[$field]
        if ($values -isnot [array] -or $values.Count -eq 0) { throw "Run scope $field must be a nonempty array." }
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($value in $values) {
            # A scope entry is one immediate folder name, never a relative path.
            if ($value -isnot [string] -or $value -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]*$') {
                throw "Run scope $field contains an unsafe folder name: $value"
            }
            if (-not $seen.Add($value)) { throw "Run scope $field contains a duplicate folder: $value" }
        }
    }
    return [pscustomobject]@{
        SchemaVersion = 1
        Client = $scope.client
        Units = [string[]] @($scope.enabledUnits)
        Roles = [string[]] @($scope.roles)
        Path = [IO.Path]::GetFullPath($Path)
    }
}

function Assert-ResidentialCareRunScopeFolders {
    param(
        [Parameter(Mandatory)] $Scope,
        [Parameter(Mandatory)] [string] $UnitsRoot,
        [Parameter(Mandatory)] [string[]] $RoleWorkbookOrder
    )

    foreach ($unit in $Scope.Units) {
        $unitPath = Join-Path $UnitsRoot $unit
        if (-not (Test-Path -LiteralPath $unitPath -PathType Container)) { throw "Enabled Unit folder is missing: $unit" }
        if ((Get-Item -LiteralPath $unitPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Enabled Unit folder is a reparse point: $unit"
        }
        foreach ($role in $Scope.Roles) {
            $rolePath = Join-Path $unitPath "2. Calculations/$role"
            if (-not (Test-Path -LiteralPath $rolePath -PathType Container)) { throw "Enabled role folder is missing: $unit/$role" }
            if ((Get-Item -LiteralPath $rolePath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Enabled role folder is a reparse point: $unit/$role"
            }
            foreach ($workbook in $RoleWorkbookOrder) {
                if (-not (Test-Path -LiteralPath (Join-Path $rolePath $workbook) -PathType Leaf)) {
                    throw "Enabled role workbook is missing: $unit/$role/$workbook"
                }
            }
        }
    }
}
