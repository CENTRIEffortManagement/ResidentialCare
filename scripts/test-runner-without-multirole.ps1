param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=Split-Path $PSScriptRoot -Parent
$temporaryRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
$testRoot=Join-Path $temporaryRoot ('rc-without-multirole-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$script:assertions=0
$beforeExcel=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
function Assert-Sequence {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:assertions++
}
function Assert-SequenceThrows {
    param([scriptblock]$Action,[string]$Message)
    $failed=$false
    try { & $Action | Out-Null } catch { $failed=$true }
    Assert-Sequence $failed $Message
}
function Get-SequenceFunctions {
    param([string]$RelativePath,[string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $RelativePath),[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw "Source did not parse: $RelativePath" }
    $definitions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false))
    $text=foreach ($name in $Names) {
        $matches=@($definitions | Where-Object Name -eq $name)
        if ($matches.Count -ne 1) { throw "Expected one $name in $RelativePath" }
        $matches[0].Extent.Text
    }
    return ($text -join "`n")
}
try {
    # Only configuration is copied. No production or synthetic Excel file is opened.
    $rootManifest=Join-Path $testRoot 'date-sequence.psd1'
    $unitsManifest=Join-Path $testRoot 'units-sequence.psd1'
    Copy-Item -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/ResidentialCare-UnitWorkbookSequence.psd1') -Destination $rootManifest
    Copy-Item -LiteralPath (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/runner/AllUnits-WorkbookSequence.psd1') -Destination $unitsManifest
    $orgSequencePath=Join-Path $testRoot 'organisation-sequence.psd1'
    [IO.File]::WriteAllText($orgSequencePath,"@{ Workbooks = @(@{ Sequence = 1; Path = 'Organisation.data' }) }")
    $unitsRoot=Join-Path $testRoot 'UNITS'
    foreach ($unitName in @('Unit1','Unit2')) {
        foreach ($folder in @('1. Input','2. Calculations','2. Calculations/AIN','2. Calculations/AINC4','2. Calculations/RN')) {
            [void][IO.Directory]::CreateDirectory((Join-Path (Join-Path $unitsRoot $unitName) $folder))
        }
    }
    $fixtureUnits=@(1,2 | ForEach-Object { [pscustomobject]@{Name="Unit$_";Number=$_;Path=(Join-Path $unitsRoot "Unit$_")} })
    $expectedLocal=@(1..12)+@(14..24)
    $expectedGlobal=@($expectedLocal)+@($expectedLocal | ForEach-Object {$_+24})
    foreach ($case in @(
        @{Source='CLIENT/DATExx-Whiddon/runner/src/Invoke-AllUnitsWorkflowRefresh.ps1';Manifest=$rootManifest},
        @{Source='CLIENT/DATExx-Whiddon/UNITS/runner/src/Invoke-WorkflowRefresh.ps1';Manifest=$unitsManifest}
    )) {
        $functions=Get-SequenceFunctions $case.Source @('Get-NormalizedFullPath','Test-IsUnderRoot','Assert-RequiredFolder','Resolve-PerUnitWorkbookEntries','Expand-AllUnitWorkbookEntries','Select-WorkbookEntries')
        & {
            . ([scriptblock]::Create($functions))
            $data=Import-PowerShellDataFile -LiteralPath $case.Manifest
            $local=@(Resolve-PerUnitWorkbookEntries $data)
            Assert-Sequence ($local.Count -eq 23) "$($case.Source): exactly 23 executable local workbooks"
            Assert-Sequence (($local.LocalSequence -join ',') -eq ($expectedLocal -join ',')) "$($case.Source): sequence 13 is reserved without renumbering"
            Assert-Sequence (-not @($local | Where-Object FileName -like '*MutliRole*').Count) "$($case.Source): MultiRole cannot be selected"
            $expanded=@(Expand-AllUnitWorkbookEntries -Units $fixtureUnits -PerUnitEntries $local)
            Assert-Sequence ($expanded.Count -eq 46) "$($case.Source): two Units have 46 executable files"
            Assert-Sequence (($expanded.Sequence -join ',') -eq ($expectedGlobal -join ',')) "$($case.Source): both global slots 13 and 37 remain reserved"
            Assert-Sequence (@($expanded | Where-Object UnitName -eq 'Unit2')[0].Sequence -eq 25) "$($case.Source): Unit2 still starts at 25"
            Assert-Sequence ($expanded[-1].Sequence -eq 48 -and $expanded[-1].FileName -eq 'Effort.xlsx') "$($case.Source): Unit2 Effort retains global ID 48"
            $StartAtWorkbook=''; $StartAtSequence=13; $EndAtSequence=0
            Assert-SequenceThrows { Select-WorkbookEntries -Entries $expanded } "$($case.Source): removed start selector 13 is rejected"
            $StartAtSequence=12; $EndAtSequence=13
            Assert-SequenceThrows { Select-WorkbookEntries -Entries $expanded } "$($case.Source): removed end selector 13 is rejected"
            $StartAtSequence=12; $EndAtSequence=14
            Assert-Sequence ((@(Select-WorkbookEntries -Entries $expanded).Sequence -join ',') -eq '12,14') "$($case.Source): a range can span the reserved slot"
            $EndAtSequence=24
            Assert-Sequence (@(Select-WorkbookEntries -Entries $expanded)[-1].Sequence -eq 24) "$($case.Source): final local ID 24 remains selectable"
            $StartAtSequence=25; $EndAtSequence=48
            Assert-Sequence (@(Select-WorkbookEntries -Entries $expanded).Count -eq 23) "$($case.Source): Unit2 full range respects its unchanged offset"
            $invalid=Import-PowerShellDataFile -LiteralPath $case.Manifest
            $invalid.Workbooks += @{Sequence=13;Folder='2. Calculations';FileName='MutliRoleCheck.xlsx'}
            Assert-SequenceThrows { Resolve-PerUnitWorkbookEntries $invalid } "$($case.Source): restoring a legacy MultiRole entry is rejected"
            $invalid=Import-PowerShellDataFile -LiteralPath $case.Manifest
            @($invalid.Workbooks | Where-Object Sequence -eq 14)[0].Sequence=13
            Assert-SequenceThrows { Resolve-PerUnitWorkbookEntries $invalid } "$($case.Source): ID 13 is rejected even when the file count is valid"
        }
    }
    $functions=Get-SequenceFunctions 'CLIENT/DATExx-Whiddon/runner/Invoke-RefreshRunner.ps1' @('Get-DiscoveredUnits','Get-UnitSequenceEntries','Get-OrgSequenceEntries','Get-GlobalSequenceEntries','Resolve-GlobalSelection')
    & {
        . ([scriptblock]::Create($functions))
        $unitSequencePath=$rootManifest
        $global=@(Get-GlobalSequenceEntries)
        Assert-Sequence ($global.Count -eq 47) 'Date runner has 46 Unit files plus the synthetic organisation entry'
        $unitEntries=@($global | Where-Object Stage -eq 'Units')
        Assert-Sequence (($unitEntries.GlobalSequence -join ',') -eq ($expectedGlobal -join ',')) 'Date runner global IDs agree with execution expansion'
        Assert-Sequence ($unitEntries[-1].GlobalSequence -eq 48 -and $unitEntries[-1].FileName -eq 'Effort.xlsx') 'Date runner last Unit Effort keeps ID 48'
        Assert-SequenceThrows { Resolve-GlobalSelection -StartSequence 13 } 'Date runner rejects removed start ID 13'
        Assert-SequenceThrows { Resolve-GlobalSelection -StartSequence 12 -EndSequence 13 } 'Date runner rejects removed end ID 13'
        $selected=Resolve-GlobalSelection -StartSequence 12 -EndSequence 14
        Assert-Sequence (($selected.SelectedEntries.GlobalSequence -join ',') -eq '12,14') 'Date runner range can span ID 13'
        Assert-Sequence ((Resolve-GlobalSelection -StartSequence 12 -EndSequence 24).SelectedEntries[-1].GlobalSequence -eq 24) 'Date runner preserves explicit end ID 24'
        Assert-Sequence ((Resolve-GlobalSelection).SelectedEntries[-1].GlobalSequence -eq 49) 'Date runner default end uses maximum ID rather than file count'
        Assert-Sequence ((Resolve-GlobalSelection -StartSequence 25 -EndSequence 48).SelectedEntries.Count -eq 23) 'Date runner Unit2 range retains 23 executable files'
        $withOrg=Resolve-GlobalSelection
        Assert-Sequence ($withOrg.OrgGlobalStart -eq 49 -and $withOrg.SelectedEntries[-1].GlobalSequence -eq 49) 'Organisation stage starts after preserved Unit ID span'
        Assert-Sequence ($withOrg.Total -eq 47) 'File count remains distinct from maximum global sequence ID'
        Assert-Sequence ((Resolve-GlobalSelection -StartSequence 49 -EndSequence 49).SelectedEntries[0].GlobalSequence -eq 49) 'The last global ID remains selectable despite a lower executable file count'
    }
    $functions=Get-SequenceFunctions 'CLIENT/DATExx-Whiddon/UNITS/runner/Invoke-AllUnitsRefresh.ps1' @('Get-DiscoveredUnits','Get-WorkbookDisplayName','Show-Sequence')
    & {
        . ([scriptblock]::Create($functions))
        $runRoot=$unitsRoot; $sequencePath=$unitsManifest
        $display=(& { Show-Sequence } 6>&1 | Out-String)
        $displayed=@([regex]::Matches($display,'(?m)^(?<global>\d{3}) \[(?<unit>Unit\d+) #(?<local>\d{2})\]'))
        Assert-Sequence ($displayed.Count -eq 46) 'All-units menu shows exactly the executable files'
        Assert-Sequence ((@($displayed | ForEach-Object {[int]$_.Groups['global'].Value}) -join ',') -eq ($expectedGlobal -join ',')) 'Menu global IDs agree with both workflow executors'
        Assert-Sequence ($display -notmatch 'MutliRoleCheck') 'Menu does not offer the removed workbook'
        Assert-Sequence ($display -match '(?m)^025 \[Unit2 #01\]' -and $display -match '(?m)^048 \[Unit2 #24\]') 'Menu retains Unit2 start and Effort IDs'
    }
    $afterExcel=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Assert-Sequence (-not @($afterExcel | Where-Object {$_ -notin $beforeExcel}).Count) 'Sequence checks start no Excel process'
    Write-Host "PASS: $script:assertions runner sequence assertions; no Excel or workbook-content reads."
} finally {
    $resolvedTest=[IO.Path]::GetFullPath($testRoot)
    if (-not $resolvedTest.StartsWith($temporaryRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture cleanup escaped its temporary root.' }
    if (Test-Path -LiteralPath $resolvedTest) { Remove-Item -LiteralPath $resolvedTest -Recurse -Force }
}
