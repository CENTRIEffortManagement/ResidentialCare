# CLI regressions for read-only output checks. Only disposable .data fixtures are observed.
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-output-cli-' + [guid]::NewGuid().ToString('N'))
$fixtureRepo = Join-Path $testRoot 'repo'
$fixtureDate = Join-Path $fixtureRepo 'CLIENT/DATExx-Whiddon'
$fixtureRunner = Join-Path $fixtureDate 'runner'
$logRoot = Join-Path $fixtureDate 'RunLogs/BatchRefresh'
$runId = 'output-cli-fixture'
$runDirectory = Join-Path $logRoot $runId
$beforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$script:checks = 0
$script:invocations = 0

function Assert-OutputCli {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}

function Get-FixtureSnapshot {
    $inventory = [Collections.Generic.List[object]]::new()
    foreach ($item in @(Get-ChildItem -LiteralPath $fixtureRepo -Recurse -Force | Sort-Object FullName)) {
        $relative = [IO.Path]::GetRelativePath($fixtureRepo,$item.FullName).Replace('\','/')
        if ($item.PSIsContainer) { $inventory.Add([ordered]@{Path=$relative;Kind='Directory'}) }
        else {
            $inventory.Add([ordered]@{Path=$relative;Kind='File';Length=$item.Length;SavedTicks=$item.LastWriteTimeUtc.Ticks;
                SHA256=(Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash})
        }
    }
    return ConvertTo-Json -InputObject @($inventory) -Depth 5 -Compress
}

function Invoke-OutputCli {
    param([string[]] $Arguments)
    $snapshot = Get-FixtureSnapshot
    $info = [Diagnostics.ProcessStartInfo]::new((Get-Command pwsh -ErrorAction Stop).Source)
    $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WindowStyle='Hidden'
    $info.WorkingDirectory=$fixtureRepo
    $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true; $info.RedirectStandardInput=$true
    foreach ($argument in @('-NoProfile','-File',(Join-Path $fixtureRunner 'Invoke-BatchRefreshRunner.ps1')) + $Arguments) {
        $info.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::Start($info)
    try {
        $process.StandardInput.Close()
        $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
        Assert-OutputCli ($process.WaitForExit(15000)) 'read-only CLI returned without prompting or hanging'
        $result = [pscustomobject]@{ExitCode=$process.ExitCode;Output=$stdout.GetAwaiter().GetResult();ErrorText=$stderr.GetAwaiter().GetResult()}
        Assert-OutputCli ((Get-FixtureSnapshot) -ceq $snapshot) 'CLI preserved all ledger, output, report and runner bytes, timestamps and file/directory inventory'
        Assert-OutputCli (@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object Id -notin $beforeExcel).Count -eq 0) 'CLI started no Excel process'
        $script:invocations++
        return $result
    }
    finally {
        if (-not $process.HasExited) { $process.Kill($true); $process.WaitForExit() }
        $process.Dispose()
    }
}

function Set-FixtureFile {
    param([string] $Path,[datetime] $SavedUtc,[string] $Text)
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path,$Text)
    [IO.File]::SetLastWriteTimeUtc($Path,$SavedUtc)
}

function Save-FixtureRun {
    [void][IO.Directory]::CreateDirectory($runDirectory)
    Write-BatchJson (Join-Path $runDirectory 'plan.json') $script:plan
    Write-BatchJson (Join-Path $runDirectory 'state.json') $script:state
    Write-BatchJson (Join-Path $logRoot 'current.json') @{RunId=$runId}
}

try {
    [void][IO.Directory]::CreateDirectory($fixtureRunner)
    $sourceRunner = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner'
    foreach ($file in @(Get-ChildItem -LiteralPath $sourceRunner -Recurse -File | Where-Object Extension -in @('.ps1','.psd1'))) {
        $target = Join-Path $fixtureRunner ([IO.Path]::GetRelativePath($sourceRunner,$file.FullName))
        [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
        Copy-Item -LiteralPath $file.FullName -Destination $target
    }
    $project = @{batchRunner=@{dateRoot='CLIENT/DATExx-Whiddon';logFolder='CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh';reportFolder='docs'}}
    [IO.File]::WriteAllText((Join-Path $fixtureRepo 'pq.project.json'),($project | ConvertTo-Json -Depth 5))
    # A regression reaching dispatch is visible even when it fails before an Excel worker starts.
    $forbidden = "throw 'Excel execution is forbidden in output-status CLI tests.'"
    foreach ($name in @('Invoke-AllUnitsExcelWorkbookRefresh.ps1','Invoke-OrgExcelWorkbookRefresh.ps1','Invoke-BatchWorkbook.ps1')) {
        [IO.File]::WriteAllText((Join-Path $fixtureRunner "src/$name"),$forbidden)
    }
    [IO.File]::AppendAllText((Join-Path $fixtureRunner 'src/BatchScheduling.ps1'), @'

function Invoke-BatchSchedule {
    param($Plan,$State,$RunDirectory,$MaxParallelBatches,$Visible,$TimeoutMinutes)
    [IO.File]::WriteAllText((Join-Path $Plan.DateRoot 'FORBIDDEN-DISPATCH.txt'),'Execution reached a read-only CLI fixture.')
    throw 'Dispatch is forbidden in output-status CLI tests.'
}
'@)
    . (Join-Path $fixtureRunner 'src/BatchPlanning.ps1')
    . (Join-Path $fixtureRunner 'src/BatchScheduling.ps1')
    . (Join-Path $fixtureRunner 'src/BatchReporting.ps1')

    $result = Invoke-OutputCli @('-CheckOutputs')
    Assert-OutputCli ($result.ExitCode -eq 1 -and $result.Output -match 'No batch run has been recorded|without a recorded run') 'missing current.json returns exit 1 with an explicit unverified result'

    $time = [datetime]::UtcNow.AddHours(-4)
    $inputPath = Join-Path $fixtureDate 'Capacity.data'
    $outputPath = Join-Path $fixtureDate 'Effort.data'
    Set-FixtureFile $inputPath $time.AddMinutes(-10) 'synthetic saved capacity'
    Set-FixtureFile $outputPath $time.AddMinutes(10) 'synthetic saved effort'
    $job = [pscustomobject]@{Id='BD/Effort';Unit='BD';Role='';Batch='U4';BatchKey='BD/U4';RelativePath='Effort.data';Path=$outputPath;Reads=@($inputPath);Dependencies=@();Exclusive=$false}
    $script:plan = [pscustomobject]@{SchemaVersion=1;Fingerprint='output-cli-metadata-fixture';SequenceProfile='Current';DateRoot=$fixtureDate;Jobs=@($job)}
    $script:state = New-BatchState $script:plan $runId
    $script:state.Status='Completed'; $script:state.ExitCode=0
    $script:state.Started=$time.ToString('o'); $script:state.Updated=$time.AddMinutes(60).ToString('o')
    $record=$script:state.Jobs[$job.Id]
    $record.Status='Completed'; $record.ExitCode=0; $record.Attempt=1
    $record.Started=$time.ToString('o'); $record.Finished=$time.AddMinutes(20).ToString('o')
    $record.OutputStamp=Get-BatchFileStamp $outputPath
    $record.Inputs[$inputPath]=Get-BatchFileStamp $inputPath
    Save-FixtureRun
    $result = Invoke-OutputCli @('-CheckOutputs')
    Assert-OutputCli ($result.ExitCode -eq 0 -and $result.Output -match 'All 1 outputs match their technical run receipts' -and [string]::IsNullOrWhiteSpace($result.ErrorText)) 'matching successful receipts, fingerprints and declared input metadata return exit 0'

    $producerPath=Join-Path $fixtureDate 'MultiRole.data'
    Set-FixtureFile $producerPath $time.AddMinutes(-5) 'synthetic old multirole output'
    $producer=[pscustomobject]@{Id='BD/MultiRole';Unit='BD';Role='';Batch='A3';BatchKey='BD/A3';RelativePath='MultiRole.data';Path=$producerPath;Reads=@();Dependencies=@();Exclusive=$false}
    $script:plan.Jobs=@($producer,$job)
    $job.Dependencies=@($producer.Id); $job.Reads=@($inputPath,$producerPath)
    $record.Inputs[$producerPath]=Get-BatchFileStamp $producerPath
    $script:state.Jobs[$producer.Id]=@{Status='Failed';Message="Expression.Error: The key didn't match any rows in the table.";Attempt=1;Started=$time.ToString('o');Finished=$time.AddMinutes(5).ToString('o');ExitCode=1;OutputStamp=(Get-BatchFileStamp $producerPath);Inputs=@{};NeedsInspection=$false;Directory=''}
    $script:state.Status='PartialOrFailed'; $script:state.ExitCode=1
    $record.Status='Blocked'; $record.ExitCode=$null; $record.Started=''; $record.Finished=''; $record.OutputStamp=''
    $record.Message='Prerequisite BD/MultiRole did not complete.'
    Set-FixtureFile $inputPath $time.AddMinutes(5) 'synthetic newly saved capacity'
    Set-FixtureFile $outputPath $time.AddMinutes(-20) 'synthetic old effort'
    Save-FixtureRun
    $result = Invoke-OutputCli @('-CheckOutputs')
    Assert-OutputCli ($result.ExitCode -eq 1 -and $result.Output -match 'BD/Effort \| Blocked' -and $result.Output -match 'BD/Effort \| NotUpdated') 'blocked Effort remains blocked and not updated with exit 1'
    Assert-OutputCli ($result.Output.Contains("The key didn't match any rows in the table.") -and $result.Output -match 'Root cause: BD/MultiRole') 'blocked output preserves its original upstream expression error'

    Set-FixtureFile $outputPath $time.AddMinutes(80) 'synthetic later manually saved effort'
    $result = Invoke-OutputCli @('-CheckOutputs')
    Assert-OutputCli ($result.ExitCode -eq 1 -and $result.Output -match 'BD/Effort \| Blocked' -and $result.Output -match 'BD/Effort \| ChangedAfterRun' -and $result.Output -match 'Historical run status remains Blocked') 'manual save becomes ChangedAfterRun without rewriting the blocked receipt'

    $result = Invoke-OutputCli @('-ShowStatus')
    Assert-OutputCli ($result.ExitCode -eq 0 -and $result.Output -match 'ChangedAfterRun' -and $result.Output -match 'BD/Effort \| Blocked') 'ShowStatus retains its informational exit 0 for blocked and changed outputs'

    $result = Invoke-OutputCli @('-CheckOutputs','-RefreshSelected')
    Assert-OutputCli ($result.ExitCode -eq 1 -and $result.ErrorText -match 'CheckOutputs is read-only; do not combine') 'mixed read-only and execution options fail before dispatch'
    Assert-OutputCli (-not (Test-Path -LiteralPath (Join-Path $fixtureDate 'FORBIDDEN-DISPATCH.txt'))) 'no execution was dispatched'
    Assert-OutputCli (@(Get-ChildItem -LiteralPath $fixtureRepo -Recurse -File | Where-Object Extension -eq '.xlsx').Count -eq 0) 'fixture contains no Excel workbooks'
    Write-Host "PASS: $script:checks output-status CLI checks across $script:invocations commands; full fixture inventories and bytes unchanged by each command; no Excel process started."
}
finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([char[]]'\/')+[IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-output-cli-*') { throw 'Unsafe output-status CLI fixture cleanup target.' }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
