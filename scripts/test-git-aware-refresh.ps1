param([switch] $LiveProcessChecks)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1')
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchScheduling.ps1')
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('rc-git-access-' + [guid]::NewGuid().ToString('N'))
[void] [IO.Directory]::CreateDirectory($testRoot)
$script:checks = 0
$script:stopRequested = $false
$script:accessMessages = [Collections.Generic.List[string]]::new()
function Assert-GitTest {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++
}
function Assert-GitError {
    param([scriptblock] $Action, [string] $Pattern)
    $message = ''; try { & $Action | Out-Null } catch { $message = $_.Exception.Message }
    Assert-GitTest ($message -like $Pattern) "Expected '$Pattern'; got '$message'"
}
function Assert-NotStopNowRequested { if ($script:stopRequested) { throw 'Stop requested.' } }
function Write-Log { param($Message, $Level) $script:accessMessages.Add($Message) }
function Write-CurrentStatus { param($State, $Message) }
function Get-TestProcessUser {
    param($Process, [string] $Name = 'Git for Windows')
    $ticks = $Process.StartTime.ToUniversalTime().ToFileTimeUtc()
    $bytes = [BitConverter]::GetBytes($ticks)
    [pscustomobject]@{ AppName=$Name; Process=[pscustomobject]@{Id=$Process.Id;Started=[pscustomobject]@{
        dwLowDateTime=[BitConverter]::ToInt32($bytes,0); dwHighDateTime=[BitConverter]::ToInt32($bytes,4)
    }} }
}
$script:RefreshState = @{phase='WaitingToSave';detail=$null;saved=$false;updatedAt=''}
$WorkerStatePath = Join-Path $testRoot 'receipt.json'
$gitPath = @(Get-RefreshGitPaths)[0]
$originalUsers = ${function:Get-RefreshFileUsers}
try {
    $validCommands = @(
        'git.exe status --porcelain=v1 -z --untracked-files=all',
        'git.exe --no-optional-locks -C "repo with spaces" -c core.quotepath=false status --porcelain=2 --branch',
        'git.exe diff --no-ext-diff --no-textconv --name-status --cached HEAD -- target.data',
        'git.exe diff-files --no-ext-diff --no-textconv --raw -z',
        'git.exe diff-index --no-ext-diff --no-textconv --quiet HEAD -- target.data',
        'git.exe diff-tree --no-ext-diff --no-textconv --no-commit-id -r HEAD',
        'git.exe ls-files --stage -z',
        'git.exe hash-object --no-filters --stdin-paths',
        'git.exe cat-file --batch-check'
    )
    foreach ($command in $validCommands) {
        Assert-GitTest (Test-RefreshReadOnlyGitCommand @(Split-RefreshCommandLine $command)) "allow known read-only command: $command"
    }
    foreach ($command in @(
        'git.exe checkout main', 'git.exe restore target.data', 'git.exe reset --hard',
        'git.exe merge branch', 'git.exe pull', 'git.exe switch main', 'git.exe add .',
        'git.exe clean -fd', 'git.exe commit -am change', 'git.exe hash-object --no-filters -w target.data',
        'git.exe hash-object target.data', 'git.exe diff --name-only',
        'git.exe diff --no-ext-diff --no-textconv --ext-diff',
        'git.exe diff --no-ext-diff --no-textconv --output=target.data',
        'git.exe diff --no-ext-diff --no-textconv --textconv',
        'git.exe --exec-path=custom status', 'git.exe -c core.fsmonitor=custom status',
        'git.exe -c alias.reader=!evil reader', 'git.exe --config-env=core.fsmonitor=VALUE status',
        'git.exe STATUS', 'git.exe cat-file --filters HEAD:target.data', 'git.exe status --help'
    )) { Assert-GitTest (-not (Test-RefreshReadOnlyGitCommand @(Split-RefreshCommandLine $command))) "block writer/unknown configuration: $command" }
    $tokens = @(Split-RefreshCommandLine '"C:\Program Files\Git\cmd\git.exe" -C "repo with spaces" status -- "a file.data"')
    Assert-GitTest ($tokens.Count -eq 6 -and $tokens[0] -eq 'C:\Program Files\Git\cmd\git.exe' -and $tokens[2] -eq 'repo with spaces') 'Windows quoting is parsed as arguments, never executed'
    $redacted = Protect-RefreshCommandDescription @('git.exe','-c','http.extraheader=Authorization: secret','status')
    Assert-GitTest ($redacted -notmatch 'secret' -and $redacted -match 'redacted') 'diagnostics redact injected configuration values'

    # Mock only CIM metadata for deterministic process classification. The PID and
    # creation time come from this real PowerShell process; no workbook is used.
    $self = Get-Process -Id $PID
    $script:testUser = Get-TestProcessUser $self
    $script:testMetadata = [pscustomobject]@{
        ExecutablePath=$gitPath; CommandLine=('"' + $gitPath + '" status --porcelain=v1')
        ParentProcessId=1; CreationDate=$self.StartTime
    }
    function Get-CimInstance {
        param($ClassName, $Filter, $ErrorAction)
        if ($Filter -eq "ProcessId = $PID") { return $script:testMetadata }
        return [pscustomobject]@{Name='fixture-editor.exe';CreationDate=$script:testMetadata.CreationDate.AddSeconds(-1)}
    }
    $decision = Get-RefreshFileUserDecision $script:testUser
    Assert-GitTest ($decision.Allowed -and $decision.ParentApplication -eq 'fixture-editor.exe') "verified Git decision and parent evidence: $($decision.Reason)"
    $script:testMetadata.ExecutablePath = Join-Path $testRoot 'git.exe'
    Assert-GitTest (-not (Get-RefreshFileUserDecision $script:testUser).Allowed) 'a lookalike executable is not trusted by filename'
    $script:testMetadata.ExecutablePath = $gitPath
    $script:testMetadata.CommandLine = ('"' + $gitPath + '" checkout main')
    Assert-GitTest (-not (Get-RefreshFileUserDecision $script:testUser).Allowed) 'verified Git writer is still blocked'
    $script:testMetadata.CommandLine = ''
    Assert-GitTest (-not (Get-RefreshFileUserDecision $script:testUser).Allowed) 'missing command line is not treated as read-only'
    $script:testMetadata.CommandLine = ('"' + $gitPath + '" status --porcelain=v1')
    $script:testUser.Process.Started.dwLowDateTime = 0
    Assert-GitTest (-not (Get-RefreshFileUserDecision $script:testUser).Allowed) 'PID reuse/mismatched start time cannot inherit approval'
    $script:testUser = Get-TestProcessUser $self
    function Get-RefreshFileUsers { param($Path) $script:testUser }
    $target = Join-Path $testRoot 'target.data'; [IO.File]::WriteAllText($target, 'original')
    Assert-NoExternalFileUsers $target
    Assert-GitTest (($script:accessMessages -join ' ') -match '"Allowed":true') 'allowed reader is diagnostic, not an immediate failure'
    $actualLock = [IO.File]::Open($target, 'Open', 'Read', [IO.FileShare]::Read)
    $osRejected = $false
    try { Assert-BatchFileAccessible $target -OutputFile -InspectFileUsers } catch { $osRejected = $true }
    finally { $actualLock.Dispose() }
    Assert-GitTest $osRejected 'allowlisted Git does not bypass an actual Windows sharing conflict'
    $script:testMetadata.CommandLine = ('"' + $gitPath + '" restore target.data')
    Assert-GitError { Assert-NoExternalFileUsers $target -WaitSeconds .02 -PollMilliseconds 5 } '*External-user policy blocked*not an Excel save error*'
    $script:stopRequested = $true
    Assert-GitError { Wait-RefreshFileAccess $target } '*Stop requested*'
    $script:stopRequested = $false
    function Get-CimInstance { param($ClassName, $Filter, $ErrorAction) throw 'Process inspection unavailable.' }
    Assert-GitTest (-not (Get-RefreshFileUserDecision $script:testUser).Allowed) 'unavailable inspection fails closed'
    Remove-Item Function:Get-CimInstance
    ${function:Get-RefreshFileUsers} = $originalUsers

    # Whole-refresh metadata protection, including changes BEFORE the access wait.
    $inputPath = Join-Path $testRoot 'input.data'; [IO.File]::WriteAllText($inputPath, 'input original')
    $inputSnapshotPath = Join-Path $testRoot 'input-snapshot.json'
    Write-RefreshJson $inputSnapshotPath @{SchemaVersion=1;Target=$target;Inputs=@{$inputPath=(Get-RefreshTargetStamp $inputPath)}}
    $snapshot = New-RefreshFileSnapshot $target $inputSnapshotPath
    Assert-RefreshFileSnapshot $snapshot
    [IO.File]::WriteAllText($target, 'changed by another writer before save')
    Assert-GitError { Assert-RefreshFileSnapshot $snapshot } '*Target workbook changed since pre-open*'
    Assert-RefreshFileSnapshot $snapshot -AfterSave
    [IO.File]::WriteAllText($inputPath, 'input edited while refreshing')
    Assert-GitError { Assert-RefreshFileSnapshot $snapshot -AfterSave } '*Declared input changed*'
    Assert-GitError { New-RefreshFileSnapshot $target $inputSnapshotPath } '*Declared input changed*'
    $missingInput = Join-Path $testRoot 'missing.data'
    Write-RefreshJson $inputSnapshotPath @{SchemaVersion=1;Target=$target;Inputs=@{$missingInput='missing'}}
    Assert-GitError { New-RefreshFileSnapshot $target $inputSnapshotPath } '*does not exist*'
    Write-RefreshJson $inputSnapshotPath @{SchemaVersion=1;Target=$inputPath;Inputs=@{}}
    Assert-GitError { New-RefreshFileSnapshot $target $inputSnapshotPath } '*does not belong*'

    foreach ($path in @('CLIENT/DATExx-Whiddon/runner/src/Invoke-AllUnitsExcelWorkbookRefresh.ps1','CLIENT/DATExx-Whiddon/UNITS/runner/src/Invoke-ExcelWorkbookRefresh.ps1')) {
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot $path))
        $capture = $source.IndexOf('$fileSnapshot = New-RefreshFileSnapshot')
        $open = $source.IndexOf('$workbook = $excel.Workbooks.Open')
        $save = $source.IndexOf('$workbook.Save()')
        $check = $source.IndexOf('Assert-RefreshFileSnapshot $fileSnapshot', $source.IndexOf("Set-RefreshPhase 'WaitingToSave'"))
        Assert-GitTest ($capture -gt 0 -and $capture -lt $open -and $check -gt $open -and $check -lt $save) "$path captures before opening and checks before save"
        Assert-GitTest ($source.IndexOf('Assert-RefreshFileSnapshot $fileSnapshot -AfterSave') -gt $save) "$path rechecks inputs after save"
    }

    if ($LiveProcessChecks) {
        # Git only operates in a disposable repository of .data fixtures. No
        # real workbook paths/content, user Git config, filters or hooks are used.
        $fixture = Join-Path $testRoot 'repository'; [void] [IO.Directory]::CreateDirectory($fixture)
        $script:gitReader = $null
        $savedUsers = ${function:Get-RefreshFileUsers}
        try {
            & $gitPath -c init.defaultBranch=main init --quiet $fixture
            if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize disposable Git repository.' }
            $info = [Diagnostics.ProcessStartInfo]::new($gitPath)
            $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WindowStyle='Hidden'
            $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
            $info.Environment['GIT_CONFIG_NOSYSTEM']='1'; $info.Environment['GIT_CONFIG_GLOBAL']='NUL'
            foreach ($arg in @('--no-pager','-C',$fixture,'hash-object','--no-filters','--stdin-paths')) { $info.ArgumentList.Add($arg) }
            $script:gitReader = [Diagnostics.Process]::Start($info)
            $stdout = $script:gitReader.StandardOutput.ReadToEndAsync(); $stderr = $script:gitReader.StandardError.ReadToEndAsync()
            $script:liveUser = Get-TestProcessUser $script:gitReader
            $liveDecision = Get-RefreshFileUserDecision $script:liveUser
            Assert-GitTest $liveDecision.Allowed "real installed Git process and command verification: $($liveDecision.Reason)"
            # Simulate repeated RM sightings of this verified live Git process.
            # Its actual read commands run alongside the synthetic scheduler.
            function Get-RefreshFileUsers { param($Path) $script:liveUser }
            $jobs = @(1..9 | ForEach-Object {
                $path = Join-Path $fixture "job$_.data"; [IO.File]::WriteAllText($path, 'test content')
                [pscustomobject]@{Id="job$_";BatchKey="batch$_";Path=$path;Reads=@();Dependencies=@();Exclusive=$false}
            })
            foreach ($limit in @(1,7)) {
                $directory=Join-Path $testRoot "schedule-$limit"; [void] [IO.Directory]::CreateDirectory($directory)
                $plan=[pscustomobject]@{DateRoot=$fixture;Jobs=$jobs;Fingerprint='fixture'}
                $state=New-BatchState $plan "git-$limit"
                $start = {
                    param($job,$directory,$options)
                    $script:gitReader.StandardInput.WriteLine($job.Path); $script:gitReader.StandardInput.Flush()
                    $stamp=New-RefreshFileSnapshot $job.Path
                    return @{Job=$job;Snapshot=$stamp;Remaining=2}
                }
                $poll = {
                    param($handle)
                    if (--$handle.Remaining -gt 0) { return $null }
                    Assert-NoExternalFileUsers $handle.Job.Path
                    Assert-RefreshFileSnapshot $handle.Snapshot
                    [IO.File]::WriteAllText($handle.Job.Path, 'synthetic saved ' + [guid]::NewGuid().ToString('N'))
                    return @{ExitCode=0;NeedsInspection=$false;Message='Synthetic save complete alongside Git reads.'}
                }
                $code=Invoke-BatchSchedule $plan $state $directory -MaxParallelBatches $limit -StartWorker $start -PollWorker $poll -ValidateJob {param($job) Assert-NoExternalFileUsers $job.Path} -PollMilliseconds 0
                Assert-GitTest ($code -eq 0 -and $state.MaxObserved -eq $limit -and @($state.Jobs.Values | Where-Object Status -eq Completed).Count -eq 9) "all nine jobs complete with live Git reads at concurrency $limit"
            }
            $script:gitReader.StandardInput.Close()
            Assert-GitTest ($script:gitReader.WaitForExit(10000) -and $script:gitReader.ExitCode -eq 0) 'background Git reader exits normally without being terminated'
            $hashes = @($stdout.GetAwaiter().GetResult().Trim() -split "\r?\n")
            Assert-GitTest ($hashes.Count -eq 18 -and @($hashes | Where-Object {$_ -notmatch '^[a-f0-9]{40,64}$'}).Count -eq 0) 'real Git completed all eighteen fixture file reads'
            Assert-GitTest ([string]::IsNullOrWhiteSpace($stderr.GetAwaiter().GetResult())) 'background Git reported no read failure'
        } finally {
            ${function:Get-RefreshFileUsers}=$savedUsers
            if ($null -ne $script:gitReader) {
                if (-not $script:gitReader.HasExited) { $script:gitReader.StandardInput.Close(); if (-not $script:gitReader.WaitForExit(3000)) { $script:gitReader.Kill($true); $script:gitReader.WaitForExit() } }
                $script:gitReader.Dispose()
            }
        }
    }
    Write-Host "PASS: $script:checks Git-aware checks; live process checks: $LiveProcessChecks. No Excel opened."
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if (-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'rc-git-access-*') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
