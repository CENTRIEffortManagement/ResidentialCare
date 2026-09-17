# Inspect existing processes only. Never launch Git against a live workbook,
# suspend Git, alter its index flags, or trust an application display name alone.
function Split-RefreshCommandLine {
    param([string] $CommandLine)
    if ([string]::IsNullOrWhiteSpace($CommandLine)) { return @() }
    if (-not ('ResidentialCare.CommandArguments' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ResidentialCare {
 public static class CommandArguments {
  [DllImport("shell32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CommandLineToArgvW(string command, out int count);
  [DllImport("kernel32.dll")] static extern IntPtr LocalFree(IntPtr memory);
  public static string[] Split(string command) {
   int count; IntPtr memory=CommandLineToArgvW(command,out count);
   if(memory==IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
   try { var args=new string[count]; for(int i=0;i<count;i++) args[i]=Marshal.PtrToStringUni(Marshal.ReadIntPtr(memory,i*IntPtr.Size)); return args; }
   finally { LocalFree(memory); }
  }
 }
}
'@
    }
    return [ResidentialCare.CommandArguments]::Split($CommandLine)
}

function Get-RefreshGitPaths {
    # PATH-resolved installed Git and its standard Windows backend. No matching
    # by filename alone and no trust inferred from an arbitrary sibling folder.
    $paths = [Collections.Generic.List[string]]::new()
    foreach ($git in @(Get-Command git.exe -CommandType Application -All -ErrorAction SilentlyContinue)) {
        $paths.Add([IO.Path]::GetFullPath($git.Source))
        $folder = Split-Path $git.Source -Parent
        if ((Split-Path $folder -Leaf) -eq 'cmd') {
            foreach ($relative in @('mingw64/bin/git.exe', 'mingw32/bin/git.exe')) {
                $backend = Join-Path (Split-Path $folder -Parent) $relative
                if (Test-Path -LiteralPath $backend -PathType Leaf) { $paths.Add([IO.Path]::GetFullPath($backend)) }
            }
        }
    }
    return @($paths | Select-Object -Unique)
}

function Test-RefreshReadOnlyGitCommand {
    param([string[]] $Arguments)
    if ($Arguments.Count -lt 2) { return $false }
    $index = 1
    while ($index -lt $Arguments.Count -and $Arguments[$index].StartsWith('-')) {
        $arg = $Arguments[$index]
        if ($arg -cin @('--no-pager', '--no-optional-locks', '--literal-pathspecs', '--no-replace-objects')) { $index++; continue }
        if ($arg -ceq '-C' -and ($index + 1) -lt $Arguments.Count) { $index += 2; continue }
        if ($arg -ceq '-c' -and ($index + 1) -lt $Arguments.Count -and
            $Arguments[$index + 1] -cmatch '^(core\.quotepath=(true|false)|color\.ui=(always|never|auto)|core\.preloadindex=(true|false))$') { $index += 2; continue }
        # Config injection, aliases, exec-path and config-env are not exempted.
        return $false
    }
    if ($index -ge $Arguments.Count) { return $false }
    $command = $Arguments[$index++]
    if ($command -cnotin @('status', 'diff', 'diff-files', 'diff-index', 'diff-tree', 'ls-files', 'hash-object', 'cat-file')) { return $false }
    $options = [Collections.Generic.List[string]]::new()
    $pathsOnly = $false
    while ($index -lt $Arguments.Count) {
        $arg = $Arguments[$index++]
        if ($pathsOnly) { continue }
        if ($arg -ceq '--') { $pathsOnly = $true; continue }
        if (-not $arg.StartsWith('-') -or $arg -ceq '-') { continue }
        $options.Add($arg)
        $allowed = switch -CaseSensitive ($command) {
            'status' { $arg -cmatch '^(-z|-s|-b|--short|--branch|--porcelain(=v?[12])?|--untracked-files(=(all|no|normal))?|-u(all|no|normal)?|--ignored(=(traditional|matching|no))?|--no-renames|--show-stash|--ahead-behind|--no-ahead-behind)$' }
            { $_ -cin @('diff', 'diff-files', 'diff-index', 'diff-tree') } {
                $arg -cmatch '^(-z|-p|-r|-m|--root|--no-commit-id|--no-ext-diff|--no-textconv|--name-only|--name-status|--raw|--numstat|--stat|--shortstat|--no-renames|--cached|--staged|--quiet|--exit-code|--no-color|--color=never|--no-prefix|--binary|--patch|--full-index|--ignore-submodules(=(all|dirty|untracked|none))?|--diff-filter=[A-Za-z*]+|--unified=\d+)$'
            }
            'ls-files' { $arg -cmatch '^(-z|-s|-v|-t|-c|-d|-m|-o|-i|--stage|--cached|--deleted|--modified|--others|--ignored|--exclude-standard|--full-name|--eol|--debug|--deduplicate)$' }
            'hash-object' { $arg -cin @('--stdin', '--stdin-paths', '--no-filters') }
            'cat-file' { $arg -cin @('--batch', '--batch-check', '--buffer', '-p', '-t', '-s', '-e') }
            default { $false }
        }
        if (-not $allowed) { return $false }
    }
    # diff configuration can run external programs unless BOTH are disabled.
    if ($command -cin @('diff', 'diff-files', 'diff-index', 'diff-tree')) {
        return ('--no-ext-diff' -cin $options -and '--no-textconv' -cin $options)
    }
    # hash-object otherwise applies repository clean filters to working files.
    if ($command -ceq 'hash-object') { return '--no-filters' -cin $options }
    return $true
}

function Protect-RefreshCommandDescription {
    param([string[]] $Arguments)
    $safe = [Collections.Generic.List[string]]::new()
    $redactNext = $false
    foreach ($arg in $Arguments) {
        if ($redactNext) { $safe.Add(($arg -split '=', 2)[0] + '=<redacted>'); $redactNext = $false; continue }
        if ($arg -cin @('-c', '--config-env')) { $safe.Add($arg); $redactNext = $true; continue }
        if ($arg -match '(?i)(password|token|secret|extraheader|authorization)=') { $safe.Add('<redacted argument>'); continue }
        $safe.Add(($arg -replace '(?i)(https?://)[^/@\s]+:[^/@\s]+@', '$1<redacted>@'))
    }
    return ($safe | ConvertTo-Json -Compress)
}

function Get-RefreshFileUserDecision {
    param($User)
    $decision = [ordered]@{
        Pid = [int] $User.Process.Id; Application = [string] $User.AppName
        Started = ''; Executable = ''; Command = ''; ParentPid = $null; ParentApplication = ''
        Allowed = $false; Reason = 'Process identity or command could not be verified.'
    }
    try {
        # A Restart Manager identity includes PID AND creation FILETIME. A
        # missing identity or PID reuse must never inherit an earlier allowance.
        $started = $User.Process.Started
        $low = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int] $started.dwLowDateTime), 0)
        $high = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int] $started.dwHighDateTime), 0)
        $fileTime = ([uint64] $high -shl 32) -bor [uint64] $low
        $row = Get-CimInstance Win32_Process -Filter "ProcessId = $($decision.Pid)" -ErrorAction Stop
        $process = Get-Process -Id $decision.Pid -ErrorAction Stop
        if ([uint64] $process.StartTime.ToUniversalTime().ToFileTimeUtc() -ne $fileTime) { throw 'Process identity changed since file-user detection.' }
        $decision.Started = $process.StartTime.ToUniversalTime().ToString('o')
        $decision.Executable = [string] $row.ExecutablePath
        $arguments = @(Split-RefreshCommandLine ([string] $row.CommandLine))
        $decision.Command = Protect-RefreshCommandDescription $arguments
        $decision.ParentPid = [int] $row.ParentProcessId
        try {
            $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($row.ParentProcessId)" -ErrorAction Stop
            if ($null -ne $parent -and $parent.CreationDate -le $row.CreationDate) { $decision.ParentApplication = [string] $parent.Name }
        } catch { $decision.ParentApplication = 'unavailable' }
        if ([string]::IsNullOrWhiteSpace($decision.Executable) -or $arguments.Count -lt 2) { throw 'Executable or command unavailable.' }
        if ([IO.Path]::GetFullPath($decision.Executable) -notin @(Get-RefreshGitPaths)) { $decision.Reason = 'Not a recognised installed Git executable.' }
        elseif (-not (Test-RefreshReadOnlyGitCommand $arguments)) { $decision.Reason = 'Git command/options are not on the read-only allowlist.' }
        else { $decision.Allowed = $true; $decision.Reason = 'Verified installed Git with an allowlisted read-only working-tree operation.' }
    } catch { $decision.Reason = "Unverified process: $($_.Exception.Message)" }
    return [pscustomobject] $decision
}

function Write-RefreshAccessDecision {
    param($Decision)
    $line = 'FILE USER: ' + ($Decision | ConvertTo-Json -Compress -Depth 5)
    if (Get-Command Write-Log -CommandType Function -ErrorAction SilentlyContinue) { Write-Log $line }
    else { Write-Host $line }
}
