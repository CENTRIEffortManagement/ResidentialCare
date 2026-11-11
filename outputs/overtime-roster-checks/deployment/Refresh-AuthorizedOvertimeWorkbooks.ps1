[CmdletBinding()]
param([string] $RunFolder='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if(-not $RunFolder){$RunFolder=(Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim()}
$runRoot=[IO.Path]::GetFullPath((Join-Path $repoRoot $RunFolder))
if(-not $runRoot.StartsWith(($repoRoot.TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence directory leaves the repository.'}
$dateRoot=Join-Path $repoRoot 'CLIENT/DATExx-Whiddon'
$worker=Join-Path $dateRoot 'runner/src/Invoke-AllUnitsExcelWorkbookRefresh.ps1'
. (Join-Path $dateRoot 'runner/src/RefreshRunGate.ps1')
$lease=$null
$rows=[Collections.Generic.List[object]]::new()
$selection=@(
    @{Id='Core';Workbook='CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx'},
    @{Id='Shifts';Workbook='CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx'},
    @{Id='Allocation';Workbook='CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx'}
)
function Get-InputStamp([string] $Path){$item=Get-Item -LiteralPath $Path;return "$($item.Length):$($item.LastWriteTimeUtc.Ticks)"}
function Get-Inputs([string] $Id){
    $paths=[Collections.Generic.List[string]]::new()
    $paths.Add((Join-Path $env:PUBLIC 'Public Scripts/CentriSyncPaths.xlsx'))
    if($Id -eq 'Core'){
        foreach($unit in @('BD','TE','JH-RY')){
            foreach($suffix in @('1. Input/Allocation/Published Roster.xlsx',"1. Input/Worker's Report.xlsx",'2. Calculations/Settings Data.xlsx')){
                $paths.Add((Join-Path $dateRoot "UNITS/$unit/$suffix"))
            }
        }
    }elseif($Id -eq 'Shifts'){
        foreach($relative in @('2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx','UNITS/BD/2. Calculations/Intervals.xlsx','UNITS/BD/2. Calculations/Settings Data.xlsx','UNITS/BD/1. Input/1-AllocationExtracted.xlsx')){$paths.Add((Join-Path $dateRoot $relative))}
    }else{
        foreach($relative in @('UNITS/BD/2. Calculations/Shifts.xlsx','UNITS/BD/2. Calculations/Settings Data.xlsx')){$paths.Add((Join-Path $dateRoot $relative))}
    }
    return $paths.ToArray()
}
try{
    # Hold the existing date-level exclusive lease for this exact three-file activation.
    # No production catalogue or sequence is edited and no upstream jobs are added.
    $lease=Enter-RefreshRunGate -DateRoot $dateRoot -Exclusive
    foreach($id in @('Core','Shifts')){
        $receipt=Get-Content -LiteralPath (Join-Path $runRoot "$id-sync-receipt.json") -Raw | ConvertFrom-Json
        if(-not $receipt.ReExtractionMatches -or -not $receipt.Saved -or -not $receipt.Closed){throw "Synchronization is not verified: $id"}
        if((Get-FileHash -LiteralPath (Join-Path $repoRoot $receipt.Source)).Hash -ne $receipt.SourceSHA256){throw "Approved source changed: $id"}
        if((Get-FileHash -LiteralPath (Join-Path $repoRoot $receipt.Workbook)).Hash -ne $receipt.WorkbookSHA256After){throw "Synchronized workbook changed before activation: $id"}
    }
    foreach($item in $selection){
        $target=Join-Path $repoRoot $item.Workbook
        if($item.Id -in @('Core','Shifts')){
            $sync=Get-Content -LiteralPath (Join-Path $runRoot "$($item.Id)-sync-receipt.json") -Raw | ConvertFrom-Json
            if((Get-FileHash -LiteralPath $target).Hash -ne $sync.WorkbookSHA256After){throw "Synchronized target changed before dispatch: $($item.Id)"}
            if((Get-FileHash -LiteralPath (Join-Path $repoRoot $sync.Source)).Hash -ne $sync.SourceSHA256){throw "Approved source changed before dispatch: $($item.Id)"}
        }
        if($item.Id -in @('Shifts','Allocation')){
            $producer=if($item.Id -eq 'Shifts'){'Core'}else{'Shifts'}
            $producerReceipt=Get-Content -LiteralPath (Join-Path $runRoot "$producer-refresh-receipt.json") -Raw | ConvertFrom-Json
            if($producerReceipt.Status -ne 'Passed' -or (Get-FileHash -LiteralPath (Join-Path $repoRoot $producerReceipt.Workbook)).Hash -ne $producerReceipt.WorkbookSHA256After){throw "Selected producer is not current and verified: $producer"}
        }
        $stream=[IO.File]::Open($target,'Open','ReadWrite','None');$stream.Dispose()
        $inputMap=[ordered]@{}
        foreach($inputPath in @(Get-Inputs $item.Id)){
            if(-not(Test-Path -LiteralPath $inputPath -PathType Leaf)){throw "Missing declared input: $inputPath"}
            $inputMap[[IO.Path]::GetFullPath($inputPath)]=Get-InputStamp $inputPath
        }
        $snapshotPath=Join-Path $runRoot "$($item.Id)-inputs.json"
        [ordered]@{SchemaVersion=1;Target=[IO.Path]::GetFullPath($target);Inputs=$inputMap} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $snapshotPath -Encoding utf8
        $before=(Get-FileHash -LiteralPath $target).Hash
        if($item.Id -eq 'Allocation'){
            $backup=Join-Path $runRoot 'backups/Allocation.before.xlsx'
            if(-not(Test-Path -LiteralPath $backup)){Copy-Item -LiteralPath $target -Destination $backup}
            if((Get-FileHash -LiteralPath $backup).Hash -ne $before){throw 'Allocation refresh backup does not match the target.'}
        }
        $started=[DateTime]::UtcNow
        $log=Join-Path $runRoot "$($item.Id)-refresh.log"
        $status=Join-Path $runRoot "$($item.Id)-refresh.status"
        Write-Host "Refreshing authorized target $($item.Id): $($item.Workbook)"
        & (Get-Command pwsh).Source -NoProfile -ExecutionPolicy Bypass -File $worker -WorkbookPath $target -WorkbookName "Overtime activation/$($item.Id)" `
            -VisibleOverride false -TimeoutMinutesOverride 30 -LogPath $log -StatusPath $status -StopRequestPath (Join-Path $runRoot 'stop.txt') -InputSnapshotPath $snapshotPath
        $code=$LASTEXITCODE
        if($code -ne 0){
            $rows.Add([ordered]@{Target=$item.Id;Workbook=$item.Workbook;Status='Failed';ExitCode=$code;Log=($RunFolder+"/$($item.Id)-refresh.log")})
            throw "Targeted refresh failed at $($item.Id); downstream targets were not started."
        }
        $stream=[IO.File]::Open($target,'Open','ReadWrite','None');$stream.Dispose()
        $row=[ordered]@{Target=$item.Id;Workbook=$item.Workbook;Status='Passed';ExitCode=0;Saved=$true;Closed=$true;
            WorkbookSHA256Before=$before;WorkbookSHA256After=(Get-FileHash -LiteralPath $target).Hash;StartedAtUTC=$started.ToString('o');FinishedAtUTC=[DateTime]::UtcNow.ToString('o');
            Log=($RunFolder+"/$($item.Id)-refresh.log");Inputs=($RunFolder+"/$($item.Id)-inputs.json");
            UnselectedUpstreamWorkbooksRefreshedThisRun=$false;
            SelectedUpstreamProducer=$(if($item.Id -eq 'Shifts'){'Core'}elseif($item.Id -eq 'Allocation'){'Shifts'}else{$null});
            SelectedUpstreamProducerRefreshedThisRun=($item.Id -ne 'Core')}
        $rows.Add($row)
        $row | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $runRoot "$($item.Id)-refresh-receipt.json") -Encoding utf8
        $rows.ToArray() | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $runRoot 'refresh-summary.json') -Encoding utf8
    }
    Write-Host 'All three authorized workbooks refreshed, saved and closed in dependency order.'
}catch{
    [ordered]@{Status='Failed';Message=$_.Exception.Message;CompletedOrAttempted=$rows.ToArray();FailedAtUTC=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runRoot 'refresh-failure.json') -Encoding utf8
    throw
}finally{if($null -ne $lease){$lease.Dispose()}}
