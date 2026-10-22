# Exercise the existing Office supervisor with an owned fake worker; no COM/build.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/reference-worker-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) {throw 'Use a fresh test directory'}
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$worker=Join-Path $root 'worker.ps1'
@'
param($InputPath,$WorkDirectory,[int]$Dpi,$ProgressLog,$StatusPath)
$ErrorActionPreference='Stop'
if ($args.Count) {throw 'Unexpected unbound arguments'}
Add-Content -LiteralPath $ProgressLog -Value 'stage:export'
'fake worker stdout'; [Console]::Error.WriteLine('fake worker stderr')
$mode=[IO.Path]::GetFileNameWithoutExtension($InputPath)
if ($mode -eq 'hang') {Start-Sleep -Seconds 60}
if ($mode -eq 'unavailable') {
    @{Status='unavailable';Stage='activation';Error='Synthetic activation failure'} | ConvertTo-Json | Set-Content $StatusPath -Encoding utf8
    exit 1
}
@{InputPath=$InputPath;WorkDirectory=$WorkDirectory;Dpi=$Dpi;ProgressLog=$ProgressLog;StatusPath=$StatusPath} | ConvertTo-Json | Set-Content (Join-Path $WorkDirectory 'arguments.json') -Encoding utf8
Copy-Item -LiteralPath $InputPath -Destination (Join-Path $WorkDirectory 'reference.pdf')
@{Status='ok';Stage='done';Error=''} | ConvertTo-Json | Set-Content $StatusPath -Encoding utf8
'@ | Set-Content $worker -Encoding utf8
$spacedWorker=Join-Path $root 'worker with spaces.ps1'
Copy-Item $worker $spacedWorker
$checks=@()
$sentinelStart=[Diagnostics.ProcessStartInfo]::new('pwsh')
$sentinelStart.UseShellExecute=$false;$sentinelStart.CreateNoWindow=$true
foreach($arg in @('-NoProfile','-Command','Start-Sleep -Seconds 60')) {$sentinelStart.ArgumentList.Add($arg)}
$sentinel=[Diagnostics.Process]::Start($sentinelStart)
try {
    foreach($id in @('plain','spaced-input','spaced-output','spaced-worker','all-spaced','unavailable','hang')) {
        $inputName=if($id -in @('spaced-input','all-spaced')) {'input ''quoted'' é.pdf'} else {$id+'.pdf'}
        $inputFile=Join-Path $root $inputName
        [IO.File]::WriteAllText($inputFile,'%PDF-Synthetic supervisor probe, not a fidelity reference')
        $output=Join-Path $root $(if($id -in @('spaced-output','all-spaced')) {$id+' output with spaces'} else {$id+'-output'})
        $selected=if($id -in @('spaced-worker','all-spaced')) {$spacedWorker} else {$worker}
        $timeout=if($id -eq 'hang') {2} else {20}
        $errorText=$null
        try { & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFile -OutputDirectory $output -WorkerScript $selected -Dpi 123 -TimeoutSeconds $timeout *> (Join-Path $root ($id+'.log')) }
        catch {$errorText=$_.Exception.Message}
        $status=Get-Content (Join-Path $output 'reference-status.json') -Raw | ConvertFrom-Json
        $expected=if($id -eq 'hang') {'timeout'} elseif($id -eq 'unavailable') {'unavailable'} else {'ok'}
        if($status.Status -ne $expected -or (($expected -eq 'ok') -eq ($null -ne $errorText))) {throw ($id+': unexpected supervisor outcome '+$errorText)}
        if($expected -eq 'ok') {
            $actual=Get-Content (Join-Path $output 'arguments.json') -Raw | ConvertFrom-Json
            if($actual.InputPath -cne $inputFile -or $actual.Dpi -ne 123 -or
                !(Split-Path $actual.WorkDirectory -Parent).Equals($output,[StringComparison]::Ordinal) -or
                [IO.File]::ReadAllText((Join-Path $output 'reference.pdf')) -ne [IO.File]::ReadAllText($inputFile)) {throw ($id+': argument/output identity differs')}
        }
        foreach($name in @('stdout','stderr')) {
            if([IO.File]::ReadAllText((Join-Path $output ('reference-worker-'+$name+'.log'))) -notlike ('*fake worker '+$name+'*')) {throw 'Worker log was lost'}
        }
        if($sentinel.HasExited) {throw 'Supervisor terminated an unrelated owned control process'}
        $checks+=@{Id=$id;Status=$status.Status;Error=$errorText}
    }
}
finally {if(!$sentinel.HasExited) {$sentinel.Kill();$null=$sentinel.WaitForExit(5000)};$sentinel.Dispose()}

# Exercise the actual worker, replacing only COM activation with managed stubs.
# Every case fails before rasterization, so these checks never build or use Office.
$actualWorkerSource=Get-Content (Join-Path $PSScriptRoot 'RenderReferenceWorker.ps1') -Raw
$activationStub=@'
function New-Object {
    param([string] $ComObject)
    if ($ComObject -notin @('PowerPoint.Application','Word.Application')) {throw 'Unexpected COM request'}
    $script:failureStage=[IO.Path]::GetFileNameWithoutExtension($InputPath)
    if ($script:failureStage -eq 'activation') {throw 'Controlled activation failure'}
    $script:probeDocument=[pscustomobject]@{}
    $save={param($Path,$Format)
        if ($script:failureStage -eq 'export') {throw 'Controlled export failure'}
        [IO.File]::WriteAllText($Path,'%PDF-control')
    }
    $script:probeDocument | Add-Member ScriptMethod SaveAs $save
    $script:probeDocument | Add-Member ScriptMethod SaveAs2 $save
    $script:probeDocument | Add-Member ScriptMethod Close {
        if ($script:failureStage -eq 'cleanup') {throw 'Controlled cleanup failure'}
    }
    $documents=[pscustomobject]@{}
    $open={
        if ($script:failureStage -eq 'open') {throw 'Controlled open failure'}
        return $script:probeDocument
    }
    $documents | Add-Member ScriptMethod Open $open
    $documents | Add-Member ScriptMethod OpenNoRepairDialog $open
    $application=[pscustomobject]@{Version='control';DisplayAlerts=0;Visible=$false;Presentations=$documents;Documents=$documents}
    $application | Add-Member ScriptMethod Quit {}
    return $application
}
'@
$marker='$ErrorActionPreference = "Stop"'
if (($actualWorkerSource.Split($marker,[StringSplitOptions]::None)).Count -ne 2) {throw 'Worker injection point changed'}
$actualWorker=Join-Path $root 'actual worker with activation stub.ps1'
$actualWorkerSource.Replace($marker,$marker+[Environment]::NewLine+$activationStub) | Set-Content $actualWorker -Encoding utf8
$stageChecks=@()
foreach($extension in @('pptx','docx')) {
    foreach($stage in @('activation','open','export','cleanup')) {
        $folder=Join-Path $root ('actual-'+$extension+'-'+$stage)
        New-Item -ItemType Directory -Path $folder | Out-Null
        $inputFile=Join-Path $folder ($stage+'.'+$extension)
        [IO.File]::WriteAllText($inputFile,'controlled input; COM never runs')
        $output=Join-Path $folder 'output';$errorText=$null
        try { & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFile -OutputDirectory $output -WorkerScript $actualWorker -TimeoutSeconds 20 *> (Join-Path $folder 'supervisor.log') }
        catch {$errorText=$_.Exception.Message}
        $status=Get-Content (Join-Path $output 'reference-status.json') -Raw | ConvertFrom-Json
        $expected=if($stage -eq 'activation') {'unavailable'} else {'export-failed'}
        if (!$errorText -or $status.Status -ne $expected -or $status.Stage -cne $stage -or
            $status.Error -notlike ('*Controlled '+$stage+' failure*')) {throw ('Actual worker attributed '+$extension+'/'+$stage+' incorrectly: '+($status | ConvertTo-Json -Compress))}
        $stageChecks+=@{Format=$extension;FailureStage=$stage;ReportedStage=$status.Stage;Status=$status.Status}
    }
}
@{Checks=$checks;ActualWorkerStageChecks=$stageChecks;Builds=0;OfficeExports=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Reference worker: $($checks.Count) supervision and $($stageChecks.Count) actual-worker stage checks pass; no build or Office."
