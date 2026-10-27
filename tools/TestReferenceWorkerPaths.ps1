# Exercise the existing Office supervisor with an owned fake worker; no COM/build.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/reference-worker-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) {throw 'Use a fresh test directory'}
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$worker=Join-Path $root 'worker.ps1'
@'
param($InputPath,$WorkDirectory,[int]$Dpi,$ProgressLog,$StatusPath,[switch]$SkipBuild)
$ErrorActionPreference='Stop'
if ($args.Count) {throw 'Unexpected unbound arguments'}
Add-Content -LiteralPath $ProgressLog -Value 'stage:export'
'fake worker stdout'; [Console]::Error.WriteLine('fake worker stderr')
$mode=[IO.Path]::GetFileNameWithoutExtension($InputPath)
if ($mode -eq 'hang') {Start-Sleep -Seconds 60}
if ($mode -eq 'no-status') {exit 1}
if ($mode -in @('export-failed','failed-reused','hang-with-status')) {
    @{Status='export-failed';Stage='export';Error='Synthetic current export failure';OfficeVersion='CURRENT-'+$mode} | ConvertTo-Json | Set-Content $StatusPath -Encoding utf8
    if ($mode -eq 'hang-with-status') {Start-Sleep -Seconds 60}
    exit 1
}
if ($mode -eq 'unavailable') {
    @{Status='unavailable';Stage='activation';Error='Synthetic activation failure';OfficeVersion='CURRENT-'+$mode} | ConvertTo-Json | Set-Content $StatusPath -Encoding utf8
    exit 1
}
@{InputPath=$InputPath;WorkDirectory=$WorkDirectory;Dpi=$Dpi;ProgressLog=$ProgressLog;StatusPath=$StatusPath;SkipBuild=[bool]$SkipBuild} | ConvertTo-Json | Set-Content (Join-Path $WorkDirectory 'arguments.json') -Encoding utf8
Copy-Item -LiteralPath $InputPath -Destination (Join-Path $WorkDirectory 'reference.pdf')
if ($mode -eq 'bad-pdf') {Remove-Item -LiteralPath (Join-Path $WorkDirectory 'reference.pdf')}
@{Status='ok';Stage='done';Error='';OfficeVersion='CURRENT-'+$mode} | ConvertTo-Json | Set-Content $StatusPath -Encoding utf8
'@ | Set-Content $worker -Encoding utf8
$spacedWorker=Join-Path $root 'worker with spaces.ps1'
Copy-Item $worker $spacedWorker
$checks=@()
$sentinelStart=[Diagnostics.ProcessStartInfo]::new('pwsh')
$sentinelStart.UseShellExecute=$false;$sentinelStart.CreateNoWindow=$true
foreach($arg in @('-NoProfile','-Command','Start-Sleep -Seconds 60')) {$sentinelStart.ArgumentList.Add($arg)}
$sentinel=[Diagnostics.Process]::Start($sentinelStart)
try {
    foreach($id in @('plain','spaced-input','spaced-output','spaced-worker','all-spaced','prebuilt','unavailable','hang','export-failed','failed-reused','no-status','hang-with-status','bad-pdf')) {
        $inputName=if($id -in @('spaced-input','all-spaced')) {'input ''quoted'' é.pdf'} else {$id+'.pdf'}
        $inputFile=Join-Path $root $inputName
        [IO.File]::WriteAllText($inputFile,'%PDF-Synthetic supervisor probe, not a fidelity reference')
        $output=Join-Path $root $(if($id -in @('spaced-output','all-spaced')) {$id+' output with spaces'} else {$id+'-output'})
        $selected=if($id -in @('spaced-worker','all-spaced')) {$spacedWorker} else {$worker}
        $timeout=if($id -in @('hang','hang-with-status')) {2} else {20}
        New-Item -ItemType Directory -Path $output | Out-Null
        $retained=Join-Path $output 'retained-evidence.txt'
        [IO.File]::WriteAllText($retained,'independent evidence survives')
        $snapshot=Join-Path $output 'reference-worker-status.json'
        if ($id -ne 'export-failed') {
            @{Status='ok';OfficeVersion='OLD-RUN'} | ConvertTo-Json | Set-Content $snapshot
        }
        $errorText=$null
        try { & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFile -OutputDirectory $output -WorkerScript $selected -Dpi 123 -TimeoutSeconds $timeout -SkipBuild:($id -eq 'prebuilt') *> (Join-Path $root ($id+'.log')) }
        catch {$errorText=$_.Exception.Message}
        $status=Get-Content (Join-Path $output 'reference-status.json') -Raw | ConvertFrom-Json
        $expected=if($id -in @('hang','hang-with-status')) {'timeout'} elseif($id -eq 'unavailable') {'unavailable'} elseif($id -in @('export-failed','failed-reused','no-status','bad-pdf')) {'export-failed'} else {'ok'}
        if($status.Status -ne $expected -or (($expected -eq 'ok') -eq ($null -ne $errorText))) {throw ($id+': unexpected supervisor outcome '+$errorText)}
        if ($id -in @('hang','no-status')) {
            if (Test-Path -LiteralPath $snapshot) {throw ($id+': stale worker metadata survived')}
        } else {
            $raw=Get-Content -LiteralPath $snapshot -Raw | ConvertFrom-Json
            if ($raw.OfficeVersion -cne ('CURRENT-'+[IO.Path]::GetFileNameWithoutExtension($inputFile))) {throw ($id+': current worker metadata was lost')}
            $rawExpected=if($id -in @('export-failed','failed-reused','hang-with-status')) {'export-failed'} elseif($id -eq 'unavailable') {'unavailable'} else {'ok'}
            if ($raw.Status -ne $rawExpected) {throw ($id+': raw worker status was changed')}
        }
        if ([IO.File]::ReadAllText($retained) -cne 'independent evidence survives' -or @(Get-ChildItem -LiteralPath $output -Directory -Filter '_worker-*').Count) {throw ($id+': evidence retention/owned cleanup failed')}
        if($expected -eq 'ok') {
            $actual=Get-Content (Join-Path $output 'arguments.json') -Raw | ConvertFrom-Json
            if($actual.InputPath -cne $inputFile -or $actual.Dpi -ne 123 -or $actual.SkipBuild -ne ($id -eq 'prebuilt') -or
                !(Split-Path $actual.WorkDirectory -Parent).Equals($output,[StringComparison]::Ordinal) -or
                [IO.File]::ReadAllText((Join-Path $output 'reference.pdf')) -ne [IO.File]::ReadAllText($inputFile)) {throw ($id+': argument/output identity differs')}
        }
        foreach($name in @('stdout','stderr')) {
            if([IO.File]::ReadAllText((Join-Path $output ('reference-worker-'+$name+'.log'))) -notlike ('*fake worker '+$name+'*')) {throw 'Worker log was lost'}
        }
        if($sentinel.HasExited) {throw 'Supervisor terminated an unrelated owned control process'}
        $checks+=@{Id=$id;Status=$status.Status;Error=$errorText;CurrentWorkerMetadata=(Test-Path -LiteralPath $snapshot);UnrelatedEvidenceRetained=$true;OwnedWorkerDirectoryRemoved=$true}
    }
}
finally {if(!$sentinel.HasExited) {$sentinel.Kill();$null=$sentinel.WaitForExit(5000)};$sentinel.Dispose()}

# Exercise the actual worker, replacing only COM activation with managed stubs.
# Rasterization is also stubbed; these checks never build or use Office/PDFium.
$actualWorkerSource=Get-Content (Join-Path $PSScriptRoot 'RenderReferenceWorker.ps1') -Raw
$activationStub=@'
function Enter-Operation([string] $Name) {
    if ($script:failureStage -eq ('delay-'+$Name)) {
        [IO.File]::WriteAllText($InputPath+'.entered',$Name)
        Start-Sleep -Seconds 60
    }
}
function New-Object {
    param([string] $ComObject)
    if ($ComObject -notin @('PowerPoint.Application','Word.Application')) {throw 'Unexpected COM request'}
    $script:failureStage=[IO.Path]::GetFileNameWithoutExtension($InputPath)
    Enter-Operation 'activation'
    if ($script:failureStage -eq 'activation') {throw 'Controlled activation failure'}
    $script:probeDocument=[pscustomobject]@{}
    $save={param($Path,$Format)
        Enter-Operation 'export'
        if ($script:failureStage -eq 'export') {throw 'Controlled export failure'}
        [IO.File]::WriteAllText($Path,'%PDF-control:'+$InputPath)
    }
    $script:probeDocument | Add-Member ScriptMethod SaveAs $save
    $script:probeDocument | Add-Member ScriptMethod SaveAs2 $save
    $script:probeDocument | Add-Member ScriptMethod Close {
        Enter-Operation 'cleanup'
        if ($script:failureStage -eq 'cleanup') {throw 'Controlled cleanup failure'}
    }
    $documents=[pscustomobject]@{}
    $open={
        Enter-Operation 'open'
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
@'
param($InputPdf,$OutputDirectory,$Dpi,[switch]$SkipBuild)
$inputSource=[IO.File]::ReadAllText($InputPdf).Substring('%PDF-control:'.Length)
if ([IO.Path]::GetFileNameWithoutExtension($inputSource) -eq 'delay-rasterize') {
    [IO.File]::WriteAllText($inputSource+'.entered','rasterize')
    Start-Sleep -Seconds 60
}
@{SkipBuild=[bool]$SkipBuild} | ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'raster-arguments.json')
'@ | Set-Content (Join-Path $root 'RasterizePdf.ps1') -Encoding utf8
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
$timeoutChecks=@();$rasterChecks=@()
foreach($extension in @('pptx','docx')) {
    foreach($stage in @('activation','open','export','cleanup','rasterize')) {
        $folder=Join-Path $root ('delay-'+$extension+'-'+$stage)
        New-Item -ItemType Directory -Path $folder | Out-Null
        $inputFile=Join-Path $folder ('delay-'+$stage+'.'+$extension)
        [IO.File]::WriteAllText($inputFile,'controlled input; COM never runs')
        $output=Join-Path $folder 'output';$errorText=$null
        try { & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFile -OutputDirectory $output -WorkerScript $actualWorker -TimeoutSeconds 3 *> (Join-Path $folder 'supervisor.log') }
        catch {$errorText=$_.Exception.Message}
        $status=Get-Content (Join-Path $output 'reference-status.json') -Raw | ConvertFrom-Json
        if (!$errorText -or $status.Status -ne 'timeout' -or $status.Stage -cne $stage -or
            [IO.File]::ReadAllText($inputFile+'.entered') -cne $stage) {throw ('Active timeout stage differs: '+$extension+'/'+$stage)}
        $timeoutChecks+=@{Format=$extension;ActiveStage=$stage;ReportedStage=$status.Stage;Status=$status.Status}
    }
    foreach($skip in @($false,$true)) {
        $folder=Join-Path $root ('raster-'+$extension+'-'+$skip)
        New-Item -ItemType Directory -Path $folder | Out-Null
        $inputFile=Join-Path $folder ('success.'+$extension)
        [IO.File]::WriteAllText($inputFile,'controlled input; COM never runs')
        $output=Join-Path $folder 'output'
        & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFile -OutputDirectory $output -WorkerScript $actualWorker -TimeoutSeconds 20 -SkipBuild:$skip *> (Join-Path $folder 'supervisor.log')
        $actual=Get-Content (Join-Path $output 'raster-arguments.json') -Raw | ConvertFrom-Json
        if ($actual.SkipBuild -ne $skip) {throw 'Actual worker did not forward rasterizer build selection'}
        $rasterChecks+=@{Format=$extension;SkipBuild=$skip;Forwarded=$actual.SkipBuild}
    }
}
@{Checks=$checks;ActualWorkerStageChecks=$stageChecks;TimeoutStageChecks=$timeoutChecks;RasterBuildChecks=$rasterChecks;Builds=0;OfficeExports=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Reference worker: $($checks.Count) supervision, $($stageChecks.Count) exception, $($timeoutChecks.Count) timeout and $($rasterChecks.Count) raster forwarding checks pass; no build or Office."
