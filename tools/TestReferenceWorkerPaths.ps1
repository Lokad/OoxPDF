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
@{Checks=$checks;Builds=0;OfficeExports=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Reference worker: $($checks.Count) checks pass; no build or Office."
