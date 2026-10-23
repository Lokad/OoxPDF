# Exercise the real build helper with stubbed builds; never invokes dotnet.
[CmdletBinding()]
param([string]$OutputDirectory, [switch]$Worker, [string]$Tool,
    [string]$Project, [string]$OutputDll, [string]$Marker, [string]$Started,
    [string]$Ready, [string]$Release, [string]$Fault)
$ErrorActionPreference = 'Stop'
if ($Worker) {
    $global:ooxPdfIsolationMarker=$Marker
    $global:ooxPdfIsolationReady=$Ready
    $global:ooxPdfIsolationRelease=$Release
    $global:ooxPdfIsolationFault=$Fault
    function global:dotnet {
        if ($global:ooxPdfIsolationReady) {
            [IO.File]::WriteAllText($global:ooxPdfIsolationReady, 'entered build')
            $deadline=[DateTime]::UtcNow.AddSeconds(45)
            while (!(Test-Path -LiteralPath $global:ooxPdfIsolationRelease)) {
                if ([DateTime]::UtcNow -ge $deadline) { throw 'Stub release deadline exceeded' }
                Start-Sleep -Milliseconds 25
            }
        }
        [IO.File]::WriteAllText($global:ooxPdfIsolationMarker, 'stub build invoked')
        $global:LASTEXITCODE=if ($global:ooxPdfIsolationFault -eq 'build') {17} else {0}
    }
    [IO.File]::WriteAllText($Started, 'invoking helper')
    & $Tool -Project $Project -OutputDll $OutputDll -Description 'Isolation check'
    exit 0
}
$repo=Split-Path $PSScriptRoot
if (!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/build-isolation-'+[guid]::NewGuid().ToString('N')) }
$output=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Use a fresh test directory' }
New-Item -ItemType Directory -Path $output | Out-Null
foreach ($label in @('checkout A','checkout B')) {
    $checkout=Join-Path $output $label
    New-Item -ItemType Directory -Path (Join-Path $checkout 'tools') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1') -Destination (Join-Path $checkout 'tools')
    foreach ($projectName in @('first','second')) {
        [IO.File]::WriteAllText((Join-Path $checkout ($projectName+'.csproj')), '<Project />')
    }
    [IO.File]::WriteAllText((Join-Path $checkout 'output.dll'), 'stub; never loaded')
}
$alias=Join-Path $output 'checkout alias'
$linkType=if ($IsWindows) {'Junction'} else {'SymbolicLink'}
New-Item -ItemType $linkType -Path $alias -Target (Join-Path $output 'checkout A') | Out-Null
$processes=[Collections.Generic.List[Diagnostics.Process]]::new()
function Start-Check([string]$Name,[string]$Checkout,[bool]$Hold=$false,[string]$Failure='',[string]$ProjectName='first') {
    $start=[Diagnostics.ProcessStartInfo]::new('pwsh')
    $start.UseShellExecute=$false; $start.CreateNoWindow=$true
    $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
    $dll=if ($Failure -eq 'missing') {'missing.dll'} else {'output.dll'}
    foreach ($arg in @('-NoProfile','-File',$PSCommandPath,'-Worker',
        '-Tool',(Join-Path $Checkout 'tools/EnsureDotnetBuild.ps1'),
        '-Project',(Join-Path $Checkout ($ProjectName+'.csproj')),
        '-OutputDll',(Join-Path $Checkout $dll),'-Marker',(Join-Path $output ($Name+'.built')),
        '-Started',(Join-Path $output ($Name+'.started')))) { $start.ArgumentList.Add($arg) }
    if ($Hold) {
        foreach ($arg in @('-Ready',(Join-Path $output 'holder.ready'),'-Release',(Join-Path $output 'release'))) { $start.ArgumentList.Add($arg) }
    }
    if ($Failure) { $start.ArgumentList.Add('-Fault'); $start.ArgumentList.Add($Failure) }
    $process=[Diagnostics.Process]::Start($start); $processes.Add($process)
    return $process
}
function Wait-File([string]$Name,[Diagnostics.Process]$Process) {
    $path=Join-Path $output $Name; $deadline=[DateTime]::UtcNow.AddSeconds(10)
    while (!(Test-Path -LiteralPath $path)) {
        if ($Process.HasExited -or [DateTime]::UtcNow -ge $deadline) { throw ('Check did not reach '+$Name) }
        Start-Sleep -Milliseconds 25
    }
}
function Finish-Check([string]$Name,[Diagnostics.Process]$Process,[string]$ExpectedError='') {
    if (!$Process.WaitForExit(10000)) { throw ('Check timed out: '+$Name) }
    $log=$Process.StandardOutput.ReadToEnd()+$Process.StandardError.ReadToEnd()
    [IO.File]::WriteAllText((Join-Path $output ($Name+'.log')), $log)
    if ($ExpectedError) {
        if ($Process.ExitCode -eq 0 -or !$log.Contains($ExpectedError)) { throw ('Expected refusal missing: '+$Name) }
    } elseif ($Process.ExitCode -ne 0) { throw ('Check failed: '+$Name+'; '+$log) }
}
try {
    $a=Start-Check 'holder' (Join-Path $output 'checkout A') $true
    Wait-File 'holder.ready' $a
    $same=Start-Check 'same' (Join-Path $output 'checkout A') $false '' 'second'
    $aliased=Start-Check 'alias' $alias
    Wait-File 'same.started' $same; Wait-File 'alias.started' $aliased
    $other=Start-Check 'independent' (Join-Path $output 'checkout B')
    Finish-Check 'independent' $other
    foreach ($name in @('same','alias')) {
        if (Test-Path -LiteralPath (Join-Path $output ($name+'.built'))) { throw ('Shared output lock bypassed: '+$name) }
    }
    if ($a.HasExited -or $same.HasExited -or $aliased.HasExited) { throw 'Shared checkout did not remain serialized' }
    [IO.File]::WriteAllText((Join-Path $output 'release'), 'release holder')
    Finish-Check 'holder' $a; Finish-Check 'same' $same; Finish-Check 'alias' $aliased
    $failed=Start-Check 'failure' (Join-Path $output 'checkout A') $false 'build'
    Finish-Check 'failure' $failed 'build failed with exit code 17'
    $missing=Start-Check 'missing' (Join-Path $output 'checkout A') $false 'missing'
    Finish-Check 'missing' $missing 'expected output is missing'
    [ordered]@{Checks=@('independent checkouts','shared referenced outputs','checkout aliases',
        'failed build rejects old output','missing output refuses');RealBuilds=0;
        HelperSha256=(Get-FileHash (Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1')).Hash} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'evidence.json') -Encoding utf8
    Write-Output 'Five build isolation/failure checks passed; no real builds.'
}
finally {
    foreach ($process in $processes) {
        if (!$process.HasExited) { $process.Kill($true); $null=$process.WaitForExit(5000) }
        $process.Dispose()
    }
}
