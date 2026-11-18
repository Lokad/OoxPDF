# Exercise the shared build helper with an isolated project and imported MSBuild props.
param([ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$buildHelper = Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1'
$scratch = Join-Path $repoRoot ('artifacts/build-incrementality/' + [guid]::NewGuid().ToString('N'))
$projectDirectory = Join-Path $scratch 'probe'
New-Item -ItemType Directory -Path $projectDirectory -Force | Out-Null
$project = Join-Path $projectDirectory 'probe.csproj'
$source = Join-Path $projectDirectory 'Program.cs'
$properties = Join-Path $scratch 'Directory.Build.props'
$outputDll = Join-Path $projectDirectory "bin/Debug/$Framework/probe.dll"
$recordPath = Join-Path $scratch 'build-info.json'
$buildArguments = @{
    Project = $project
    Framework = $Framework
    OutputDll = $outputDll
    Description = 'Shared build incrementality probe'
}
Set-Content -LiteralPath $project -Value ('<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>'+ $Framework + '</TargetFramework><OutputType>Exe</OutputType></PropertyGroup></Project>')
Set-Content -LiteralPath $source -Value @'
#if UPDATED_BUILD
System.Console.WriteLine("updated");
#else
System.Console.WriteLine("initial");
#endif
'@
Set-Content -LiteralPath $properties -Value '<Project />'
& $buildHelper @buildArguments
if ((& dotnet $outputDll) -ne 'initial') { throw 'Initial build did not run the initial candidate.' }
Write-Host 'PASS Initial candidate builds and executes'
$sourceTimestamp = (Get-Item -LiteralPath $source).LastWriteTimeUtc
$projectTimestamp = (Get-Item -LiteralPath $project).LastWriteTimeUtc
$initialHash = (Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash

Set-Content -LiteralPath $properties -Value '<Project><PropertyGroup><DefineConstants>UPDATED_BUILD</DefineConstants></PropertyGroup></Project>'
& $buildHelper @buildArguments
if ((& dotnet $outputDll) -ne 'updated') { throw 'A changed parent build property silently reused the previous candidate.' }
if ((Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash -eq $initialHash) { throw 'The changed build property did not change candidate bytes.' }
if ((Get-Item -LiteralPath $source).LastWriteTimeUtc -ne $sourceTimestamp -or
    (Get-Item -LiteralPath $project).LastWriteTimeUtc -ne $projectTimestamp) { throw 'The probe must leave source/project timestamps unchanged.' }
Write-Host 'PASS Parent build properties refresh candidate bytes with unchanged source timestamps'

$updatedHash = (Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash
& $buildHelper @buildArguments -RecordPath $recordPath
$record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
if ($record.Framework -cne $Framework -or $record.Configuration -cne 'Debug' -or
    $record.OutputDll -cne $outputDll -or $record.OutputSha256 -cne $updatedHash.ToLowerInvariant()) {
    throw 'Build provenance does not describe the executed candidate.'
}
if ((Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash -ne $updatedHash) { throw 'An unchanged incremental build changed candidate bytes.' }
Write-Host 'PASS Unchanged builds retain candidate bytes and record their hash'

# The previous candidate still exists, but malformed build inputs must fail
# before any caller can execute it as the current candidate.
Set-Content -LiteralPath $properties -Value '<Project><PropertyGroup>'
$failed = $false
try {
    & $buildHelper @buildArguments
} catch {
    $failed = $_.Exception.Message -like '*build failed with exit code*'
}
if (-not $failed -or -not (Test-Path -LiteralPath $outputDll)) { throw 'A failed build must reject the existing candidate DLL.' }
Write-Host 'PASS Failed builds reject existing candidate output'
[ordered]@{
    Framework = $Framework
    HelperSha256 = (Get-FileHash -LiteralPath $buildHelper).Hash
    InitialSha256 = $initialHash
    RebuiltSha256 = $updatedHash
    Checks = @('initial execution', 'imported property rebuild', 'incremental identity', 'failed build refusal')
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $scratch 'evidence.json') -Encoding utf8
Write-Host "Four shared build checks passed. Evidence: $scratch"
exit 0
