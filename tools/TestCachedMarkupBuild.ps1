# Exercise the cached comparison's production build gate with an isolated
# project. Parent MSBuild properties affect the output without touching C#.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot 'CompareCachedDocxMarkupReference.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Cached comparison script did not parse.' }
$buildFunction = $ast.FindAll({ param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-DotnetBuildIfStale'
}, $false) | Select-Object -First 1
if ($null -eq $buildFunction) { throw 'Cached comparison build gate is missing.' }
. ([scriptblock]::Create($buildFunction.Extent.Text))

$scratch = Join-Path $repoRoot ('artifacts/cached-markup-build-test/' + [guid]::NewGuid().ToString('N'))
$projectDirectory = Join-Path $scratch 'probe'
New-Item -ItemType Directory -Path $projectDirectory -Force | Out-Null
$project = Join-Path $projectDirectory 'probe.csproj'
$source = Join-Path $projectDirectory 'Program.cs'
$properties = Join-Path $scratch 'Directory.Build.props'
$outputDll = Join-Path $projectDirectory 'bin/Debug/net10.0/probe.dll'
Set-Content -LiteralPath $project -Value '<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>net10.0</TargetFramework><OutputType>Exe</OutputType></PropertyGroup></Project>'
Set-Content -LiteralPath $source -Value @'
#if UPDATED_BUILD
System.Console.WriteLine("updated");
#else
System.Console.WriteLine("initial");
#endif
'@
Set-Content -LiteralPath $properties -Value '<Project />'
Invoke-DotnetBuildIfStale -Project $project -OutputDll $outputDll -Description 'Cached markup build probe'
if ((& dotnet $outputDll) -ne 'initial') { throw 'Initial build did not run the initial candidate.' }
Write-Host 'PASS Initial candidate builds and executes'
$sourceTimestamp = (Get-Item -LiteralPath $source).LastWriteTimeUtc
$projectTimestamp = (Get-Item -LiteralPath $project).LastWriteTimeUtc
$initialHash = (Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash

Set-Content -LiteralPath $properties -Value '<Project><PropertyGroup><DefineConstants>UPDATED_BUILD</DefineConstants></PropertyGroup></Project>'
Invoke-DotnetBuildIfStale -Project $project -OutputDll $outputDll -Description 'Cached markup build probe'
if ((& dotnet $outputDll) -ne 'updated') { throw 'A changed parent build property silently reused the previous candidate.' }
if ((Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash -eq $initialHash) { throw 'The changed build property did not change candidate bytes.' }
if ((Get-Item -LiteralPath $source).LastWriteTimeUtc -ne $sourceTimestamp -or
    (Get-Item -LiteralPath $project).LastWriteTimeUtc -ne $projectTimestamp) { throw 'The probe must leave source/project timestamps unchanged.' }
Write-Host 'PASS Parent build properties refresh candidate bytes with unchanged source timestamps'

$updatedHash = (Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash
$recordPath = Join-Path $scratch 'build-info.json'
Invoke-DotnetBuildIfStale -Project $project -OutputDll $outputDll -Description 'Cached markup build probe' -RecordPath $recordPath
$record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
if ($record.OutputSha256 -ne $updatedHash.ToLowerInvariant()) { throw 'Build provenance does not describe the executed candidate.' }
if ((Get-FileHash -LiteralPath $outputDll -Algorithm SHA256).Hash -ne $updatedHash) { throw 'An unchanged incremental build changed candidate bytes.' }
Write-Host 'PASS Unchanged builds retain candidate bytes and record their hash'

# The previous candidate still exists, but malformed build inputs must fail
# before any caller can execute it as the current candidate.
Set-Content -LiteralPath $properties -Value '<Project><PropertyGroup>'
$failed = $false
try {
    Invoke-DotnetBuildIfStale -Project $project -OutputDll $outputDll -Description 'Cached markup build probe'
} catch {
    $failed = $_.Exception.Message -like '*build failed with exit code*'
}
if (-not $failed -or -not (Test-Path -LiteralPath $outputDll)) { throw 'A failed build must reject the existing candidate DLL.' }
Write-Host 'PASS Failed builds reject existing candidate output'
Write-Host "Cached markup build checks passed. Evidence: $scratch"
exit 0
