# Build the existing allocation probe against an already qualified library.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $ArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$false
$artifact=(Resolve-Path -LiteralPath $ArtifactDirectory).Path
$pin=Get-Content -LiteralPath (Join-Path $artifact 'identity.json') -Raw | ConvertFrom-Json
$library=Join-Path $artifact 'src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll'
if ((Get-FileHash -LiteralPath $library).Hash -ne $pin.LibrarySha256) { throw 'Frozen library identity differs' }
$output=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Use a fresh allocation runtime directory' }
$sources=@(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'Lokad.OoxPdf.AllocProbe') -Filter '*.cs' -File)
if (!$sources.Count) { throw 'Allocation probe source is missing' }
New-Item -ItemType Directory -Path (Join-Path $output 'source') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $output 'library') | Out-Null
$sourceHashes=[ordered]@{}
foreach ($source in $sources) {
    $hash=(Get-FileHash -LiteralPath $source.FullName).Hash
    $target=Join-Path $output ('source/'+$source.Name)
    Copy-Item -LiteralPath $source.FullName -Destination $target
    if ((Get-FileHash -LiteralPath $target).Hash -ne $hash) { throw 'Probe source copy differs' }
    $sourceHashes[$source.Name]=$hash
}
$snapshotLibrary=Join-Path $output 'library/Lokad.OoxPdf.dll'
Copy-Item -LiteralPath $library -Destination $snapshotLibrary
if ((Get-FileHash -LiteralPath $snapshotLibrary).Hash -ne $pin.LibrarySha256) { throw 'Probe library copy differs' }
$project=Join-Path $output 'Lokad.OoxPdf.AllocProbe.csproj'
@'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <AssemblyName>Lokad.OoxPdf.AllocProbe</AssemblyName>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <EnableDefaultCompileItems>false</EnableDefaultCompileItems>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="source/*.cs" />
    <Reference Include="Lokad.OoxPdf">
      <HintPath>library/Lokad.OoxPdf.dll</HintPath>
    </Reference>
  </ItemGroup>
</Project>
'@ | Set-Content -LiteralPath $project -Encoding utf8
# No ProjectReference: this builds only the probe and never the library.
& dotnet build $project -c Release --nologo -v minimal *> (Join-Path $output 'build.log')
if ($LASTEXITCODE -ne 0) { throw ('Allocation probe build failed; inspect '+(Join-Path $output 'build.log')) }
$runtime=Join-Path $output 'bin/Release/net10.0'
if ((Get-FileHash -LiteralPath (Join-Path $runtime 'Lokad.OoxPdf.dll')).Hash -ne $pin.LibrarySha256 -or
    (Get-FileHash -LiteralPath $library).Hash -ne $pin.LibrarySha256) { throw 'Qualified library changed during probe preparation' }
$probe=Join-Path $runtime 'Lokad.OoxPdf.AllocProbe.dll'
& dotnet $probe --self-test *> (Join-Path $output 'self-test.log')
if ($LASTEXITCODE -ne 0) { throw ('Allocation probe self-test failed; inspect '+(Join-Path $output 'self-test.log')) }
[ordered]@{
    FrozenRevision=$pin.Revision;LibrarySha256=$pin.LibrarySha256
    ProbeBuilt=$true;LibraryRebuilt=$false;SourceHashes=$sourceHashes
    ProjectSha256=(Get-FileHash -LiteralPath $project).Hash
    ProbeSha256=(Get-FileHash -LiteralPath $probe).Hash
    SelfTestPassed=$true;ProbePath=$probe
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
Write-Output ('Allocation probe prepared against frozen '+$pin.Revision+'; library not rebuilt.')
