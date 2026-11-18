# Self-test an existing probe or derive one against an already qualified library.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $ArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [string] $ProbePath
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$false
$artifact=(Resolve-Path -LiteralPath $ArtifactDirectory).Path
$pin=Get-Content -LiteralPath (Join-Path $artifact 'identity.json') -Raw | ConvertFrom-Json
$framework=if ($pin.Framework) {$pin.Framework} else {'net10.0'}
if ($framework -notin @('net10.0','net11.0')) { throw 'Unknown frozen target framework' }
$library=Join-Path $artifact "src/Lokad.OoxPdf/bin/Release/$framework/Lokad.OoxPdf.dll"
if ((Get-FileHash -LiteralPath $library).Hash -ne $pin.LibrarySha256) { throw 'Frozen library identity differs' }
# Exercise real SFNT preparation with a font belonging to this frozen input set.
$fontRelative='tests/Lokad.OoxPdf.Tests/Cases/PortableFonts/DejaVuSans.ttf'
$font=Join-Path $artifact $fontRelative
$fontHash=$pin.Inputs.$fontRelative
if (!$fontHash -or (Get-FileHash -LiteralPath $font).Hash -ne $fontHash) { throw 'Frozen self-test font identity differs' }
$output=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Use a fresh allocation runtime directory' }
$sourceHashes=$null
$projectHash=$null
$selfTestFont=$font
if ($ProbePath) {
    $probe=(Resolve-Path -LiteralPath $ProbePath).Path
} else {
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
    @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>$framework</TargetFramework>
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
"@ | Set-Content -LiteralPath $project -Encoding utf8
    # No ProjectReference: this builds only the probe.
    & dotnet build $project -c Release --nologo -v minimal *> (Join-Path $output 'build.log')
    if ($LASTEXITCODE -ne 0) { throw ('Allocation probe build failed; inspect '+(Join-Path $output 'build.log')) }
    $probe=Join-Path $output "bin/Release/$framework/Lokad.OoxPdf.AllocProbe.dll"
    $projectHash=(Get-FileHash -LiteralPath $project).Hash
    $selfTestFont=Join-Path $output 'DejaVuSans.ttf'
    # Copy bytes without inheriting frozen read-only permissions on shared drives.
    [IO.File]::WriteAllBytes($selfTestFont, [IO.File]::ReadAllBytes($font))
    if ((Get-FileHash -LiteralPath $selfTestFont).Hash -ne $fontHash) { throw 'Self-test font copy differs' }
}

$probeLibrary=Join-Path (Split-Path -Parent $probe) 'Lokad.OoxPdf.dll'
if ((Get-FileHash -LiteralPath $probeLibrary).Hash -ne $pin.LibrarySha256) { throw 'Probe library differs from frozen identity' }
$probeHash=(Get-FileHash -LiteralPath $probe).Hash
New-Item -ItemType Directory -Path $output -Force | Out-Null
$selfTestLog=Join-Path $output 'self-test.log'
& dotnet $probe --self-test --font-file $selfTestFont *> $selfTestLog
if ($LASTEXITCODE -ne 0) { throw ('Allocation probe self-test failed; inspect '+$selfTestLog) }
$expectedIdentity='Self-test library SHA256: '+$pin.LibrarySha256.ToUpperInvariant()
if ($expectedIdentity -cnotin (Get-Content -LiteralPath $selfTestLog)) {
    throw 'Self-test did not report the frozen loaded library identity'
}
if ((Get-FileHash -LiteralPath $probe).Hash -ne $probeHash -or
    (Get-FileHash -LiteralPath $probeLibrary).Hash -ne $pin.LibrarySha256 -or
    (Get-FileHash -LiteralPath $library).Hash -ne $pin.LibrarySha256 -or
    (Get-FileHash -LiteralPath $selfTestFont).Hash -ne $fontHash) { throw 'Allocation self-test inputs changed during execution' }
[ordered]@{
    FrozenRevision=$pin.Revision;Framework=$framework;LibrarySha256=$pin.LibrarySha256
    LoadedLibrarySha256=$pin.LibrarySha256;ProbeBuilt=(!$ProbePath);LibraryRebuilt=$false;SourceHashes=$sourceHashes
    ProjectSha256=$projectHash;ProbeSha256=$probeHash
    SelfTestPassed=$true;ProbePath=$probe
    SelfTestFont=$fontRelative;SelfTestFontSha256=$fontHash;SelfTestStages=$true
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
Write-Output ('Allocation probe self-test passed against frozen '+$pin.Revision+'.')
