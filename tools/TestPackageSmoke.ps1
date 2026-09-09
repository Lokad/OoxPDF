# Package smoke test (T08): packs the library as shipped, then converts a
# document through the packed package only: no project references, no tools,
# no Office/COM, and no reference cache. Windows-only like CI (the default
# font resolver covers Windows installed fonts).

param(
    [string] $Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot

$libraryProject = Join-Path $repoRoot "src/Lokad.OoxPdf/Lokad.OoxPdf.csproj"
& dotnet pack $libraryProject -c $Configuration --nologo
if ($LASTEXITCODE -ne 0) {
    throw "Library packing failed with exit code $LASTEXITCODE."
}

$package = Get-ChildItem -LiteralPath (Join-Path $repoRoot "artifacts/nuget") -Filter "Lokad.OoxPdf.*.nupkg" |
    Where-Object { $_.Name -notlike "*.snupkg" } |
    Sort-Object Name -Descending |
    Select-Object -First 1
if ($null -eq $package) {
    throw "No packed Lokad.OoxPdf package found under artifacts/nuget."
}
$version = $package.BaseName.Substring("Lokad.OoxPdf.".Length)
Write-Host ("Smoking package {0}." -f $package.Name)

$scratch = Join-Path $repoRoot "artifacts/package-smoke"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

Copy-Item -LiteralPath (Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/docx-ladder-01-plain-paragraph.docx") -Destination (Join-Path $scratch "input.docx")

Set-Content -LiteralPath (Join-Path $scratch "NuGet.config") -Encoding UTF8 -Value @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="local-pack" value="../nuget" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
</configuration>
"@

$csproj = @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Lokad.OoxPdf" Version="$version" />
  </ItemGroup>
</Project>
"@
Set-Content -LiteralPath (Join-Path $scratch "smoke.csproj") -Encoding UTF8 -Value $csproj

$program = @'
using System.Text;
using Lokad.OoxPdf;

string input = args[0];
string output = args[1];
OoxPdfConverter.Convert(input, output);
byte[] header = new byte[5];
using (FileStream stream = File.OpenRead(output))
{
    if (stream.Read(header, 0, header.Length) != header.Length)
    {
        throw new InvalidDataException("Smoke output is shorter than the PDF header.");
    }
}
if (Encoding.ASCII.GetString(header) != "%PDF-")
{
    throw new InvalidDataException("Smoke output does not start with the PDF header.");
}
Console.WriteLine("Packed conversion produced " + new FileInfo(output).Length + " bytes.");
'@
Set-Content -LiteralPath (Join-Path $scratch "Program.cs") -Encoding UTF8 -Value $program

$input = Join-Path $scratch "input.docx"
$output = Join-Path $scratch "output.pdf"
& dotnet run --project (Join-Path $scratch "smoke.csproj") -c $Configuration -- $input $output
if ($LASTEXITCODE -ne 0) {
    throw "Smoke conversion failed with exit code $LASTEXITCODE."
}

$pdfHeader = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($output)[0..4])
if ($pdfHeader -ne "%PDF-") {
    throw "Smoke output does not start with the PDF header: $pdfHeader."
}
Write-Host "Package smoke test passed through the packed library."
