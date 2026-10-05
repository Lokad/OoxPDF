# Package smoke test (T08): packs the library as shipped, then converts a
# document through the packed package only: no project references, no tools,
# no Office/COM, and no reference cache. Each run has its own package feed,
# restore cache, and consumer directory so older packages cannot satisfy the test.
# On hosts without Windows fonts,
# pass an embeddable TrueType font through -FontPath to exercise a custom resolver.

param(
    [string] $Configuration = "Release",
    [string] $FontPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not [string]::IsNullOrWhiteSpace($FontPath)) {
    $FontPath = (Resolve-Path -LiteralPath $FontPath).Path
}

$libraryProject = Join-Path $repoRoot "src/Lokad.OoxPdf/Lokad.OoxPdf.csproj"
$runId = [Guid]::NewGuid().ToString("N")
$packDirectory = Join-Path $repoRoot "artifacts/nuget/package-smoke/$runId"
$scratch = Join-Path $repoRoot "artifacts/package-smoke/$runId"
New-Item -ItemType Directory -Force -Path $packDirectory, $scratch | Out-Null
& dotnet pack $libraryProject -c $Configuration --nologo --output $packDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Library packing failed with exit code $LASTEXITCODE."
}

$packages = @(Get-ChildItem -LiteralPath $packDirectory -Filter "Lokad.OoxPdf.*.nupkg")
if ($packages.Count -ne 1) {
    throw "Expected exactly one freshly packed Lokad.OoxPdf package, found $($packages.Count)."
}
$package = $packages[0]
$version = $package.BaseName.Substring("Lokad.OoxPdf.".Length)
Write-Host ("Smoking package {0}." -f $package.Name)

$archive = [IO.Compression.ZipFile]::OpenRead($package.FullName)
try {
    $libraryEntry = $archive.GetEntry("lib/net10.0/Lokad.OoxPdf.dll")
    if ($null -eq $libraryEntry) { throw "Packed library DLL is missing." }
    $libraryStream = $libraryEntry.Open()
    try {
        $librarySha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($libraryStream))
    }
    finally { $libraryStream.Dispose() }
}
finally { $archive.Dispose() }

Copy-Item -LiteralPath (Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/docx-ladder-01-plain-paragraph.docx") -Destination (Join-Path $scratch "input.docx")

$escapedFeedPath = [Security.SecurityElement]::Escape($packDirectory)
Set-Content -LiteralPath (Join-Path $scratch "NuGet.config") -Encoding UTF8 -Value @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="local-pack" value="$escapedFeedPath" />
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
using System.Security.Cryptography;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

string input = args[0];
string output = args[1];
string expectedLibrarySha256 = args[2];
string loadedLibrarySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(OoxPdfConverter).Assembly.Location)));
if (!string.Equals(expectedLibrarySha256, loadedLibrarySha256, StringComparison.Ordinal))
{
    throw new InvalidDataException("Smoke consumer loaded a DLL different from the freshly packed library.");
}
Console.WriteLine("Loaded packed library SHA-256 " + loadedLibrarySha256 + ".");
var options = new OoxPdfOptions
{
    FontResolver = args.Length > 3 ? new SmokeFontResolver(args[3]) : null
};
OoxPdfConverter.Convert(input, output, options);
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
string pdf = Encoding.ASCII.GetString(File.ReadAllBytes(output));
if (!pdf.Contains("/FontFile2", StringComparison.Ordinal) || !pdf.Contains("/ToUnicode", StringComparison.Ordinal))
{
    throw new InvalidDataException("Smoke output must embed a TrueType font and its Unicode map.");
}

sealed class SmokeFontResolver(string path) : IFontResolver
{
    private readonly IFontProgramSource source = new MemoryFontProgramSource("smoke-font", File.ReadAllBytes(path));

    public FontFaceResolution Resolve(FontRequest request) => new(
        request.FamilyName,
        "Smoke font",
        new FontStyleKey(false, false, 400, 0, false),
        source,
        IsFallback: true);
}
'@
Set-Content -LiteralPath (Join-Path $scratch "Program.cs") -Encoding UTF8 -Value $program

$input = Join-Path $scratch "input.docx"
$output = Join-Path $scratch "output.pdf"
$smokeArguments = @($input, $output, $librarySha256)
if (-not [string]::IsNullOrWhiteSpace($FontPath)) { $smokeArguments += $FontPath }
$smokeProject = Join-Path $scratch "smoke.csproj"
& dotnet restore $smokeProject --packages (Join-Path $scratch "packages") --configfile (Join-Path $scratch "NuGet.config") --nologo
if ($LASTEXITCODE -ne 0) {
    throw "Isolated package restore failed with exit code $LASTEXITCODE."
}
& dotnet run --no-restore --project $smokeProject -c $Configuration -- @smokeArguments
if ($LASTEXITCODE -ne 0) {
    throw "Smoke conversion failed with exit code $LASTEXITCODE."
}

$pdfHeader = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($output)[0..4])
if ($pdfHeader -ne "%PDF-") {
    throw "Smoke output does not start with the PDF header: $pdfHeader."
}
Write-Host "Package smoke test passed through the packed library."
[ordered]@{
    Package = $package.FullName
    PackageVersion = $version
    PackageSha256 = (Get-FileHash -LiteralPath $package.FullName -Algorithm SHA256).Hash
    LibrarySha256 = $librarySha256
    OutputSha256 = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
    OutputBytes = (Get-Item -LiteralPath $output).Length
    RestorePackages = Join-Path $scratch "packages"
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $scratch "summary.json") -Encoding UTF8
Write-Host ("Evidence: {0}" -f (Join-Path $scratch "summary.json"))
