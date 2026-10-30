# Package smoke test (T08): packs the library, or reuses -PackagePath, then checks
# DOCX/PPTX conversion and failure contracts through the packed package only:
# no project references, no tools,
# no Office/COM, and no reference cache. Each run has its own package feed,
# restore cache, and consumer directory so older packages cannot satisfy the test.
# On hosts without Windows fonts,
# pass an embeddable TrueType font through -FontPath to exercise a custom resolver.

param(
    [string] $Configuration = "Release",
    [string] $FontPath = "",
    [switch] $NoBuild,
    [string] $PackagePath = "",
    [string] $ExpectedLibrarySha256 = "",
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0'
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
if ($PackagePath) {
    $PackagePath = (Resolve-Path -LiteralPath $PackagePath).Path
    if (-not $ExpectedLibrarySha256) { throw 'Reusing a package requires the qualified library hash.' }
}
if (-not [string]::IsNullOrWhiteSpace($FontPath)) {
    $FontPath = (Resolve-Path -LiteralPath $FontPath).Path
}

$libraryProject = Join-Path $repoRoot "src/Lokad.OoxPdf/Lokad.OoxPdf.csproj"
$runId = [Guid]::NewGuid().ToString("N")
$packDirectory = Join-Path $repoRoot "artifacts/nuget/package-smoke/$runId"
$scratch = Join-Path $repoRoot "artifacts/package-smoke/$runId"
New-Item -ItemType Directory -Force -Path $packDirectory, $scratch | Out-Null
if ($PackagePath) {
    Copy-Item -LiteralPath $PackagePath -Destination $packDirectory
}
else {
    $packArguments = @("pack", $libraryProject, "-c", $Configuration, "--nologo", "--output", $packDirectory)
    if ($NoBuild) { $packArguments += "--no-build" }
    & dotnet @packArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Library packing failed with exit code $LASTEXITCODE."
    }
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
    $localDocumentationLinks = 0
    foreach ($required in @('README.md', 'CHANGELOG.md', 'docs/PublicApiChanges.md', 'docs/PreviewRelease.md')) {
        $entry = $archive.GetEntry($required)
        if ($null -eq $entry) { throw "Packed caller documentation is missing: $required" }
    }
    foreach ($entry in @($archive.Entries | Where-Object { $_.FullName.EndsWith('.md', [StringComparison]::OrdinalIgnoreCase) })) {
        $required = $entry.FullName
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $markdown = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        # Relative repository links must also work in the extracted package.
        # Check every shipped guide, including their transitive local links.
        foreach ($link in [regex]::Matches($markdown, '(?<!!)\[[^\]]+\]\(([^)\s]+)\)')) {
            $target = $link.Groups[1].Value
            if ($target.StartsWith('#') -or $target -match '^[a-zA-Z][a-zA-Z0-9+.-]*:') { continue }
            $resolved = [Uri]::new([Uri]('https://package.invalid/' + $required), $target).AbsolutePath.TrimStart('/')
            if ($null -eq $archive.GetEntry($resolved)) { throw "Packed guide $required links to absent document $resolved" }
            $localDocumentationLinks++
        }
    }
    $libraryEntry = $archive.GetEntry("lib/$Framework/Lokad.OoxPdf.dll")
    if ($null -eq $libraryEntry) { throw "Packed library DLL is missing." }
    $libraryStream = $libraryEntry.Open()
    try {
        $librarySha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($libraryStream))
    }
    finally { $libraryStream.Dispose() }
}
finally { $archive.Dispose() }
if ($ExpectedLibrarySha256 -and $librarySha256 -ne $ExpectedLibrarySha256) {
    throw "Packed library differs from the qualified artifact."
}

Copy-Item -LiteralPath (Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/docx-ladder-01-plain-paragraph.docx") -Destination (Join-Path $scratch "input.docx")
Copy-Item -LiteralPath (Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-10-composite-table-port.pptx") -Destination (Join-Path $scratch "input.pptx")

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
    <TargetFramework>$Framework</TargetFramework>
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
using System.Reflection;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

string input = args[0];
string output = args[1];
string expectedLibrarySha256 = args[2];
string packageVersion = args[3];
Assembly library = typeof(OoxPdfConverter).Assembly;
string informationalVersion = library.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;
string fileVersion = library.GetCustomAttribute<AssemblyFileVersionAttribute>()!.Version;
Version numericVersion = new Version(Version.Parse(packageVersion.Split('-')[0]).Major,
    Version.Parse(packageVersion.Split('-')[0]).Minor, Version.Parse(packageVersion.Split('-')[0]).Build, 0);
if (library.GetName().Version != numericVersion || fileVersion != numericVersion.ToString() ||
    informationalVersion.Split('+')[0] != packageVersion)
{
    throw new InvalidDataException("Package, assembly, file and informational versions disagree.");
}
string loadedLibrarySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(OoxPdfConverter).Assembly.Location)));
if (!string.Equals(expectedLibrarySha256, loadedLibrarySha256, StringComparison.Ordinal))
{
    throw new InvalidDataException("Smoke consumer loaded a DLL different from the freshly packed library.");
}
Console.WriteLine("Loaded packed library SHA-256 " + loadedLibrarySha256 + ".");
var admission = new SmokeAdmission();
var options = new OoxPdfOptions
{
    FontResolver = args[4].Length != 0 ? new SmokeFontResolver(args[4]) : null,
    MemoryAdmission = admission
};
OoxPdfConverter.Convert(input, output, options);
if (admission.Requests != 1 || admission.Releases != 1) { throw new InvalidOperationException("Packed admission ownership is unbalanced."); }
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
string pptxOutput = Path.ChangeExtension(output, ".pptx.pdf");
OoxPdfConverter.Convert(args[5], pptxOutput, options);
string pptxPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pptxOutput));
if (!pptxPdf.StartsWith("%PDF-", StringComparison.Ordinal) || !pptxPdf.Contains("/FontFile2", StringComparison.Ordinal))
    throw new InvalidDataException("Packed PPTX conversion must produce a PDF with embedded text.");
if (admission.Requests != 2 || admission.Releases != 2)
    throw new InvalidOperationException("Packed DOCX/PPTX admission ownership is unbalanced.");

string protectedOutput = Path.ChangeExtension(output, ".protected.pdf");
const string retained = "Retain the existing destination after failure.";
File.WriteAllText(protectedOutput, retained);
using (var cancellation = new CancellationTokenSource())
{
    cancellation.Cancel();
    var error = ExpectFailure<OperationCanceledException>(() => OoxPdfConverter.Convert(input, protectedOutput, options, cancellation.Token));
    if (error.CancellationToken != cancellation.Token)
        throw new InvalidOperationException("Packed cancellation must retain the caller token.");
}
string malformed = Path.ChangeExtension(input, ".malformed.docx");
File.WriteAllText(malformed, "This is not an OOXML ZIP package.");
ExpectFailure<InvalidDataException>(() => OoxPdfConverter.Convert(malformed, protectedOutput, options));
var refusal = new RefusingAdmission();
ExpectFailure<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(args[5], protectedOutput,
    new OoxPdfOptions { FontResolver = options.FontResolver, MemoryAdmission = refusal }));
if (refusal.Requests != 1 || admission.Requests != admission.Releases || File.ReadAllText(protectedOutput) != retained)
    throw new InvalidOperationException("Packed failure must balance ownership and preserve the destination.");
Console.WriteLine("Packed PPTX conversion, cancellation, malformed-input recovery, admission refusal and destination preservation passed.");

static T ExpectFailure<T>(Action action) where T : Exception
{
    try { action(); }
    catch (T error) { return error; }
    throw new InvalidOperationException("Expected packed failure: " + typeof(T).Name);
}

sealed class RefusingAdmission : IOoxPdfMemoryAdmission
{
    public int Requests;
    public IDisposable? TryAcquire(CancellationToken token) { token.ThrowIfCancellationRequested(); Requests++; return null; }
}

sealed class SmokeAdmission : IOoxPdfMemoryAdmission, IDisposable
{
    public int Requests;
    public int Releases;
    public IDisposable? TryAcquire(CancellationToken token) { token.ThrowIfCancellationRequested(); Requests++; return this; }
    public void Dispose() { Releases++; }
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
$smokeArguments = @($input, $output, $librarySha256, $version, $FontPath, (Join-Path $scratch 'input.pptx'))
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
    Framework = $Framework
    LibraryRebuilt = -not $NoBuild -and -not $PackagePath
    ReusedPackage = [bool]$PackagePath
    ExpectedLibrarySha256 = $ExpectedLibrarySha256
    MemoryAdmissionChecked = $true
    PptxConversionChecked = $true
    CancellationChecked = $true
    MalformedInputChecked = $true
    AdmissionRefusalChecked = $true
    FailurePreservesDestinationChecked = $true
    VersionMetadataChecked = $true
    CallerDocumentationChecked = $true
    LocalDocumentationLinksChecked = $localDocumentationLinks
    Package = $package.FullName
    PackageVersion = $version
    PackageSha256 = (Get-FileHash -LiteralPath $package.FullName -Algorithm SHA256).Hash
    LibrarySha256 = $librarySha256
    OutputSha256 = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
    OutputBytes = (Get-Item -LiteralPath $output).Length
    PptxOutputSha256 = (Get-FileHash -LiteralPath ([IO.Path]::ChangeExtension($output, '.pptx.pdf')) -Algorithm SHA256).Hash
    RestorePackages = Join-Path $scratch "packages"
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $scratch "summary.json") -Encoding UTF8
Write-Host ("Evidence: {0}" -f (Join-Path $scratch "summary.json"))
