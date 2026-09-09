param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pdfium = Join-Path $PSScriptRoot "vendor/pdfium/win-x64/bin/pdfium.dll"
if (-not (Test-Path -LiteralPath $pdfium)) {
    throw "Missing PDFium DLL: $pdfium. Retrieve https://github.com/bblanchon/pdfium-binaries/releases/latest/download/pdfium-win-x64.tgz and unpack it under tools/vendor/pdfium/win-x64."
}

$inputFull = (Resolve-Path -LiteralPath $InputPdf).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path

$rasterizerProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj"
$rasterizerDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/net10.0/Lokad.OoxPdf.PdfiumRasterizer.dll"
if ($SkipBuild) {
    if (-not (Test-Path -LiteralPath $rasterizerDll)) {
        throw "PDFium rasterizer DLL is missing and -SkipBuild was requested: $rasterizerDll."
    }
}
else {
    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Project $rasterizerProject -OutputDll $rasterizerDll -Description "PDFium rasterizer"
}

dotnet $rasterizerDll $inputFull $outputFull $Dpi
if ($LASTEXITCODE -ne 0) {
    throw "PDFium rasterizer failed with exit code $LASTEXITCODE."
}
