param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144
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
$sourceNewest = Get-ChildItem -LiteralPath (Split-Path -Parent $rasterizerProject) -Recurse -Include *.cs,*.csproj |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
if (-not (Test-Path -LiteralPath $rasterizerDll) -or $sourceNewest.LastWriteTimeUtc -gt (Get-Item -LiteralPath $rasterizerDll).LastWriteTimeUtc) {
    dotnet build $rasterizerProject --tl:off --nologo -v minimal
    if ($LASTEXITCODE -ne 0) {
        throw "PDFium rasterizer build failed with exit code $LASTEXITCODE."
    }
}

dotnet $rasterizerDll $inputFull $outputFull $Dpi
if ($LASTEXITCODE -ne 0) {
    throw "PDFium rasterizer failed with exit code $LASTEXITCODE."
}
