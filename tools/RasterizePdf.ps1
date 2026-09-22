param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $SkipBuild,

    # R21: bound native validation time from the outside. PDFium calls are
    # synchronous, so an over-long render is killed at the process boundary
    # instead of hanging the validation gate.
    [int] $TimeoutSeconds = 600
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pdfium = Join-Path $PSScriptRoot "vendor/pdfium/win-x64/bin/pdfium.dll"
if (-not (Test-Path -LiteralPath $pdfium)) {
    throw "Missing PDFium DLL: $pdfium. Retrieve https://github.com/bblanchon/pdfium-binaries/releases/latest/download/pdfium-win-x64.tgz and unpack it under tools/vendor/pdfium/win-x64."
}

$pinFile = Join-Path $PSScriptRoot "vendor/pdfium/win-x64/pdfium.sha256"
$pinnedHash = @((Get-Content -LiteralPath $pinFile) | Where-Object { $_ -match "^[0-9a-fA-F]{64}" } | ForEach-Object { ($_.Substring(0, 64).ToLowerInvariant()) })
if ($pinnedHash.Count -ne 1) {
    throw "PDFium pin file is missing or ambiguous."
}
$actualHash = (Get-FileHash -LiteralPath $pdfium -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualHash -ne $pinnedHash[0]) {
    throw "PDFium binary does not match the pinned hash."
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

$rasterizer = Start-Process -FilePath "dotnet" -ArgumentList @($rasterizerDll, $inputFull, $outputFull, $Dpi) -NoNewWindow -PassThru
if (-not $rasterizer.WaitForExit($TimeoutSeconds * 1000)) {
    try { $rasterizer.Kill() } catch { }
    throw "PDFium rasterizer timed out after $TimeoutSeconds seconds on $InputPdf."
}
if ($rasterizer.ExitCode -ne 0) {
    throw "PDFium rasterizer failed with exit code $($rasterizer.ExitCode)."
}
