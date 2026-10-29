param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $SkipBuild,

    # R21: bound native validation time from the outside. PDFium calls are
    # synchronous, so an over-long render is killed at the process boundary
    # instead of hanging the validation gate.
    [int] $TimeoutSeconds = 600,
    [string] $RasterizerPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pdfium = Join-Path $PSScriptRoot "vendor/pdfium/win-x64/bin/pdfium.dll"
if (-not (Test-Path -LiteralPath $pdfium)) {
    throw "Missing PDFium DLL: $pdfium. Restore the tracked pinned binary; see tools/vendor/pdfium/win-x64/VERSION, pdfium.sha256 and docs/VisualValidation.md."
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
if ((Test-Path -LiteralPath $OutputDirectory -PathType Container) -and
    @(Get-ChildItem -LiteralPath $OutputDirectory -Filter '*.png' -File).Count) {
    throw 'Raster output already contains PNG images; choose a fresh directory.'
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path

$rasterizerProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj"
$rasterizerDll = if ($RasterizerPath) { (Resolve-Path -LiteralPath $RasterizerPath).Path }
    else { Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework/Lokad.OoxPdf.PdfiumRasterizer.dll" }
if ($SkipBuild -or $RasterizerPath) {
    if (-not (Test-Path -LiteralPath $rasterizerDll)) {
        throw "PDFium rasterizer DLL is missing and -SkipBuild was requested: $rasterizerDll."
    }
}
else {
    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Framework $Framework -Project $rasterizerProject -OutputDll $rasterizerDll -Description "PDFium rasterizer"
}

$start = [Diagnostics.ProcessStartInfo]::new("dotnet")
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
foreach ($argument in @($rasterizerDll, $inputFull, $outputFull, $Dpi.ToString([Globalization.CultureInfo]::InvariantCulture))) {
    $start.ArgumentList.Add($argument)
}
$rasterizer = [Diagnostics.Process]::Start($start)
try {
    if (-not $rasterizer.WaitForExit($TimeoutSeconds * 1000)) {
        try {
            $rasterizer.Kill($true)
            $null = $rasterizer.WaitForExit(5000)
        } catch { }
        throw "PDFium rasterizer timed out after $TimeoutSeconds seconds on $InputPdf."
    }
    if ($rasterizer.ExitCode -ne 0) {
        throw "PDFium rasterizer failed with exit code $($rasterizer.ExitCode)."
    }
}
finally {
    $rasterizer.Dispose()
}
