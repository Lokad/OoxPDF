# Dumps DOCX layout/mode inspection JSON (helper for markup triage).

param(
    [Parameter(Mandatory = $true)]
    [string] $InputDocx,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [ValidateSet("final", "original", "simple", "all", "simple-markup", "all-markup")]
    [string] $DocxMarkup = "final",

    [ValidateSet("preserve", "preserve-layout", "preserve-document-layout", "reserve", "reserve-margin", "markup-margin", "reserve-markup-margin", "word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup")]
    [string] $DocxMarkupGeometry = "preserve",
    [string] $InspectorPath,
    [switch] $SkipBuild,
    [string] $ExpectedLibrarySha256
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.DocxInspect/Lokad.OoxPdf.DocxInspect.csproj"
$dll = if ($InspectorPath) { (Resolve-Path -LiteralPath $InspectorPath).Path } else {
    Join-Path $repoRoot "tools/Lokad.OoxPdf.DocxInspect/bin/Debug/net10.0/Lokad.OoxPdf.DocxInspect.dll"
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$buildRecordPath = Join-Path $OutputDirectory 'docx-inspect-build-info.json'
if (!$InspectorPath -and !$SkipBuild) {
    & (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $project -OutputDll $dll -Description 'DOCX inspect' -RecordPath $buildRecordPath
}
$libraryPath = Join-Path (Split-Path $dll) 'Lokad.OoxPdf.dll'
$inspectorHash = (Get-FileHash -LiteralPath $dll).Hash
$libraryHash = (Get-FileHash -LiteralPath $libraryPath).Hash
if ($ExpectedLibrarySha256 -and $libraryHash -ne $ExpectedLibrarySha256) { throw 'DOCX inspector library differs from the qualified converter' }
if ($InspectorPath -or $SkipBuild) {
    @{OutputDll=$dll;OutputSha256=$inspectorHash;LibrarySha256=$libraryHash;BuildRequested=$false} |
        ConvertTo-Json | Set-Content -LiteralPath $buildRecordPath
}

dotnet $dll (Resolve-Path -LiteralPath $InputDocx).Path $OutputDirectory --docx-markup $DocxMarkup --docx-markup-geometry $DocxMarkupGeometry
if ($LASTEXITCODE -ne 0) {
    throw "DOCX inspect failed with exit code $LASTEXITCODE."
}
if ((Get-FileHash $dll).Hash -ne $inspectorHash -or (Get-FileHash $libraryPath).Hash -ne $libraryHash) { throw 'DOCX inspection runtime changed during execution' }
