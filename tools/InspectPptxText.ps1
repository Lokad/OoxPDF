param(
    [Parameter(Mandatory = $true)]
    [string] $InputPptx,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int[]] $Slide,

    [switch] $IncludeText
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.PptxInspect/Lokad.OoxPdf.PptxInspect.csproj"
$dll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PptxInspect/bin/Debug/net10.0/Lokad.OoxPdf.PptxInspect.dll"
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$buildRecordPath = Join-Path $OutputDirectory 'pptx-inspect-build-info.json'
& (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $project -OutputDll $dll -Description 'PPTX text inspect' -RecordPath $buildRecordPath

$arguments = @(
    (Resolve-Path -LiteralPath $InputPptx).Path,
    $OutputDirectory
)
foreach ($slideNumber in $Slide) {
    $arguments += "--slide"
    $arguments += $slideNumber.ToString([Globalization.CultureInfo]::InvariantCulture)
}
if ($IncludeText) {
    $arguments += "--include-text"
}

dotnet $dll @arguments
if ($LASTEXITCODE -ne 0) {
    throw "PPTX text inspect failed with exit code $LASTEXITCODE."
}
