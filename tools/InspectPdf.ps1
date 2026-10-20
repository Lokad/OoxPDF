param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [string] $OutputDirectory,

    [switch] $TextOnly,

    [int[]] $Page,
    [string] $InspectorPath,
    [switch] $SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj"
$dll = if ($InspectorPath) { (Resolve-Path $InspectorPath).Path } else { Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/net10.0/Lokad.OoxPdf.PdfInspect.dll" }
$buildRecordPath = $null
if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $buildRecordPath = Join-Path $OutputDirectory 'pdf-inspect-build-info.json'
}
if (!$InspectorPath -and !$SkipBuild) {
    & (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $project -OutputDll $dll -Description 'PDF inspect' -RecordPath $buildRecordPath
}
$inspectorHash = (Get-FileHash $dll).Hash
if (($InspectorPath -or $SkipBuild) -and $buildRecordPath) {
    @{OutputDll=$dll;OutputSha256=$inspectorHash;BuildRequested=$false} | ConvertTo-Json | Set-Content $buildRecordPath
}

$arguments = @((Resolve-Path -LiteralPath $InputPdf).Path)
if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $arguments += $OutputDirectory
}

if ($TextOnly) {
    $arguments += "--text-only"
}
foreach ($pageNumber in $Page) {
    $arguments += "--page"
    $arguments += $pageNumber.ToString([Globalization.CultureInfo]::InvariantCulture)
}

dotnet $dll @arguments

if ($LASTEXITCODE -ne 0) {
    throw "PDF inspect failed with exit code $LASTEXITCODE."
}
if ((Get-FileHash $dll).Hash -ne $inspectorHash) { throw 'PDF inspection runtime changed during execution' }
