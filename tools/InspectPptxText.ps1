param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $InputPptx,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int[]] $Slide,

    [switch] $IncludeText,
    [string] $InspectorPath,
    [switch] $SkipBuild,
    [string] $ExpectedLibrarySha256,
    [string] $FontDirectory
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.PptxInspect/Lokad.OoxPdf.PptxInspect.csproj"
$dll = if ($InspectorPath) { (Resolve-Path $InspectorPath).Path } else { Join-Path $repoRoot "tools/Lokad.OoxPdf.PptxInspect/bin/Debug/$Framework/Lokad.OoxPdf.PptxInspect.dll" }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$buildRecordPath = Join-Path $OutputDirectory 'pptx-inspect-build-info.json'
if (!$InspectorPath -and !$SkipBuild) {
    & (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Framework $Framework -Project $project -OutputDll $dll -Description 'PPTX text inspect' -RecordPath $buildRecordPath
}
$libraryPath = Join-Path (Split-Path $dll) 'Lokad.OoxPdf.dll'
$inspectorHash = (Get-FileHash $dll).Hash
$libraryHash = (Get-FileHash $libraryPath).Hash
if ($ExpectedLibrarySha256 -and $ExpectedLibrarySha256 -ne $libraryHash) { throw 'PPTX inspector library differs from the qualified converter' }
if ($InspectorPath -or $SkipBuild) {
    @{OutputDll=$dll;OutputSha256=$inspectorHash;LibrarySha256=$libraryHash;BuildRequested=$false} | ConvertTo-Json | Set-Content $buildRecordPath
}

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
if ($FontDirectory) {
    $arguments += '--font-directory'
    $arguments += (Resolve-Path -LiteralPath $FontDirectory).Path
}

dotnet $dll @arguments
if ($LASTEXITCODE -ne 0) {
    throw "PPTX text inspect failed with exit code $LASTEXITCODE."
}
if ((Get-FileHash $dll).Hash -ne $inspectorHash -or (Get-FileHash $libraryPath).Hash -ne $libraryHash) { throw 'PPTX inspection runtime changed during execution' }
