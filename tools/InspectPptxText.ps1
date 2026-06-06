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
$sourceRoots = @(
    (Split-Path -Parent $project),
    (Join-Path $repoRoot "src/Lokad.OoxPdf")
)
$sourceNewest = $sourceRoots |
    ForEach-Object { Get-ChildItem -LiteralPath $_ -Recurse -Include *.cs,*.csproj } |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
if (-not (Test-Path -LiteralPath $dll) -or $sourceNewest.LastWriteTimeUtc -gt (Get-Item -LiteralPath $dll).LastWriteTimeUtc) {
    dotnet build $project --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "PPTX text inspect build failed with exit code $LASTEXITCODE."
    }
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

dotnet $dll @arguments
if ($LASTEXITCODE -ne 0) {
    throw "PPTX text inspect failed with exit code $LASTEXITCODE."
}
