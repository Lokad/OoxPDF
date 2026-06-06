param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [string] $OutputDirectory,

    [switch] $TextOnly,

    [int[]] $Page
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj"
$dll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/net10.0/Lokad.OoxPdf.PdfInspect.dll"
$sourceNewest = Get-ChildItem -LiteralPath (Split-Path -Parent $project) -Recurse -Include *.cs,*.csproj |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
if (-not (Test-Path -LiteralPath $dll) -or $sourceNewest.LastWriteTimeUtc -gt (Get-Item -LiteralPath $dll).LastWriteTimeUtc) {
    dotnet build $project --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "PDF inspect build failed with exit code $LASTEXITCODE."
    }
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

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    dotnet $dll @arguments
}
else {
    dotnet $dll @arguments
}

if ($LASTEXITCODE -ne 0) {
    throw "PDF inspect failed with exit code $LASTEXITCODE."
}
