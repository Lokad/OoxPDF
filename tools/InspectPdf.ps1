param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [string] $OutputDirectory
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
    dotnet build $project --tl:off --nologo -v minimal
    if ($LASTEXITCODE -ne 0) {
        throw "PDF inspect build failed with exit code $LASTEXITCODE."
    }
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    dotnet $dll (Resolve-Path -LiteralPath $InputPdf).Path
}
else {
    dotnet $dll (Resolve-Path -LiteralPath $InputPdf).Path $OutputDirectory
}

if ($LASTEXITCODE -ne 0) {
    throw "PDF inspect failed with exit code $LASTEXITCODE."
}
