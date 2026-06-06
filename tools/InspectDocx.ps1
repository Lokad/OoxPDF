param(
    [Parameter(Mandatory = $true)]
    [string] $InputDocx,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.DocxInspect/Lokad.OoxPdf.DocxInspect.csproj"
$dll = Join-Path $repoRoot "tools/Lokad.OoxPdf.DocxInspect/bin/Debug/net10.0/Lokad.OoxPdf.DocxInspect.dll"
dotnet build $project --nologo
if ($LASTEXITCODE -ne 0) {
    throw "DOCX inspect build failed with exit code $LASTEXITCODE."
}

dotnet $dll (Resolve-Path -LiteralPath $InputDocx).Path $OutputDirectory
if ($LASTEXITCODE -ne 0) {
    throw "DOCX inspect failed with exit code $LASTEXITCODE."
}
