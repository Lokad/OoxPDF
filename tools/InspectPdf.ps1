param(
    [Parameter(Mandatory = $true)]
    [string] $InputPdf,

    [string] $OutputDirectory
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj"
dotnet build $project --tl:off --nologo -v minimal
if ($LASTEXITCODE -ne 0) {
    throw "PDF inspect build failed with exit code $LASTEXITCODE."
}

$dll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/net10.0/Lokad.OoxPdf.PdfInspect.dll"
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    dotnet $dll (Resolve-Path -LiteralPath $InputPdf).Path
}
else {
    dotnet $dll (Resolve-Path -LiteralPath $InputPdf).Path $OutputDirectory
}

if ($LASTEXITCODE -ne 0) {
    throw "PDF inspect failed with exit code $LASTEXITCODE."
}
