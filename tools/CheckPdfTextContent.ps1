# Source-text preservation gate (P06 slice 1): converts each case document
# with the CLI, extracts candidate text with the independent PdfInspect
# decoder (own PDF parser, own ToUnicode maps; no Office needed), and checks
# hand-authored expected strings appear in order. Expectations live beside
# the inputs they describe, never generated from candidate output.
param(
    [string] $CasesDir = "",
    [string] $Out = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($CasesDir)) {
    $CasesDir = Join-Path $repoRoot "tools/text-content-cases"
}
if ([string]::IsNullOrWhiteSpace($Out)) {
    $Out = Join-Path $repoRoot "artifacts/text-content"
}
New-Item -ItemType Directory -Force -Path $Out | Out-Null

$cliProject = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj"
$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.Cli.dll"
$buildInfoDir = Join-Path $Out "build-info"
New-Item -ItemType Directory -Force -Path $buildInfoDir | Out-Null
& (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Project $cliProject -OutputDll $cliDll -Description "CLI" -Configuration Release -RecordPath (Join-Path $buildInfoDir "cli.json")

$failures = 0
$cases = Get-ChildItem -LiteralPath $CasesDir -Filter *.json | Sort-Object Name
if ($cases.Count -eq 0) { throw "No text-content cases in $CasesDir." }
foreach ($caseFile in $cases) {
    $case = Get-Content -Raw -LiteralPath $caseFile.FullName | ConvertFrom-Json
    $input = Join-Path $repoRoot $case.input
    if (-not (Test-Path -LiteralPath $input)) { throw "Missing input $($case.input)." }
    $caseDir = Join-Path $Out ([IO.Path]::GetFileNameWithoutExtension($caseFile.Name))
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null
    $pdf = Join-Path $caseDir "output.pdf"
    $inspectDir = Join-Path $caseDir "inspect"
    Write-Host "[$($caseFile.BaseName)] converting..."
    & dotnet $cliDll convert $input $pdf
    if ($LASTEXITCODE -ne 0) { throw "Conversion failed for $($case.input)." }
    & (Join-Path $repoRoot "tools/InspectPdf.ps1") -InputPdf $pdf -OutputDirectory $inspectDir -TextOnly
    $ops = Get-Content -Raw -LiteralPath (Join-Path $inspectDir "text-operations.json") | ConvertFrom-Json
    $pages = $ops | Group-Object { $_.PageNumber } | Sort-Object { [int]$_.Name }
    $fullText = ($pages | ForEach-Object { ($_.Group | ForEach-Object { $_.DecodedText }) -join "" }) -join "`n"
    $cursor = 0
    foreach ($expected in @($case.expected)) {
        $found = $fullText.IndexOf($expected, $cursor, [StringComparison]::Ordinal)
        if ($found -lt 0) {
            $failures++
            Write-Host "[$($caseFile.BaseName)] MISSING: $expected"
        }
        else { $cursor = $found + $expected.Length }
    }
    if ($failures -eq 0) { Write-Host "[$($caseFile.BaseName)] ok" }
}
if ($failures -ne 0) { throw "$failures expected string(s) missing from extracted text." }
Write-Host "Text-content gate passed for $($cases.Count) case(s)."
