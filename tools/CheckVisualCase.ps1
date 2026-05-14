param(
    [Parameter(Mandatory = $true)]
    [string] $Case
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$caseFull = (Resolve-Path -LiteralPath $Case).Path
$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$runRoot = Join-Path $repoRoot ("artifacts/visual/{0}/{1}" -f $manifest.id, $runId)
$referenceDir = Join-Path $runRoot "reference"
$candidateDir = Join-Path $runRoot "candidate"
$comparisonDir = Join-Path $runRoot "comparison"

New-Item -ItemType Directory -Force -Path $referenceDir, $candidateDir, $comparisonDir | Out-Null

$inputPath = Join-Path $caseDirectory $manifest.input
$inputFull = (Resolve-Path -LiteralPath $inputPath).Path
$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
$dpi = if ($manifest.dpi -ne $null) { [int]$manifest.dpi } else { 144 }

dotnet build (Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj") --tl:off --nologo -v minimal
if ($LASTEXITCODE -ne 0) {
    throw "CLI build failed with exit code $LASTEXITCODE."
}

$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
dotnet $cliDll convert $inputFull $candidatePdf --diagnostics $diagnostics
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "RenderReference.ps1") -InputPath $inputFull -OutputDirectory $referenceDir -Dpi $dpi
& (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $candidatePdf -OutputDirectory $candidateDir -Dpi $dpi

dotnet build (Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj") --tl:off --nologo -v minimal
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff build failed with exit code $LASTEXITCODE."
}

$visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/net10.0/Lokad.OoxPdf.VisualDiff.dll"
dotnet $visualDiffDll $referenceDir $candidateDir $comparisonDir
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff failed with exit code $LASTEXITCODE."
}

$assessment = @"
# Visual assessment: $($manifest.id)

Input: $inputFull
Kind: $($manifest.kind)
Run: $runId
Agent rating: <0-5>

## Summary

<One paragraph summary of visual fidelity.>

## Pages/slides reviewed

- Page 1: <rating>, <main differences>

## Major defects

1. <Defect, page, suspected feature>

## Diagnostics reviewed

<Note important warnings/errors from diagnostics.json.>

## Next implementation target

<The smallest renderer improvement likely to improve this case.>
"@

Set-Content -LiteralPath (Join-Path $runRoot "assessment.md") -Value $assessment
Write-Host "Visual case artifacts: $runRoot"
