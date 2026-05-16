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

& (Join-Path $PSScriptRoot "RenderCachedReference.ps1") -InputPath $inputFull -OutputDirectory $referenceDir -Dpi $dpi
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

$metricsPath = Join-Path $comparisonDir "metrics.json"
$metrics = Get-Content -Raw -LiteralPath $metricsPath | ConvertFrom-Json
$referencePages = @(Get-ChildItem -LiteralPath $referenceDir -Filter "page-*.png")
$candidatePages = @(Get-ChildItem -LiteralPath $candidateDir -Filter "page-*.png")

if ($manifest.expected.pageCountMustMatch -eq $true -and $referencePages.Count -ne $candidatePages.Count) {
    throw "Page count mismatch. Reference: $($referencePages.Count). Candidate: $($candidatePages.Count)."
}

if ($manifest.expected.dimensionsMustMatch -eq $true) {
    $mismatched = @($metrics | Where-Object { $_.DimensionsMatch -ne $true })
    if ($mismatched.Count -ne 0) {
        $pages = ($mismatched | ForEach-Object { $_.Page }) -join ", "
        throw "Dimension mismatch on page(s): $pages."
    }
}

if ($manifest.expected.maxMeanAbsoluteError -ne $null) {
    $maxMae = [double]$manifest.expected.maxMeanAbsoluteError
    $exceeded = @($metrics | Where-Object { [double]$_.MeanAbsoluteError -gt $maxMae })
    if ($exceeded.Count -ne 0) {
        $worst = $exceeded | Sort-Object -Property MeanAbsoluteError -Descending | Select-Object -First 1
        throw "Mean absolute error gate failed. Page $($worst.Page) was $($worst.MeanAbsoluteError), limit is $maxMae."
    }
}

if ($manifest.expected.maxChangedPixelRatioAtThreshold16 -ne $null) {
    $maxChanged = [double]$manifest.expected.maxChangedPixelRatioAtThreshold16
    $exceeded = @($metrics | Where-Object { [double]$_.ChangedPixelRatioAtThreshold16 -gt $maxChanged })
    if ($exceeded.Count -ne 0) {
        $worst = $exceeded | Sort-Object -Property ChangedPixelRatioAtThreshold16 -Descending | Select-Object -First 1
        throw "Changed-pixel ratio gate failed. Page $($worst.Page) was $($worst.ChangedPixelRatioAtThreshold16), limit is $maxChanged."
    }
}

$diagnosticsJson = Get-Content -Raw -LiteralPath $diagnostics
$diagnosticItems = if ($diagnosticsJson.Trim() -eq "[]") {
    @()
}
else {
    @(ConvertFrom-Json -InputObject $diagnosticsJson)
}
if ($manifest.expected.diagnosticsMustBeEmpty -eq $true -and $diagnosticItems.Count -ne 0) {
    $ids = ($diagnosticItems | ForEach-Object { $_.Id } | Sort-Object -Unique) -join ", "
    throw "Expected no diagnostics, but found: $ids."
}

$allowedUnsupportedFeatures = @($manifest.allowedUnsupportedFeatures)
if ($allowedUnsupportedFeatures.Count -ne 0 -and $diagnosticItems.Count -ne 0) {
    $unexpectedDiagnostics = @($diagnosticItems | Where-Object {
        $allowedUnsupportedFeatures -notcontains $_.Id -and
            $allowedUnsupportedFeatures -notcontains $_.Feature
    })
    if ($unexpectedDiagnostics.Count -ne 0) {
        $ids = ($unexpectedDiagnostics | ForEach-Object { $_.Id } | Sort-Object -Unique) -join ", "
        throw "Unexpected diagnostics: $ids."
    }
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
