# Foreground-recall probe (T06 slice 1): VisualDiff reports the share of
# reference foreground pixels the candidate reproduces (null when the
# reference has no foreground, i.e. the metric does not apply). The legacy
# foreground histogram correlation returns 1 whenever the smaller foreground
# side is below 1.5% page occupancy, and reports 1 (both sides empty) or 0
# (one side empty) where there is nothing to correlate. No Office/COM or
# reference cache is needed: both PDFs are converted from tracked synthetic
# fixtures with the CLI and rasterized with the local PDFium rasterizer.

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = Join-Path $repoRoot 'artifacts/visual-diff-metrics-probe'
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $scratch | Out-Null
$buildInfoDir = Join-Path $scratch 'build-info'
New-Item -ItemType Directory -Force -Path $buildInfoDir | Out-Null

$cliProject = Join-Path $repoRoot 'src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj'
$cliDll = Join-Path $repoRoot 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.Cli.dll'
& (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $cliProject -OutputDll $cliDll -Description 'CLI' -Configuration Release -RecordPath (Join-Path $buildInfoDir 'cli.json')

$rasterizerProject = Join-Path $repoRoot 'tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj'
$rasterizerDll = Join-Path $repoRoot 'tools/Lokad.OoxPdf.PdfiumRasterizer/bin/Release/net10.0/Lokad.OoxPdf.PdfiumRasterizer.dll'
& (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $rasterizerProject -OutputDll $rasterizerDll -Description 'PDFium rasterizer' -Configuration Release

$visualDiffProject = Join-Path $repoRoot 'tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj'
$visualDiffDll = Join-Path $repoRoot 'tools/Lokad.OoxPdf.VisualDiff/bin/Release/net10.0/Lokad.OoxPdf.VisualDiff.dll'
& (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1') -Project $visualDiffProject -OutputDll $visualDiffDll -Description 'VisualDiff' -Configuration Release

$blankDocx = Join-Path $repoRoot 'tests/Lokad.OoxPdf.Tests/Cases/docx-ladder-00-blank.docx'
$titledDocx = Join-Path $repoRoot 'tests/Lokad.OoxPdf.Tests/Cases/docx-ladder-01-plain-paragraph.docx'
foreach ($fixture in @($blankDocx, $titledDocx)) {
    if (-not (Test-Path -LiteralPath $fixture)) {
        throw ('Missing probe fixture: ' + $fixture)
    }
}

$blankPdf = Join-Path $scratch 'blank.pdf'
$titledPdf = Join-Path $scratch 'titled.pdf'
& dotnet $cliDll convert $blankDocx $blankPdf
if ($LASTEXITCODE -ne 0) {
    throw 'Probe conversion failed for the blank fixture.'
}
& dotnet $cliDll convert $titledDocx $titledPdf
if ($LASTEXITCODE -ne 0) {
    throw 'Probe conversion failed for the titled fixture.'
}

$blankPng = Join-Path $scratch 'png-blank'
$titledPng = Join-Path $scratch 'png-titled'
New-Item -ItemType Directory -Force -Path $blankPng, $titledPng | Out-Null
& dotnet $rasterizerDll $blankPdf $blankPng 144
if ($LASTEXITCODE -ne 0) {
    throw 'Probe rasterization failed for the blank PDF.'
}
& dotnet $rasterizerDll $titledPdf $titledPng 144
if ($LASTEXITCODE -ne 0) {
    throw 'Probe rasterization failed for the titled PDF.'
}

function Read-ProbePage([string] $Reference, [string] $Candidate, [string] $Name, [string] $Regions = '') {
    $out = Join-Path $scratch $Name
    New-Item -ItemType Directory -Force -Path $out | Out-Null
    if ([string]::IsNullOrWhiteSpace($Regions)) {
        & dotnet $visualDiffDll $Reference $Candidate $out
    }
    else {
        & dotnet $visualDiffDll $Reference $Candidate $out $Regions
    }
    if ($LASTEXITCODE -ne 0) {
        throw ('VisualDiff failed for ' + $Name)
    }
    $metrics = @(Get-Content -Raw -LiteralPath (Join-Path $out 'metrics.json') | ConvertFrom-Json)
    if ($metrics.Count -ne 1) {
        throw ('Expected one compared page in ' + $Name)
    }
    return [pscustomobject]@{ Page = $metrics[0]; Directory = $out }
}

$missing = Read-ProbePage $titledPng $blankPng 'comparison-missing'
if ($missing.Page.ForegroundRecall -ne 0) {
    throw ('Missing-title recall must be 0, was ' + $missing.Page.ForegroundRecall)
}
if ($missing.Page.ForegroundColorHistogramCorrelation -ne 0) {
    throw 'Probe premise broken: expected histogram 0 (empty candidate side).'
}
Write-Host 'PASS: missing title -> ForegroundRecall=0 (lost sparse content is not averaged away).'

$emptyRef = Read-ProbePage $blankPng $titledPng 'comparison-empty-ref'
if ($null -ne $emptyRef.Page.ForegroundRecall) {
    throw 'Empty-reference recall must be null (the metric does not apply).'
}
if ($emptyRef.Page.ForegroundColorHistogramCorrelation -ne 0) {
    throw 'Probe premise broken: expected histogram 0 for the empty reference side.'
}
Write-Host 'PASS: empty reference -> ForegroundRecall=null while histogram reports 0.'

$identical = Read-ProbePage $titledPng $titledPng 'comparison-identical'
if ($identical.Page.ForegroundRecall -ne 1) {
    throw 'Identical-page recall must be 1.'
}
Write-Host 'PASS: identical pages -> ForegroundRecall=1.'

$blankBlank = Read-ProbePage $blankPng $blankPng 'comparison-blank-blank'
if ($null -ne $blankBlank.Page.ForegroundRecall) {
    throw 'Blank-vs-blank recall must be null (the metric does not apply).'
}
if ($blankBlank.Page.ForegroundColorHistogramCorrelation -ne 1) {
    throw 'Probe premise broken: expected histogram 1 for two empty pages.'
}
Write-Host 'PASS: blank vs blank -> ForegroundRecall=null while histogram reports 1.'

$regionSpecs = Join-Path $scratch 'regions.json'
Set-Content -LiteralPath $regionSpecs -Encoding UTF8 -Value '[{"page":1,"region":"full","xRatio":0,"yRatio":0,"widthRatio":1,"heightRatio":1},{"page":1,"region":"bottom-half","xRatio":0,"yRatio":0.5,"widthRatio":1,"heightRatio":0.5}]'

$missingRegions = Read-ProbePage $titledPng $blankPng 'comparison-missing-regions' $regionSpecs
$missingRegionMetrics = @(Get-Content -Raw -LiteralPath (Join-Path $missingRegions.Directory 'region-metrics.json') | ConvertFrom-Json)
$full = @($missingRegionMetrics | Where-Object { $_.Region -eq 'full' })
$bottomMissing = @($missingRegionMetrics | Where-Object { $_.Region -eq 'bottom-half' })
if ($full.Count -ne 1 -or $full[0].ForegroundRecall -ne 0) {
    throw 'Full-page region recall must be 0 for the missing-title pair.'
}
if ($bottomMissing.Count -ne 1 -or $null -ne $bottomMissing[0].ForegroundRecall) {
    throw 'Empty bottom-half region recall must be null for the missing-title pair.'
}
Write-Host 'PASS: regions carry recall (full=0, empty=null).'

$identicalRegions = Read-ProbePage $titledPng $titledPng 'comparison-identical-regions' $regionSpecs
$identicalBottom = @(Get-Content -Raw -LiteralPath (Join-Path $identicalRegions.Directory 'region-metrics.json') | ConvertFrom-Json | Where-Object { $_.Region -eq 'bottom-half' })
if ($identicalBottom.Count -ne 1 -or $null -ne $identicalBottom[0].ForegroundRecall) {
    throw 'Empty bottom-half region recall must be null on identical pages.'
}
if ($identicalBottom[0].ForegroundColorHistogramCorrelation -ne 1) {
    throw 'Probe premise broken: expected histogram 1 for the empty bottom-half region.'
}
Write-Host 'PASS: empty bottom-half region -> ForegroundRecall=null while histogram reports 1.'

Write-Host 'VisualDiff foreground-recall probe passed (6 cases).'
