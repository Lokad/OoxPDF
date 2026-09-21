# Allocation measurement driver (G01 slice 1): builds AllocProbe, runs it over
# tracked public fixtures, and prints a whole-pipeline summary table.
# Outputs go under ignored artifacts/alloc/. No Office/COM is required.

param(
    [string[]] $Corpus = @(),

    [int] $Warmup = 1,

    [int] $Iterations = 3,

    [string] $Out = "artifacts/alloc/report.json",

    [switch] $SkipBuild,

    [switch] $Stages,

    [switch] $SelfTest,

    [ValidateSet("buffer", "file")]
    [string] $OutputMode = "buffer",

    [switch] $Isolate
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$defaultCorpus = @(
    "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-08-composite-port-a.pptx",
    "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-11-composite-chart-port.pptx",
    "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-10-table-font-fragmentation.pptx",
    "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-04-typography-font-families.pptx",
    "tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-07-jpeg-image.pptx",
    "tests/Lokad.OoxPdf.Tests/Cases/docx-basic-paragraphs.docx",
    "tests/Lokad.OoxPdf.Tests/Cases/docx-tables.docx",
    "tests/Lokad.OoxPdf.Tests/Cases/docx-images.docx",
    "tests/Lokad.OoxPdf.Tests/Cases/docx-markup-note-links-fields.docx"
)
if ($Corpus.Count -eq 0) {
    $Corpus = $defaultCorpus
}

$probeProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.AllocProbe/Lokad.OoxPdf.AllocProbe.csproj"
$probeDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.AllocProbe/bin/Release/net10.0/Lokad.OoxPdf.AllocProbe.dll"
$allocRoot = Join-Path $repoRoot "artifacts/alloc"
New-Item -ItemType Directory -Force -Path $allocRoot | Out-Null

if ($SkipBuild) {
    if (-not (Test-Path -LiteralPath $probeDll)) {
        throw "AllocProbe DLL is missing and -SkipBuild was requested: $probeDll. Build once without -SkipBuild first."
    }
}
else {
    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Project $probeProject -OutputDll $probeDll -Description "AllocProbe" -Configuration Release -RecordPath (Join-Path $allocRoot "build-info.json")
}

if ($SelfTest) {
    & dotnet $probeDll --self-test
    if ($LASTEXITCODE -ne 0) {
        throw "AllocProbe self-test failed with exit code $LASTEXITCODE."
    }

    return
}

$inputs = @()
foreach ($entry in $Corpus) {
    $full = if ([System.IO.Path]::IsPathRooted($entry)) { $entry } else { Join-Path $repoRoot $entry }
    if (-not (Test-Path -LiteralPath $full)) {
        throw "Corpus input is missing: $entry. Pass explicit -Corpus paths that exist."
    }

    $inputs += (Resolve-Path -LiteralPath $full).Path
}

$outFull = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $repoRoot $Out }
$probeArgs = @("--out", $outFull, "--warmup", $Warmup, "--iterations", $Iterations, "--output-mode", $OutputMode)
if ($Stages) {
    $probeArgs += "--stages"
}
if ($Isolate) {
    $probeArgs += "--isolate"
}
$probeArgs += $inputs
& dotnet $probeDll $probeArgs
if ($LASTEXITCODE -ne 0) {
    throw "AllocProbe failed with exit code $LASTEXITCODE."
}

$report = Get-Content -Raw -LiteralPath $outFull | ConvertFrom-Json
Write-Host ("{0,-55} {1,9} {2,9} {3,5} {4,12} {5,12} {6,9} {7,6}" -f "input", "in KB", "out KB", "pages", "cold MB", "warm MB", "warm ms", "stable")
foreach ($entry in @($report.inputs)) {
    $warmAllocs = @($entry.warm | ForEach-Object { [double]$_.allocatedBytes }) | Sort-Object
    $warmMs = @($entry.warm | ForEach-Object { [double]$_.elapsedMilliseconds }) | Sort-Object
    $warmAllocMedian = $warmAllocs[[math]::Floor($warmAllocs.Count / 2)]
    $warmMsMedian = $warmMs[[math]::Floor($warmMs.Count / 2)]
    Write-Host ("{0,-55} {1,9} {2,9} {3,5} {4,12} {5,12} {6,9} {7,6}" -f $entry.name, ([math]::Round($entry.inputBytes / 1KB, 1)), ([math]::Round($entry.outputBytes / 1KB, 1)), $entry.pageCount, ([math]::Round($entry.cold.allocatedBytes / 1MB, 2)), ([math]::Round($warmAllocMedian / 1MB, 2)), ([math]::Round($warmMsMedian, 1)), $entry.outputStable)
    if ($null -ne $entry.stages) {
        Write-Host ("  stages [{0}] match={1}" -f $entry.stages.stageSet, $entry.stages.stagedMatchesWhole)
        foreach ($stageName in @("open", "read", "scene", "render", "write")) {
            $stage = $entry.stages.$stageName
            if ($null -ne $stage) {
                Write-Host ("    {0,-8} {1,12} {2,9}" -f $stageName, ([math]::Round($stage.allocatedBytes / 1MB, 2)), ([math]::Round($stage.elapsedMilliseconds, 1)))
            }
        }
    }
}

Write-Host ""
Write-Host ("Report: {0} (build {1}, {2}, serverGc={3})." -f $outFull, $report.buildConfiguration, $report.framework, $report.serverGc)
Write-Host ("Scope: {0}" -f $report.allocationScope)
