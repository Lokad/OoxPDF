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

function Invoke-DotnetBuildIfStale {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Project,

        [Parameter(Mandatory = $true)]
        [string] $OutputDll,

        [Parameter(Mandatory = $true)]
        [string] $Description,

        [string[]] $AdditionalSourceDirectories = @()
    )

    $projectDirectory = Split-Path -Parent $Project
    $sourceDirectories = @($projectDirectory) + $AdditionalSourceDirectories
    $sourceNewest = $sourceDirectories | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -Recurse -Include *.cs,*.csproj
    } |
        Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
    if ((Test-Path -LiteralPath $OutputDll) -and $sourceNewest.LastWriteTimeUtc -le (Get-Item -LiteralPath $OutputDll).LastWriteTimeUtc) {
        return
    }

    dotnet build $Project --tl:off --nologo -v minimal
    if ($LASTEXITCODE -ne 0) {
        throw "$Description build failed with exit code $LASTEXITCODE."
    }
}

$inputPath = Join-Path $caseDirectory $manifest.input
$inputFull = (Resolve-Path -LiteralPath $inputPath).Path
$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
$dpi = if ($manifest.dpi -ne $null) { [int]$manifest.dpi } else { 144 }

$cliProject = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj"
$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
$librarySourceDirectory = Join-Path $repoRoot "src/Lokad.OoxPdf"
Invoke-DotnetBuildIfStale -Project $cliProject -OutputDll $cliDll -Description "CLI" -AdditionalSourceDirectories @($librarySourceDirectory)
dotnet $cliDll convert $inputFull $candidatePdf --diagnostics $diagnostics
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "RenderCachedReference.ps1") -InputPath $inputFull -OutputDirectory $referenceDir -Dpi $dpi
& (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $candidatePdf -OutputDirectory $candidateDir -Dpi $dpi

$visualDiffProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj"
$visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/net10.0/Lokad.OoxPdf.VisualDiff.dll"
Invoke-DotnetBuildIfStale -Project $visualDiffProject -OutputDll $visualDiffDll -Description "VisualDiff"
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

if ($manifest.expected.minStructuralSimilarity -ne $null) {
    $minStructuralSimilarity = [double]$manifest.expected.minStructuralSimilarity
    $exceeded = @($metrics | Where-Object { $_.StructuralSimilarity -eq $null -or [double]$_.StructuralSimilarity -lt $minStructuralSimilarity })
    if ($exceeded.Count -ne 0) {
        $worst = $exceeded | Sort-Object -Property StructuralSimilarity | Select-Object -First 1
        throw "Structural similarity gate failed. Page $($worst.Page) was $($worst.StructuralSimilarity), minimum is $minStructuralSimilarity."
    }
}

if ($manifest.expected.minForegroundColorHistogramCorrelation -ne $null) {
    $minColorHistogram = [double]$manifest.expected.minForegroundColorHistogramCorrelation
    $exceeded = @($metrics | Where-Object { $_.ForegroundColorHistogramCorrelation -eq $null -or [double]$_.ForegroundColorHistogramCorrelation -lt $minColorHistogram })
    if ($exceeded.Count -ne 0) {
        $worst = $exceeded | Sort-Object -Property ForegroundColorHistogramCorrelation | Select-Object -First 1
        throw "Foreground color histogram gate failed. Page $($worst.Page) was $($worst.ForegroundColorHistogramCorrelation), minimum is $minColorHistogram."
    }
}

if ($manifest.expected.maxTextOperationPositionDelta -ne $null) {
    $textInspectRoot = Join-Path $comparisonDir "pdf-text"
    $referenceTextInspect = Join-Path $textInspectRoot "reference"
    $candidateTextInspect = Join-Path $textInspectRoot "candidate"
    New-Item -ItemType Directory -Force -Path $referenceTextInspect, $candidateTextInspect | Out-Null

    & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
        -InputPdf (Join-Path $referenceDir "reference.pdf") `
        -OutputDirectory $referenceTextInspect
    & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
        -InputPdf $candidatePdf `
        -OutputDirectory $candidateTextInspect

    $referenceTextOperations = Join-Path $referenceTextInspect "text-operations.json"
    $candidateTextOperations = Join-Path $candidateTextInspect "text-operations.json"
    if (-not (Test-Path -LiteralPath $referenceTextOperations) -or -not (Test-Path -LiteralPath $candidateTextOperations)) {
        throw "PDF text operation gate failed because one inspected PDF had no text operations."
    }

    $positionTolerance = [double]$manifest.expected.maxTextOperationPositionDelta
    $fontSizeTolerance = if ($manifest.expected.maxTextOperationFontSizeDelta -ne $null) {
        [double]$manifest.expected.maxTextOperationFontSizeDelta
    }
    else {
        0.01
    }
    $characterSpacingTolerance = if ($manifest.expected.maxTextOperationCharacterSpacingDelta -ne $null) {
        [double]$manifest.expected.maxTextOperationCharacterSpacingDelta
    }
    else {
        0.001
    }

    $compareTextArgs = @{
        Reference = $referenceTextOperations
        Candidate = $candidateTextOperations
        PositionTolerance = $positionTolerance
        FontSizeTolerance = $fontSizeTolerance
        CharacterSpacingTolerance = $characterSpacingTolerance
    }
    if ($manifest.expected.matchTextOperationsByPosition -eq $true) {
        $compareTextArgs.MatchByPosition = $true
    }
    if ($manifest.expected.compareTextOperationsWithEffectiveMatrix -eq $true) {
        $compareTextArgs.UseEffectiveMatrix = $true
    }
    if ($manifest.expected.compareDecodedTextOperations -eq $true) {
        $compareTextArgs.CompareDecodedText = $true
    }

    & (Join-Path $PSScriptRoot "ComparePdfTextOperations.ps1") @compareTextArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF text operation gate failed."
    }
}

if ($manifest.expected.maxTextLineStartDelta -ne $null) {
    $textInspectRoot = Join-Path $comparisonDir "pdf-text"
    $referenceTextInspect = Join-Path $textInspectRoot "reference"
    $candidateTextInspect = Join-Path $textInspectRoot "candidate"
    New-Item -ItemType Directory -Force -Path $referenceTextInspect, $candidateTextInspect | Out-Null

    $referenceTextOperations = Join-Path $referenceTextInspect "text-operations.json"
    $candidateTextOperations = Join-Path $candidateTextInspect "text-operations.json"
    if (-not (Test-Path -LiteralPath $referenceTextOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf (Join-Path $referenceDir "reference.pdf") `
            -OutputDirectory $referenceTextInspect
    }

    if (-not (Test-Path -LiteralPath $candidateTextOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf $candidatePdf `
            -OutputDirectory $candidateTextInspect
    }

    if (-not (Test-Path -LiteralPath $referenceTextOperations) -or -not (Test-Path -LiteralPath $candidateTextOperations)) {
        throw "PDF text line-start gate failed because one inspected PDF had no text operations."
    }

    $lineStartTolerance = [double]$manifest.expected.maxTextLineStartDelta
    $lineYTolerance = if ($manifest.expected.maxTextLineStartYDelta -ne $null) {
        [double]$manifest.expected.maxTextLineStartYDelta
    }
    else {
        0.75
    }

    if ($manifest.expected.matchTextLineStartsByPosition -eq $true) {
        & (Join-Path $PSScriptRoot "ComparePdfTextLineStarts.ps1") `
            -Reference $referenceTextOperations `
            -Candidate $candidateTextOperations `
            -StartTolerance $lineStartTolerance `
            -LineTolerance $lineYTolerance `
            -MatchByPosition
    }
    else {
        & (Join-Path $PSScriptRoot "ComparePdfTextLineStarts.ps1") `
            -Reference $referenceTextOperations `
            -Candidate $candidateTextOperations `
            -StartTolerance $lineStartTolerance `
            -LineTolerance $lineYTolerance
    }
    if ($LASTEXITCODE -ne 0) {
        throw "PDF text line-start gate failed."
    }
}

if ($manifest.expected.maxGraphicsOperationBoundsDelta -ne $null) {
    $graphicsInspectRoot = Join-Path $comparisonDir "pdf-graphics"
    $referenceGraphicsInspect = Join-Path $graphicsInspectRoot "reference"
    $candidateGraphicsInspect = Join-Path $graphicsInspectRoot "candidate"
    New-Item -ItemType Directory -Force -Path $referenceGraphicsInspect, $candidateGraphicsInspect | Out-Null

    $referenceGraphicsOperations = Join-Path $referenceGraphicsInspect "graphics-operations.json"
    $candidateGraphicsOperations = Join-Path $candidateGraphicsInspect "graphics-operations.json"
    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf (Join-Path $referenceDir "reference.pdf") `
            -OutputDirectory $referenceGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf $candidatePdf `
            -OutputDirectory $candidateGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations) -or -not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        throw "PDF graphics operation gate failed because one inspected PDF had no graphics operations."
    }

    $compareGraphicsArgs = @{
        Reference = $referenceGraphicsOperations
        Candidate = $candidateGraphicsOperations
        BoundsTolerance = [double]$manifest.expected.maxGraphicsOperationBoundsDelta
    }

    if ($manifest.expected.maxGraphicsOperationLineWidthDelta -ne $null) {
        $compareGraphicsArgs.LineWidthTolerance = [double]$manifest.expected.maxGraphicsOperationLineWidthDelta
    }
    if ($manifest.expected.compareGraphicsOperationKinds -ne $null) {
        $compareGraphicsArgs.Kinds = @($manifest.expected.compareGraphicsOperationKinds)
    }
    if ($manifest.expected.compareGraphicsOperationPageNumber -ne $null) {
        $compareGraphicsArgs.PageNumber = [int]$manifest.expected.compareGraphicsOperationPageNumber
    }
    if ($manifest.expected.matchGraphicsOperationsByBounds -eq $true) {
        $compareGraphicsArgs.MatchByBounds = $true
    }

    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareGraphicsArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF graphics operation gate failed."
    }
}

if ($manifest.expected.maxChartGraphicsStructureBoundsDelta -ne $null) {
    $graphicsInspectRoot = Join-Path $comparisonDir "pdf-graphics"
    $referenceGraphicsInspect = Join-Path $graphicsInspectRoot "reference"
    $candidateGraphicsInspect = Join-Path $graphicsInspectRoot "candidate"
    New-Item -ItemType Directory -Force -Path $referenceGraphicsInspect, $candidateGraphicsInspect | Out-Null

    $referenceGraphicsOperations = Join-Path $referenceGraphicsInspect "graphics-operations.json"
    $candidateGraphicsOperations = Join-Path $candidateGraphicsInspect "graphics-operations.json"
    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf (Join-Path $referenceDir "reference.pdf") `
            -OutputDirectory $referenceGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf $candidatePdf `
            -OutputDirectory $candidateGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations) -or -not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        throw "PDF chart graphics structure gate failed because one inspected PDF had no graphics operations."
    }

    $referenceChartStructures = Join-Path $referenceGraphicsInspect "chart-structures.json"
    $candidateChartStructures = Join-Path $candidateGraphicsInspect "chart-structures.json"
    $classifyReferenceArgs = @{
        InputPath = $referenceGraphicsOperations
        Output = $referenceChartStructures
    }
    $classifyCandidateArgs = @{
        InputPath = $candidateGraphicsOperations
        Output = $candidateChartStructures
    }
    if ($manifest.expected.compareChartGraphicsStructurePageNumber -ne $null) {
        $classifyReferenceArgs.PageNumber = [int]$manifest.expected.compareChartGraphicsStructurePageNumber
        $classifyCandidateArgs.PageNumber = [int]$manifest.expected.compareChartGraphicsStructurePageNumber
    }
    if ($manifest.expected.chartGraphicsLineTolerance -ne $null) {
        $classifyReferenceArgs.LineTolerance = [double]$manifest.expected.chartGraphicsLineTolerance
        $classifyCandidateArgs.LineTolerance = [double]$manifest.expected.chartGraphicsLineTolerance
    }
    if ($manifest.expected.chartGraphicsMinLineLength -ne $null) {
        $classifyReferenceArgs.MinLineLength = [double]$manifest.expected.chartGraphicsMinLineLength
        $classifyCandidateArgs.MinLineLength = [double]$manifest.expected.chartGraphicsMinLineLength
    }
    if ($manifest.expected.chartGraphicsMarkerMaxSize -ne $null) {
        $classifyReferenceArgs.MarkerMaxSize = [double]$manifest.expected.chartGraphicsMarkerMaxSize
        $classifyCandidateArgs.MarkerMaxSize = [double]$manifest.expected.chartGraphicsMarkerMaxSize
    }

    & (Join-Path $PSScriptRoot "ClassifyPdfChartGraphics.ps1") @classifyReferenceArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart graphics structure gate failed while classifying the reference PDF."
    }
    & (Join-Path $PSScriptRoot "ClassifyPdfChartGraphics.ps1") @classifyCandidateArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart graphics structure gate failed while classifying the candidate PDF."
    }

    $compareChartArgs = @{
        Reference = $referenceChartStructures
        Candidate = $candidateChartStructures
        BoundsTolerance = [double]$manifest.expected.maxChartGraphicsStructureBoundsDelta
        MatchByBounds = $true
    }
    if ($manifest.expected.maxChartGraphicsStructureLineWidthDelta -ne $null) {
        $compareChartArgs.LineWidthTolerance = [double]$manifest.expected.maxChartGraphicsStructureLineWidthDelta
    }
    if ($manifest.expected.compareChartGraphicsStructureKinds -ne $null) {
        $compareChartArgs.Kinds = @($manifest.expected.compareChartGraphicsStructureKinds)
    }
    else {
        $compareChartArgs.Kinds = @("HorizontalLine", "VerticalLine", "PlotBoxCandidate", "FilledRegion", "MarkerCandidate", "ClipBox")
    }

    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareChartArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart graphics structure gate failed."
    }
}

if ($manifest.expected.maxChartTextStructurePositionDelta -ne $null) {
    $textInspectRoot = Join-Path $comparisonDir "pdf-text"
    $referenceTextInspect = Join-Path $textInspectRoot "reference"
    $candidateTextInspect = Join-Path $textInspectRoot "candidate"
    $graphicsInspectRoot = Join-Path $comparisonDir "pdf-graphics"
    $referenceGraphicsInspect = Join-Path $graphicsInspectRoot "reference"
    $candidateGraphicsInspect = Join-Path $graphicsInspectRoot "candidate"
    New-Item -ItemType Directory -Force -Path $referenceTextInspect, $candidateTextInspect, $referenceGraphicsInspect, $candidateGraphicsInspect | Out-Null

    $referenceTextOperations = Join-Path $referenceTextInspect "text-operations.json"
    $candidateTextOperations = Join-Path $candidateTextInspect "text-operations.json"
    if (-not (Test-Path -LiteralPath $referenceTextOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf (Join-Path $referenceDir "reference.pdf") `
            -OutputDirectory $referenceTextInspect
    }

    if (-not (Test-Path -LiteralPath $candidateTextOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf $candidatePdf `
            -OutputDirectory $candidateTextInspect
    }

    if (-not (Test-Path -LiteralPath $referenceTextOperations) -or -not (Test-Path -LiteralPath $candidateTextOperations)) {
        throw "PDF chart text structure gate failed because one inspected PDF had no text operations."
    }

    $referenceGraphicsOperations = Join-Path $referenceGraphicsInspect "graphics-operations.json"
    $candidateGraphicsOperations = Join-Path $candidateGraphicsInspect "graphics-operations.json"
    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf (Join-Path $referenceDir "reference.pdf") `
            -OutputDirectory $referenceGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        & (Join-Path $PSScriptRoot "InspectPdf.ps1") `
            -InputPdf $candidatePdf `
            -OutputDirectory $candidateGraphicsInspect
    }

    if (-not (Test-Path -LiteralPath $referenceGraphicsOperations) -or -not (Test-Path -LiteralPath $candidateGraphicsOperations)) {
        throw "PDF chart text structure gate failed because one inspected PDF had no graphics operations for plot-box classification."
    }

    $referenceChartStructures = Join-Path $referenceGraphicsInspect "chart-structures.json"
    $candidateChartStructures = Join-Path $candidateGraphicsInspect "chart-structures.json"
    $classifyReferenceGraphicsArgs = @{
        InputPath = $referenceGraphicsOperations
        Output = $referenceChartStructures
    }
    $classifyCandidateGraphicsArgs = @{
        InputPath = $candidateGraphicsOperations
        Output = $candidateChartStructures
    }
    if ($manifest.expected.compareChartTextStructurePageNumber -ne $null) {
        $classifyReferenceGraphicsArgs.PageNumber = [int]$manifest.expected.compareChartTextStructurePageNumber
        $classifyCandidateGraphicsArgs.PageNumber = [int]$manifest.expected.compareChartTextStructurePageNumber
    }

    & (Join-Path $PSScriptRoot "ClassifyPdfChartGraphics.ps1") @classifyReferenceGraphicsArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart text structure gate failed while classifying reference graphics."
    }
    & (Join-Path $PSScriptRoot "ClassifyPdfChartGraphics.ps1") @classifyCandidateGraphicsArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart text structure gate failed while classifying candidate graphics."
    }

    $referenceChartTextStructures = Join-Path $referenceTextInspect "chart-text-structures.json"
    $candidateChartTextStructures = Join-Path $candidateTextInspect "chart-text-structures.json"
    $classifyReferenceTextArgs = @{
        InputPath = $referenceTextOperations
        ChartStructures = $referenceChartStructures
        Output = $referenceChartTextStructures
    }
    $classifyCandidateTextArgs = @{
        InputPath = $candidateTextOperations
        ChartStructures = $candidateChartStructures
        Output = $candidateChartTextStructures
    }
    if ($manifest.expected.compareChartTextStructurePageNumber -ne $null) {
        $classifyReferenceTextArgs.PageNumber = [int]$manifest.expected.compareChartTextStructurePageNumber
        $classifyCandidateTextArgs.PageNumber = [int]$manifest.expected.compareChartTextStructurePageNumber
    }
    if ($manifest.expected.chartTextPlotTolerance -ne $null) {
        $classifyReferenceTextArgs.PlotTolerance = [double]$manifest.expected.chartTextPlotTolerance
        $classifyCandidateTextArgs.PlotTolerance = [double]$manifest.expected.chartTextPlotTolerance
    }

    & (Join-Path $PSScriptRoot "ClassifyPdfChartText.ps1") @classifyReferenceTextArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart text structure gate failed while classifying reference text."
    }
    & (Join-Path $PSScriptRoot "ClassifyPdfChartText.ps1") @classifyCandidateTextArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart text structure gate failed while classifying candidate text."
    }

    $compareChartTextArgs = @{
        Reference = $referenceChartTextStructures
        Candidate = $candidateChartTextStructures
        BoundsTolerance = [double]$manifest.expected.maxChartTextStructurePositionDelta
        LineWidthTolerance = 0
        MatchByBounds = $true
    }
    if ($manifest.expected.compareChartTextStructureKinds -ne $null) {
        $compareChartTextArgs.Kinds = @($manifest.expected.compareChartTextStructureKinds)
    }
    else {
        $compareChartTextArgs.Kinds = @("AbovePlotText", "BelowPlotText", "InsidePlotText", "LeftAxisText", "RightSideText", "OuterChartText", "ChartText")
    }

    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareChartTextArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart text structure gate failed."
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
