param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [switch] $CacheOnlyReference,

    [string] $RunId,

    [switch] $SkipBuild,

    [string] $CliPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pdfInspectorProject = Join-Path $repoRoot 'tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj'
$pdfInspectorDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/$Framework/Lokad.OoxPdf.PdfInspect.dll"
$caseFull = (Resolve-Path -LiteralPath $Case).Path
$caseDirectory = Split-Path -Parent $caseFull
$caseSha256 = (Get-FileHash -LiteralPath $caseFull).Hash
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
$runId = if (-not [string]::IsNullOrWhiteSpace($RunId)) { ($RunId -replace '[^A-Za-z0-9_.-]', '-') } else { "{0}-p{1}-{2}" -f (Get-Date -Format "yyyyMMdd-HHmmss-fff"), $PID, ([System.Guid]::NewGuid().ToString("N").Substring(0, 8)) }
if ($runId -in '.', '..') { throw 'RunId must name a run directory.' }
$runRoot = Join-Path $repoRoot ("artifacts/visual/{0}/{1}" -f $manifest.id, $runId)
if (Test-Path -LiteralPath $runRoot) { throw "Visual run directory already exists. Choose a new RunId. $runRoot" }
[IO.Directory]::CreateDirectory($runRoot) | Out-Null
# CreateNew also refuses a competing invocation that passed the directory check.
$runMarker = Join-Path $runRoot '.started'
try {
    [IO.File]::Open($runMarker, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None).Dispose()
}
catch [IO.IOException] {
    if (Test-Path -LiteralPath $runMarker) { throw "Visual run directory already exists. Choose a new RunId. $runRoot" }
    throw
}
$referenceDir = Join-Path $runRoot "reference"
$candidateDir = Join-Path $runRoot "candidate"
$comparisonDir = Join-Path $runRoot "comparison"

New-Item -ItemType Directory -Force -Path $referenceDir, $candidateDir, $comparisonDir | Out-Null

function Invoke-RepoBuild([string] $Project, [string] $OutputDll, [string] $Description, [string] $RecordPath) {
    if ($SkipBuild) {
        if (-not (Test-Path -LiteralPath $OutputDll)) {
            throw "$Description DLL is missing and -SkipBuild was requested: $OutputDll. Build once without -SkipBuild first."
        }

        return
    }

    $params = @{ Project = $Project; OutputDll = $OutputDll; Description = $Description }
    if (-not [string]::IsNullOrWhiteSpace($RecordPath)) {
        $params.RecordPath = $RecordPath
    }

    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Framework $Framework @params
}
$pdfInspections = @{}
function Get-ComparisonPdfInspection([ValidateSet('reference','candidate')][string] $Side) {
    if (!$pdfInspections.ContainsKey($Side)) {
        if ($pdfInspections.Count -eq 0) {
            Invoke-RepoBuild $pdfInspectorProject $pdfInspectorDll 'PDF inspect' (Join-Path $buildInfoDir 'pdf-inspect.json')
        }
        $inputPdf = if ($Side -eq 'reference') { Join-Path $referenceDir 'reference.pdf' } else { $candidatePdf }
        $outputDirectory = Join-Path $comparisonDir "pdf-text/$Side"
        & (Join-Path $PSScriptRoot 'InspectPdf.ps1') -Framework $Framework -InputPdf $inputPdf `
            -OutputDirectory $outputDirectory -InspectorPath $pdfInspectorDll -SkipBuild | Out-Host
        foreach ($reportName in 'text-operations.json', 'graphics-operations.json', 'font-resources.json') {
            if (!(Test-Path -LiteralPath (Join-Path $outputDirectory $reportName) -PathType Leaf)) {
                throw "PDF inspection did not produce $reportName for the $Side PDF."
            }
        }
        $pdfInspections[$Side] = $outputDirectory
    }
    return $pdfInspections[$Side]
}
function Get-ComparisonChartGraphics([string] $Side, [hashtable] $Options, [string] $ReportName) {
    $inspection = Get-ComparisonPdfInspection $Side
    $output = Join-Path $comparisonDir "pdf-graphics/$Side/$ReportName"
    & (Join-Path $PSScriptRoot 'ClassifyPdfChartGraphics.ps1') -InputPath (Join-Path $inspection 'graphics-operations.json') `
        -TextOperations (Join-Path $inspection 'text-operations.json') -Output $output @Options | Out-Host
    if ($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath $output -PathType Leaf)) {
        throw "PDF chart graphics classification failed for the $Side PDF."
    }
    return $output
}
. (Join-Path $PSScriptRoot "JsonArray.ps1")

$inputPath = Join-Path $caseDirectory $manifest.input
$inputFull = (Resolve-Path -LiteralPath $inputPath).Path
$inputSha256 = (Get-FileHash -LiteralPath $inputFull).Hash
$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
$dpi = if ($manifest.dpi -ne $null) { [int]$manifest.dpi } else { 144 }
$docxMarkup = if ($manifest.PSObject.Properties.Name -contains "docxMarkup" -and $manifest.docxMarkup -ne $null) { [string]$manifest.docxMarkup } else { $null }
$docxMarkupGeometry = if ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and $manifest.docxMarkupGeometry -ne $null) { [string]$manifest.docxMarkupGeometry } else { $null }

$cliProject = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj"
$cliDll = if ($CliPath) { (Resolve-Path -LiteralPath $CliPath).Path } else {
    Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/$Framework/Lokad.OoxPdf.Cli.dll"
}
$buildInfoDir = Join-Path $runRoot "build-info"
New-Item -ItemType Directory -Force -Path $buildInfoDir | Out-Null
if (!$CliPath) { Invoke-RepoBuild $cliProject $cliDll "CLI" (Join-Path $buildInfoDir "cli.json") }
else {
    $explicitLibrary = Join-Path (Split-Path -Parent $cliDll) 'Lokad.OoxPdf.dll'
    [ordered]@{CliDll=$cliDll;CliSha256=(Get-FileHash -LiteralPath $cliDll).Hash;
        LibraryDll=$explicitLibrary;LibrarySha256=(Get-FileHash -LiteralPath $explicitLibrary).Hash;
        InputSha256=$inputSha256;CaseSha256=$caseSha256;
        BuildRequested=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $buildInfoDir 'explicit-cli.json') -Encoding utf8
}
$candidateArgs = @("convert", $inputFull, $candidatePdf, "--diagnostics", $diagnostics)
if (-not [string]::IsNullOrWhiteSpace($docxMarkup)) {
    if ($manifest.kind -ne "docx") {
        throw "Case '$($manifest.id)' sets docxMarkup but is not a DOCX case."
    }

    $candidateArgs += @("--docx-markup", $docxMarkup)
}
if (-not [string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
    if ($manifest.kind -ne "docx") {
        throw "Case '$($manifest.id)' sets docxMarkupGeometry but is not a DOCX case."
    }

    $candidateArgs += @("--docx-markup-geometry", $docxMarkupGeometry)
}

Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Converting candidate PDF...";
dotnet $cliDll @candidateArgs
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

$referenceArgs = @{
    InputPath = $inputFull
    OutputDirectory = $referenceDir
    Dpi = $dpi
    CaseId = $manifest.id
    SkipBuild = $SkipBuild
}
if ($CacheOnlyReference) {
    $referenceArgs.CacheOnly = $true
    if (-not [string]::IsNullOrWhiteSpace($docxMarkup) -or -not [string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
        . (Join-Path $PSScriptRoot "DocxMarkupReference.ps1")
        $referenceArgs.CacheVariant = (Resolve-DocxMarkupReferenceSettings $manifest).Variant
    }
}
Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Rendering reference...";
& (Join-Path $PSScriptRoot "RenderCachedReference.ps1") -Framework $Framework @referenceArgs
Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Rasterizing candidate...";
& (Join-Path $PSScriptRoot "RasterizePdf.ps1") -Framework $Framework -InputPdf $candidatePdf -OutputDirectory $candidateDir -Dpi $dpi -SkipBuild:$SkipBuild

$visualDiffProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj"
$visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/$Framework/Lokad.OoxPdf.VisualDiff.dll"
Invoke-RepoBuild $visualDiffProject $visualDiffDll "VisualDiff" (Join-Path $buildInfoDir "visual-diff.json")
Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Comparing pages...";
dotnet $visualDiffDll $referenceDir $candidateDir $comparisonDir
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff failed with exit code $LASTEXITCODE."
}

$metricsPath = Join-Path $comparisonDir "metrics.json"
$metrics = Get-Content -Raw -LiteralPath $metricsPath | ConvertFrom-Json
$referencePages = @(Get-ChildItem -LiteralPath $referenceDir -Filter "page-*.png")
$candidatePages = @(Get-ChildItem -LiteralPath $candidateDir -Filter "page-*.png")

# Bind the measurements even when a later visual or content gate fails.
if ((Get-FileHash -LiteralPath $inputFull).Hash -ne $inputSha256 -or
    (Get-FileHash -LiteralPath $caseFull).Hash -ne $caseSha256) {
    throw "Source or case manifest changed during the visual run."
}
$runInfo = [ordered]@{
    Schema = "ooxpdf-visual-run-info/1"
    RunId = $runId
    CaseId = $manifest.id
    Kind = $manifest.kind
    Input = $inputFull.Substring($repoRoot.Length).TrimStart("\", "/").Replace("\", "/")
    InputSha256 = $inputSha256
    CaseSha256 = $caseSha256
    ReferencePdfSha256 = (Get-FileHash -LiteralPath (Join-Path $referenceDir "reference.pdf")).Hash
    PdfiumPinSha256 = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "vendor/pdfium/win-x64/pdfium.sha256")).Hash
    Dpi = $dpi
    DocxMarkup = $docxMarkup
    DocxMarkupGeometry = $docxMarkupGeometry
    CliDll = $cliDll
    CliSha256 = if (Test-Path -LiteralPath $cliDll) { (Get-FileHash -LiteralPath $cliDll -Algorithm SHA256).Hash.ToLowerInvariant() } else { $null }
    LibraryDll = $null
    VisualDiffDll = $visualDiffDll
    VisualDiffSha256 = if (Test-Path -LiteralPath $visualDiffDll) { (Get-FileHash -LiteralPath $visualDiffDll -Algorithm SHA256).Hash.ToLowerInvariant() } else { $null }
    RasterizerDll = $null
    GitCommit = $null
    GitDirty = $null
    GeneratedAtUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
}
$libraryDll = Join-Path (Split-Path -Parent $cliDll) "Lokad.OoxPdf.dll"
if (Test-Path -LiteralPath $libraryDll) {
    $runInfo.LibraryDll = $libraryDll
    $runInfo.LibrarySha256 = (Get-FileHash -LiteralPath $libraryDll -Algorithm SHA256).Hash.ToLowerInvariant()
}

$rasterizerDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework/Lokad.OoxPdf.PdfiumRasterizer.dll"
if (Test-Path -LiteralPath $rasterizerDll) {
    $runInfo.RasterizerDll = $rasterizerDll
    $runInfo.RasterizerSha256 = (Get-FileHash -LiteralPath $rasterizerDll -Algorithm SHA256).Hash.ToLowerInvariant()
}

if (Test-Path -LiteralPath (Join-Path $repoRoot '.git')) {
    try {
        $runInfo.GitCommit = (& git -C $repoRoot rev-parse HEAD 2>$null | Select-Object -First 1).Trim()
        $runInfo.GitDirty = [bool](& git --no-optional-locks -C $repoRoot status --short 2>$null | Select-Object -First 1)
    }
    catch {
    }
}

$runInfo | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runRoot "run-info.json") -Encoding UTF8

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

if ($manifest.expected.minForegroundRecall -ne $null) {
    $minForegroundRecall = [double]$manifest.expected.minForegroundRecall
    $exceeded = @($metrics | Where-Object { $_.ForegroundRecall -eq $null -or [double]$_.ForegroundRecall -lt $minForegroundRecall })
    if ($exceeded.Count -ne 0) {
        $worst = $exceeded | Sort-Object -Property ForegroundRecall | Select-Object -First 1
        throw "Foreground recall gate failed. Page $($worst.Page) was $($worst.ForegroundRecall), minimum is $minForegroundRecall."
    }
}

if ($manifest.expected.maxTextOperationPositionDelta -ne $null) {
    $referenceTextOperations = Join-Path (Get-ComparisonPdfInspection reference) 'text-operations.json'
    $candidateTextOperations = Join-Path (Get-ComparisonPdfInspection candidate) 'text-operations.json'

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
    $referenceTextOperations = Join-Path (Get-ComparisonPdfInspection reference) 'text-operations.json'
    $candidateTextOperations = Join-Path (Get-ComparisonPdfInspection candidate) 'text-operations.json'

    $lineStartTolerance = [double]$manifest.expected.maxTextLineStartDelta
    $lineYTolerance = if ($manifest.expected.maxTextLineStartYDelta -ne $null) {
        [double]$manifest.expected.maxTextLineStartYDelta
    }
    else {
        0.75
    }

    $lineStartArgs = @{
        Reference = $referenceTextOperations
        Candidate = $candidateTextOperations
        StartTolerance = $lineStartTolerance
        LineTolerance = $lineYTolerance
    }
    if ($manifest.expected.maxTextLineStartYDelta -ne $null) {
        $lineStartArgs.BaselineTolerance = $lineYTolerance
    }
    if ($manifest.expected.matchTextLineStartsByPosition -eq $true) {
        $lineStartArgs.MatchByPosition = $true
    }
    if ($manifest.expected.compareFirstTextLineStartOnly -eq $true) {
        $lineStartArgs.FirstStartOnly = $true
    }
    if ($manifest.expected.ignoreWhitespaceOnlyTextLineStartOperations -eq $true) {
        $lineStartArgs.IgnoreWhitespaceOnlyOperations = $true
    }
    & (Join-Path $PSScriptRoot "ComparePdfTextLineStarts.ps1") @lineStartArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF text line-start gate failed."
    }
}

if ($manifest.expected.maxGraphicsOperationBoundsDelta -ne $null) {
    $referenceGraphicsOperations = Join-Path (Get-ComparisonPdfInspection reference) 'graphics-operations.json'
    $candidateGraphicsOperations = Join-Path (Get-ComparisonPdfInspection candidate) 'graphics-operations.json'

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
    if ($manifest.expected.compareGraphicsOperationOperators -ne $null) {
        $compareGraphicsArgs.Operators = @($manifest.expected.compareGraphicsOperationOperators)
    }
    if ($manifest.expected.compareGraphicsOperationPageNumber -ne $null) {
        $compareGraphicsArgs.PageNumber = [int]$manifest.expected.compareGraphicsOperationPageNumber
    }
    if ($manifest.expected.matchGraphicsOperationsByBounds -eq $true) {
        $compareGraphicsArgs.MatchByBounds = $true
    }
    if ($manifest.expected.matchGraphicsOperationOperators -eq $true) {
        $compareGraphicsArgs.MatchOperator = $true
    }
    if ($manifest.expected.compareGraphicsOperationSegmentCounts -eq $true) {
        $compareGraphicsArgs.MatchSegmentCount = $true
    }
    if ($manifest.expected.compareGraphicsOperationPathCommandCounts -eq $true) {
        $compareGraphicsArgs.MatchPathCommandCounts = $true
    }
    if ($manifest.expected.compareGraphicsOperationPathOperators -eq $true) {
        $compareGraphicsArgs.MatchPathOperators = $true
    }
    if ($manifest.expected.compareGraphicsOperationPathCommandCoordinates -eq $true -or $manifest.expected.maxGraphicsOperationPathCommandCoordinateDelta -ne $null) {
        $compareGraphicsArgs.MatchPathCommandCoordinates = $true
    }
    if ($manifest.expected.maxGraphicsOperationPathCommandCoordinateDelta -ne $null) {
        $compareGraphicsArgs.PathCommandCoordinateTolerance = [double]$manifest.expected.maxGraphicsOperationPathCommandCoordinateDelta
    }

    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareGraphicsArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF graphics operation gate failed."
    }
}

$hasChartGraphicsStructureGlobalBoundsTolerance = $manifest.expected.maxChartGraphicsStructureBoundsDelta -ne $null
$hasChartGraphicsStructureKindBoundsTolerances = $manifest.expected.maxChartGraphicsStructureBoundsDeltaByKind -ne $null
$chartGraphicsOptions = $null
if ($hasChartGraphicsStructureGlobalBoundsTolerance -or $hasChartGraphicsStructureKindBoundsTolerances) {
    $chartGraphicsOptions = @{}
    if ($null -ne $manifest.expected.compareChartGraphicsStructurePageNumber) {
        $chartGraphicsOptions.PageNumber = [int]$manifest.expected.compareChartGraphicsStructurePageNumber
    }
    if ($null -ne $manifest.expected.chartGraphicsLineTolerance) {
        $chartGraphicsOptions.LineTolerance = [double]$manifest.expected.chartGraphicsLineTolerance
    }
    if ($null -ne $manifest.expected.chartGraphicsMinLineLength) {
        $chartGraphicsOptions.MinLineLength = [double]$manifest.expected.chartGraphicsMinLineLength
    }
    if ($null -ne $manifest.expected.chartGraphicsMarkerMaxSize) {
        $chartGraphicsOptions.MarkerMaxSize = [double]$manifest.expected.chartGraphicsMarkerMaxSize
    }
    if ($null -ne $manifest.expected.chartGraphicsGridlineBoundsTolerance) {
        $chartGraphicsOptions.GridlineBoundsTolerance = [double]$manifest.expected.chartGraphicsGridlineBoundsTolerance
    }
    $referenceChartStructures = Get-ComparisonChartGraphics reference $chartGraphicsOptions 'chart-structures.json'
    $candidateChartStructures = Get-ComparisonChartGraphics candidate $chartGraphicsOptions 'chart-structures.json'

    $compareChartArgs = @{
        Reference = $referenceChartStructures
        Candidate = $candidateChartStructures
        MatchByBounds = $true
    }
    if ($hasChartGraphicsStructureGlobalBoundsTolerance) {
        $compareChartArgs.BoundsTolerance = [double]$manifest.expected.maxChartGraphicsStructureBoundsDelta
        $compareChartArgs.UseBoundsToleranceForUnlistedKinds = $true
    }
    if ($hasChartGraphicsStructureKindBoundsTolerances) {
        $boundsToleranceByKind = @{}
        foreach ($entry in $manifest.expected.maxChartGraphicsStructureBoundsDeltaByKind.PSObject.Properties) {
            $boundsToleranceByKind[[string]$entry.Name] = [double]$entry.Value
        }

        $compareChartArgs.BoundsToleranceByKind = $boundsToleranceByKind
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
    if ($manifest.expected.compareChartGraphicsStructureOperators -eq $true) {
        $compareChartArgs.MatchOperator = $true
    }
    if ($manifest.expected.compareChartGraphicsStructureSegmentCounts -eq $true) {
        $compareChartArgs.MatchSegmentCount = $true
    }
    if ($manifest.expected.compareChartGraphicsStructurePathCommandCounts -eq $true) {
        $compareChartArgs.MatchPathCommandCounts = $true
    }
    if ($manifest.expected.compareChartGraphicsStructurePathOperators -eq $true) {
        $compareChartArgs.MatchPathOperators = $true
    }
    if ($manifest.expected.compareChartGraphicsStructurePathCommandCoordinates -eq $true -or $manifest.expected.maxChartGraphicsStructurePathCommandCoordinateDelta -ne $null) {
        $compareChartArgs.MatchPathCommandCoordinates = $true
    }
    if ($manifest.expected.maxChartGraphicsStructurePathCommandCoordinateDelta -ne $null) {
        $compareChartArgs.PathCommandCoordinateTolerance = [double]$manifest.expected.maxChartGraphicsStructurePathCommandCoordinateDelta
    }
    if ($manifest.expected.compareChartGraphicsStructureStrokeColors -eq $true) {
        $compareChartArgs.MatchStrokeColor = $true
    }
    if ($manifest.expected.compareChartGraphicsStructureFillColors -eq $true) {
        $compareChartArgs.MatchFillColor = $true
    }
    if ($manifest.expected.compareChartGraphicsStructureLineCaps -eq $true) {
        $compareChartArgs.MatchLineCap = $true
    }
    if ($manifest.expected.compareChartGraphicsStructureLineJoins -eq $true) {
        $compareChartArgs.MatchLineJoin = $true
    }
    $hasChartGraphicsPathGeometryKindTolerances = $manifest.expected.maxChartGraphicsStructurePathGeometryDeltaByKind -ne $null
    if ($manifest.expected.compareChartGraphicsStructurePathGeometry -eq $true -or $hasChartGraphicsPathGeometryKindTolerances) {
        $compareChartArgs.MatchPathGeometry = $true
    }
    if ($manifest.expected.maxChartGraphicsStructurePathGeometryDelta -ne $null) {
        $compareChartArgs.PathGeometryTolerance = [double]$manifest.expected.maxChartGraphicsStructurePathGeometryDelta
        $compareChartArgs.UsePathGeometryToleranceForUnlistedKinds = $true
    }
    if ($hasChartGraphicsPathGeometryKindTolerances) {
        $pathGeometryToleranceByKind = @{}
        foreach ($entry in $manifest.expected.maxChartGraphicsStructurePathGeometryDeltaByKind.PSObject.Properties) {
            $pathGeometryToleranceByKind[[string]$entry.Name] = [double]$entry.Value
        }

        $compareChartArgs.PathGeometryToleranceByKind = $pathGeometryToleranceByKind
    }

    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareChartArgs
    if ($LASTEXITCODE -ne 0) {
        throw "PDF chart graphics structure gate failed."
    }

    foreach ($additionalComparison in @($manifest.expected.additionalChartGraphicsStructureComparisons)) {
        if ($null -eq $additionalComparison) {
            continue
        }

        $additionalCompareArgs = @{
            Reference = $referenceChartStructures
            Candidate = $candidateChartStructures
            MatchByBounds = $true
        }
        if ($additionalComparison.pageNumber -ne $null) {
            $additionalCompareArgs.PageNumber = [int]$additionalComparison.pageNumber
        }
        if ($additionalComparison.kinds -ne $null) {
            $additionalCompareArgs.Kinds = @($additionalComparison.kinds)
        }
        if ($additionalComparison.maxBoundsDelta -ne $null) {
            $additionalCompareArgs.BoundsTolerance = [double]$additionalComparison.maxBoundsDelta
            $additionalCompareArgs.UseBoundsToleranceForUnlistedKinds = $true
        }
        if ($additionalComparison.maxBoundsDeltaByKind -ne $null) {
            $boundsToleranceByKind = @{}
            foreach ($entry in $additionalComparison.maxBoundsDeltaByKind.PSObject.Properties) {
                $boundsToleranceByKind[[string]$entry.Name] = [double]$entry.Value
            }

            $additionalCompareArgs.BoundsToleranceByKind = $boundsToleranceByKind
        }
        if ($additionalComparison.maxLineWidthDelta -ne $null) {
            $additionalCompareArgs.LineWidthTolerance = [double]$additionalComparison.maxLineWidthDelta
        }
        if ($additionalComparison.matchFillColors -eq $true) {
            $additionalCompareArgs.MatchFillColor = $true
        }
        if ($additionalComparison.matchStrokeColors -eq $true) {
            $additionalCompareArgs.MatchStrokeColor = $true
        }
        if ($additionalComparison.matchOperators -eq $true) {
            $additionalCompareArgs.MatchOperator = $true
        }
        if ($additionalComparison.matchSegmentCounts -eq $true) {
            $additionalCompareArgs.MatchSegmentCount = $true
        }
        if ($additionalComparison.matchPathCommandCounts -eq $true) {
            $additionalCompareArgs.MatchPathCommandCounts = $true
        }
        if ($additionalComparison.matchPathOperators -eq $true) {
            $additionalCompareArgs.MatchPathOperators = $true
        }

        & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @additionalCompareArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Additional PDF chart graphics structure gate failed."
        }
    }
}

$hasChartTextGlobalTolerance = $manifest.expected.maxChartTextStructurePositionDelta -ne $null
$hasChartTextKindTolerances = $manifest.expected.maxChartTextStructurePositionDeltaByKind -ne $null
if ($hasChartTextGlobalTolerance -or $hasChartTextKindTolerances) {
    $referenceTextInspect = Get-ComparisonPdfInspection reference
    $candidateTextInspect = Get-ComparisonPdfInspection candidate
    $referenceTextOperations = Join-Path $referenceTextInspect 'text-operations.json'
    $candidateTextOperations = Join-Path $candidateTextInspect 'text-operations.json'

    $textPlotOptions = @{}
    if ($null -ne $manifest.expected.compareChartTextStructurePageNumber) {
        $textPlotOptions.PageNumber = [int]$manifest.expected.compareChartTextStructurePageNumber
    }
    if ($null -ne $manifest.expected.chartTextGraphicsGridlineBoundsTolerance) {
        $textPlotOptions.GridlineBoundsTolerance = [double]$manifest.expected.chartTextGraphicsGridlineBoundsTolerance
    }
    $matchesGraphics = $null -ne $chartGraphicsOptions -and $chartGraphicsOptions.Count -eq $textPlotOptions.Count -and
        !@($textPlotOptions.Keys | Where-Object {
            !$chartGraphicsOptions.ContainsKey($_) -or $chartGraphicsOptions[$_] -ne $textPlotOptions[$_]
        }).Count
    if (!$matchesGraphics) {
        $referenceChartStructures = Get-ComparisonChartGraphics reference $textPlotOptions 'chart-text-plot-structures.json'
        $candidateChartStructures = Get-ComparisonChartGraphics candidate $textPlotOptions 'chart-text-plot-structures.json'
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
        BoundsTolerance = 0d
        LineWidthTolerance = 0
        MatchByBounds = $true
    }
    if ($hasChartTextGlobalTolerance) {
        $compareChartTextArgs.BoundsTolerance = [double]$manifest.expected.maxChartTextStructurePositionDelta
    }
    if ($manifest.expected.compareChartTextStructureKinds -ne $null) {
        $compareChartTextArgs.Kinds = @($manifest.expected.compareChartTextStructureKinds)
    }
    else {
        $compareChartTextArgs.Kinds = @(
            "ChartTitleText",
            "CategoryAxisTickLabel",
            "ValueAxisTickLabel",
            "DataLabelText",
            "LegendText",
            "AbovePlotText",
            "BelowPlotText",
            "InsidePlotText",
            "LeftAxisText",
            "RightSideText",
            "OuterChartText",
            "ChartText")
    }
    if ($manifest.expected.compareChartTextStructureTextHash -eq $true) {
        $compareChartTextArgs.MatchTextHash = $true
    }

    if ($hasChartTextGlobalTolerance) {
        & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @compareChartTextArgs
        if ($LASTEXITCODE -ne 0) {
            throw "PDF chart text structure gate failed."
        }
    }

    if ($hasChartTextKindTolerances) {
        foreach ($entry in $manifest.expected.maxChartTextStructurePositionDeltaByKind.PSObject.Properties) {
            $kindTolerance = [double]$entry.Value
            $kindCompareArgs = $compareChartTextArgs.Clone()
            $kindCompareArgs.BoundsTolerance = $kindTolerance
            $kindCompareArgs.Kinds = @([string]$entry.Name)
            & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") @kindCompareArgs
            if ($LASTEXITCODE -ne 0) {
                throw "PDF chart text structure gate failed for kind '$($entry.Name)'."
            }
        }
    }
}

$requiresCandidateFontInspection =
    $manifest.expected.candidateFontsMustBeEmbedded -eq $true -or
    $manifest.expected.candidateFontsMustHaveToUnicode -eq $true -or
    $manifest.expected.candidateFontsMustBeSubset -eq $true
if ($requiresCandidateFontInspection) {
    $candidateFontResources = Join-Path (Get-ComparisonPdfInspection candidate) 'font-resources.json'

    $readFontResources = Read-JsonArray $candidateFontResources
    $fontResources = @($readFontResources)
    if ($fontResources.Count -eq 0) {
        throw "PDF font gate failed because the candidate PDF had no page font resources."
    }

    if ($manifest.expected.candidateFontsMustBeEmbedded -eq $true) {
        $missingEmbeddedFonts = @($fontResources | Where-Object { $_.Embedded -ne $true })
        if ($missingEmbeddedFonts.Count -ne 0) {
            $names = ($missingEmbeddedFonts | ForEach-Object { "p$($_.PageNumber)/$($_.ResourceName):$($_.BaseFont)" }) -join ", "
            throw "PDF font embedding gate failed. Non-embedded candidate fonts: $names."
        }
    }

    if ($manifest.expected.candidateFontsMustHaveToUnicode -eq $true) {
        $missingToUnicodeFonts = @($fontResources | Where-Object { $_.HasToUnicode -ne $true })
        if ($missingToUnicodeFonts.Count -ne 0) {
            $names = ($missingToUnicodeFonts | ForEach-Object { "p$($_.PageNumber)/$($_.ResourceName):$($_.BaseFont)" }) -join ", "
            throw "PDF font ToUnicode gate failed. Candidate fonts without ToUnicode maps: $names."
        }
    }

    if ($manifest.expected.candidateFontsMustBeSubset -eq $true) {
        $nonSubsetFonts = @($fontResources | Where-Object { $_.IsSubset -ne $true })
        if ($nonSubsetFonts.Count -ne 0) {
            $names = ($nonSubsetFonts | ForEach-Object { "p$($_.PageNumber)/$($_.ResourceName):$($_.BaseFont)" }) -join ", "
            throw "PDF font subsetting gate failed. Non-subset candidate fonts: $names."
        }
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

$allowedUnsupportedFeatures = @($manifest.allowedUnsupportedFeatures | Where-Object { $_ -is [string] -and $_.Length -ne 0 })
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

$requiredDiagnostics = @($manifest.expected.requiredDiagnostics | Where-Object { $_ -is [string] -and $_.Length -ne 0 })
if ($requiredDiagnostics.Count -ne 0) {
    $missingDiagnostics = @($requiredDiagnostics | Where-Object {
        $required = $_
        -not ($diagnosticItems | Where-Object {
            $_.Id -eq $required -or $_.Feature -eq $required
        } | Select-Object -First 1)
    })
    if ($missingDiagnostics.Count -ne 0) {
        throw "Missing required diagnostics: $($missingDiagnostics -join ', ')."
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
