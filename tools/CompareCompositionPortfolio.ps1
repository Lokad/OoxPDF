# Independent Office comparisons for a measured portfolio; no library builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CandidateDirectory,
    [Parameter(Mandatory=$true)] [string] $ReferenceDirectory,
    [Parameter(Mandatory=$true)] [string] $VisualDiffPath,
    [string] $ValidationToolsDirectory = $PSScriptRoot,
    [string[]] $ReuseComparisonDirectory = @()
)
$ErrorActionPreference = 'Stop'
$candidate = (Resolve-Path -LiteralPath $CandidateDirectory).Path
$reference = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
$outcomes = Get-Content -LiteralPath (Join-Path $candidate 'outcomes.json') -Raw | ConvertFrom-Json
$repo = Split-Path -Parent $PSScriptRoot
$docxThresholdSource = Join-Path $repo 'visual-cases/cases/docx-math-normal-fraction-inline-arial-normal/case.json'
$pptxThresholdSource = Join-Path $repo 'visual-cases/cases/pptx-ladder-11-composite-chart-port/case.json'
$thresholds = @{
    docx=(Get-Content $docxThresholdSource -Raw | ConvertFrom-Json).expected
    pptx=(Get-Content $pptxThresholdSource -Raw | ConvertFrom-Json).expected
}
$thresholdHashes = @{
    docx=(Get-FileHash $docxThresholdSource).Hash
    pptx=(Get-FileHash $pptxThresholdSource).Hash
}
$cached = @(foreach ($directory in $ReuseComparisonDirectory) {
    $path = Join-Path (Resolve-Path -LiteralPath $directory).Path 'office-comparisons.json'
    $report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    if ($report.DocxThresholdSourceSha256 -eq $thresholdHashes.docx -and
        $report.PptxThresholdSourceSha256 -eq $thresholdHashes.pptx) {
        foreach ($prior in $report.Rows) { [pscustomobject]@{Path=$path;Row=$prior} }
    }
})
$rows = @()
foreach ($row in $outcomes.Rows) {
    $id = $row.Id + '-' + $row.Mode
    $run = Join-Path $candidate $id
    $source = Join-Path $reference $id
    $pdf = Join-Path $run 'candidate.pdf'
    $referencePdf = Join-Path $source 'reference.pdf'
    if ((Get-FileHash $pdf).Hash -ne $row.PdfSha256) { throw ('Candidate identity differs: ' + $id) }
    $referenceHash = (Get-FileHash $referencePdf).Hash
    $reuse = @($cached | Where-Object {
        $_.Row.Id -eq $row.Id -and $_.Row.Mode -eq $row.Mode -and
        $_.Row.CandidatePdfSha256 -eq $row.PdfSha256 -and $_.Row.ReferencePdfSha256 -eq $referenceHash
    } | Select-Object -First 1)
    if ($reuse.Count) {
        $prior = $reuse[0].Row
        $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;
            CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;
            Failures=$prior.Failures;Metrics=$prior.Metrics;Thresholds=$prior.Thresholds;
            ContentSeverity=$row.Severity;ReusedComparison=$reuse[0].Path;
            ReusedComparisonSha256=(Get-FileHash $reuse[0].Path).Hash}
        Write-Output ($id + ': reused identical PDFs and threshold sources')
        continue
    }
    foreach ($item in @(@{Pdf=$pdf;Folder=(Join-Path $run 'raster')},@{Pdf=$referencePdf;Folder=(Join-Path $source 'raster')})) {
        # A failed rasterizer startup can leave an empty output directory.
        # Retry that incomplete stage instead of reporting missing PDF pages.
        if (!(Test-Path -LiteralPath $item.Folder) -or
            !(Get-ChildItem -LiteralPath $item.Folder -Filter 'page-*.png' -File | Select-Object -First 1)) {
            & (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1') -InputPdf $item.Pdf -OutputDirectory $item.Folder -Dpi 144 -SkipBuild *> ($item.Folder+'.log')
        }
        if (!(Get-ChildItem -LiteralPath $item.Folder -Filter 'page-*.png' -File | Select-Object -First 1)) {
            throw ('Rasterization produced no pages: ' + $item.Pdf)
        }
    }
    $diff = Join-Path $run 'office-diff'
    if (Test-Path -LiteralPath $diff) { throw ('Comparison already exists: ' + $id) }
    dotnet $VisualDiffPath (Join-Path $source 'raster') (Join-Path $run 'raster') $diff *> (Join-Path $run 'office-diff.log')
    if ($LASTEXITCODE -ne 0) { throw ('Visual comparison failed: ' + $id) }
    $metrics = @(Get-Content (Join-Path $diff 'metrics.json') -Raw | ConvertFrom-Json)
    $limits = if ($row.Mode -eq 'presentation') {$thresholds.pptx} else {$thresholds.docx}
    $failures = @()
    foreach ($page in $metrics) {
        if (!$page.ReferenceFile -or !$page.CandidateFile) { $failures += 'Pagination'; continue }
        if (!$page.DimensionsMatch) { $failures += 'Dimensions' }
        if ($page.MeanAbsoluteError -gt $limits.maxMeanAbsoluteError) { $failures += 'MeanAbsoluteError' }
        if ($page.ChangedPixelRatioAtThreshold16 -gt $limits.maxChangedPixelRatioAtThreshold16) { $failures += 'ChangedPixels' }
        if ($page.ForegroundRecall -lt $limits.minForegroundRecall) { $failures += 'ForegroundRecall' }
        if ($limits.minStructuralSimilarity -and $page.StructuralSimilarity -lt $limits.minStructuralSimilarity) { $failures += 'StructuralSimilarity' }
    }
    $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;Failures=@($failures | Sort-Object -Unique);Metrics=$metrics;Thresholds=$limits;ContentSeverity=$row.Severity}
    if ((Get-FileHash $referencePdf).Hash -ne $referenceHash) { throw 'Reference changed during comparison' }
    Write-Output ($id + ': ' + ($failures | Sort-Object -Unique) -join ', ')
}
[ordered]@{
    CandidateLibrarySha256=$outcomes.LibrarySha256
    DocxThresholdSourceSha256=$thresholdHashes.docx
    PptxThresholdSourceSha256=$thresholdHashes.pptx
    Scope='Existing numerical visual thresholds retained; no claim of complete content or layout fidelity'
    Rows=$rows
} | ConvertTo-Json -Depth 14 | Set-Content (Join-Path $candidate 'office-comparisons.json') -Encoding utf8
