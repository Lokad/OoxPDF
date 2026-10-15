# Independent Office comparisons for a measured portfolio; no library builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CandidateDirectory,
    [Parameter(Mandatory=$true)] [string] $ReferenceDirectory,
    [Parameter(Mandatory=$true)] [string] $VisualDiffPath,
    [string] $ValidationToolsDirectory = $PSScriptRoot
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
$rows = @()
foreach ($row in $outcomes.Rows) {
    $id = $row.Id + '-' + $row.Mode
    $run = Join-Path $candidate $id
    $source = Join-Path $reference $id
    $pdf = Join-Path $run 'candidate.pdf'
    $referencePdf = Join-Path $source 'reference.pdf'
    if ((Get-FileHash $pdf).Hash -ne $row.PdfSha256) { throw ('Candidate identity differs: ' + $id) }
    $referenceHash = (Get-FileHash $referencePdf).Hash
    foreach ($item in @(@{Pdf=$pdf;Folder=(Join-Path $run 'raster')},@{Pdf=$referencePdf;Folder=(Join-Path $source 'raster')})) {
        if (!(Test-Path -LiteralPath $item.Folder)) {
            & (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1') -InputPdf $item.Pdf -OutputDirectory $item.Folder -Dpi 144 -SkipBuild *> ($item.Folder+'.log')
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
    DocxThresholdSourceSha256=(Get-FileHash $docxThresholdSource).Hash
    PptxThresholdSourceSha256=(Get-FileHash $pptxThresholdSource).Hash
    Scope='Existing numerical visual thresholds retained; no claim of complete content or layout fidelity'
    Rows=$rows
} | ConvertTo-Json -Depth 14 | Set-Content (Join-Path $candidate 'office-comparisons.json') -Encoding utf8
