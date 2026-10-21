# Independent Office comparisons for a measured portfolio; no library builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CandidateDirectory,
    [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [string[]] $ReferenceDirectory,
    [Parameter(Mandatory=$true)] [string] $VisualDiffPath,
    [string] $ValidationToolsDirectory = $PSScriptRoot,
    [string[]] $ReuseComparisonDirectory = @()
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReferenceCache.ps1')
$candidate = (Resolve-Path -LiteralPath $CandidateDirectory).Path
$references = @($ReferenceDirectory | ForEach-Object { (Resolve-Path -LiteralPath $_).Path } | Select-Object -Unique)
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
$ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
# Validate the complete batch before raster/diff stages create any output. A
# missing late reference must not strand completed early comparisons.
$jobs = @(foreach ($row in $outcomes.Rows) {
    $id = $row.Id + '-' + $row.Mode
    if (!$ids.Add($id)) { throw ('Duplicate comparison outcome: ' + $id) }
    $run = Join-Path $candidate $id
    $pdf = Join-Path $run 'candidate.pdf'
    if ((Get-FileHash $pdf).Hash -ne $row.PdfSha256) { throw ('Candidate identity differs: ' + $id) }
    $sources = @(foreach ($directory in $references) {
        $source = Join-Path $directory $id
        if (!(Test-Path -LiteralPath $source)) { continue }
        $referencePdf = Join-Path $source 'reference.pdf'
        if (!(Test-Path -LiteralPath $referencePdf -PathType Leaf)) { throw ('Reference PDF missing: ' + $id) }
        $referenceHash = (Get-FileHash -LiteralPath $referencePdf).Hash
        $identity = Join-Path $source 'source-identity.json'
        $identityHash = $null; $exportLog = $null; $exportLogHash = $null
        # Existing single-directory bare PDF comparisons remain usable. Selecting
        # across caches requires a pin, so stale or conflicting source/mode
        # entries cannot silently win by directory order.
        if ($references.Count -gt 1 -or (Test-Path -LiteralPath $identity)) {
            if (!(Test-Path -LiteralPath $identity -PathType Leaf)) { throw ('Reference identity missing: ' + $id) }
            $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json
            if ($pin.Id -cne $row.Id -or $pin.Mode -cne $row.Mode -or
                $pin.InputSha256 -ne $row.InputSha256 -or $pin.PdfSha256 -ne $referenceHash) {
                throw ('Reference identity differs: ' + $id)
            }
            $identityHash = (Get-FileHash -LiteralPath $identity).Hash
            if ($pin.ExportLogSha256) {
                $exportLog = Join-Path $source 'export.log'
                if (!(Test-Path -LiteralPath $exportLog -PathType Leaf) -or
                    (Get-FileHash -LiteralPath $exportLog).Hash -ne $pin.ExportLogSha256) {
                    throw ('Reference export provenance differs: ' + $id)
                }
                $exportLogHash = $pin.ExportLogSha256
            }
        }
        [pscustomobject]@{Source=$source;Pdf=$referencePdf;Hash=$referenceHash;
            Identity=$identity;IdentityHash=$identityHash;ExportLog=$exportLog;ExportLogHash=$exportLogHash}
    })
    if (!$sources.Count) { throw ('Reference missing: ' + $id) }
    if (@($sources.Hash | Select-Object -Unique).Count -ne 1) { throw ('Conflicting references: ' + $id) }
    $selected = $sources[0]
    $source = $selected.Source; $referencePdf = $selected.Pdf; $referenceHash = $selected.Hash
    $reuse = @($cached | Where-Object {
        $_.Row.Id -eq $row.Id -and $_.Row.Mode -eq $row.Mode -and
        $_.Row.CandidatePdfSha256 -eq $row.PdfSha256 -and $_.Row.ReferencePdfSha256 -eq $referenceHash
    } | Select-Object -First 1)
    $diff = Join-Path $run 'office-diff'
    if (!$reuse.Count -and (Test-Path -LiteralPath $diff)) { throw ('Comparison already exists: ' + $id) }
    [pscustomobject]@{Row=$row;Id=$id;Run=$run;Source=$source;Pdf=$pdf;
        ReferencePdf=$referencePdf;ReferenceHash=$referenceHash;Identity=$selected.Identity;IdentityHash=$selected.IdentityHash;
        ExportLog=$selected.ExportLog;ExportLogHash=$selected.ExportLogHash;Reuse=$reuse;Diff=$diff}
})
# Match the reference cache's PDF/DPI/tool/page-hash convention. Legacy raster
# folders stay untouched: only derivatives completed by this command may be reused.
$rasterizerDll = Join-Path $ValidationToolsDirectory 'Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/net10.0/Lokad.OoxPdf.PdfiumRasterizer.dll'
$rasterizerIdentity = [ordered]@{
    SourceAndNativeId = Get-RasterizerId
    WrapperSha256 = (Get-FileHash (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1')).Hash
    AssemblySha256 = if (Test-Path -LiteralPath $rasterizerDll) { (Get-FileHash $rasterizerDll).Hash } else { $null }
}
$rasterizerId = Get-ShortSha256 ([Text.Encoding]::UTF8.GetBytes(($rasterizerIdentity | ConvertTo-Json -Compress))) 32
function Get-PortfolioRaster([string] $Pdf, [string] $Folder) {
    $pdfHash = (Get-FileHash -LiteralPath $Pdf).Hash
    $identityPath = Join-Path $Folder 'raster-identity.json'
    $reuseRaster = Test-Path -LiteralPath $Folder
    if ($reuseRaster) {
        if (!(Test-Path -LiteralPath $identityPath -PathType Leaf)) { throw ('Unpinned or incomplete portfolio raster: ' + $Folder) }
        $pin = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
        if ($pin.PdfSha256 -ne $pdfHash -or $pin.Dpi -ne 144 -or $pin.RasterizerId -ne $rasterizerId) {
            throw ('Portfolio raster identity differs: ' + $Folder)
        }
    }
    else {
        & (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1') -InputPdf $Pdf -OutputDirectory $Folder -Dpi 144 -SkipBuild *> ($Folder+'.log')
        # The wrapper throws on native failure; Start-Process does not set LASTEXITCODE.
    }
    $files = @(Get-ChildItem -LiteralPath $Folder -Filter 'page-*.png' -File)
    if (!$files.Count -or @($files | Where-Object Name -CNotMatch '^page-[0-9]+\.png$').Count) { throw ('Portfolio raster page set is incomplete: ' + $Folder) }
    $files = @($files | Sort-Object { [int]$_.BaseName.Substring(5) })
    $pages = @(for ($index=0; $index -lt $files.Count; $index++) {
        if ($files[$index].Name -cne ('page-{0:000}.png' -f ($index+1))) { throw ('Portfolio raster page set is incomplete: ' + $Folder) }
        [ordered]@{Name=$files[$index].Name;Sha256=(Get-FileHash -LiteralPath $files[$index].FullName).Hash}
    })
    if ($reuseRaster) {
        if ($pin.Pages.Count -ne $pages.Count) { throw ('Portfolio raster page set differs: ' + $Folder) }
        for ($index=0; $index -lt $pages.Count; $index++) {
            if ($pin.Pages[$index].Name -cne $pages[$index].Name -or $pin.Pages[$index].Sha256 -ne $pages[$index].Sha256) {
                throw ('Portfolio raster page identity differs: ' + $Folder)
            }
        }
    }
    if ((Get-FileHash -LiteralPath $Pdf).Hash -ne $pdfHash) { throw ('Rasterized PDF changed: ' + $Pdf) }
    if (!$reuseRaster) {
        @{PdfSha256=$pdfHash;Dpi=144;RasterizerId=$rasterizerId;Pages=$pages} | ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath $identityPath -Encoding utf8
    }
    return $Folder
}
$rows = @()
foreach ($job in $jobs) {
    $row = $job.Row; $id = $job.Id; $run = $job.Run; $source = $job.Source
    $pdf = $job.Pdf; $referencePdf = $job.ReferencePdf; $referenceHash = $job.ReferenceHash
    $reuse = @($job.Reuse); $diff = $job.Diff
    if ((Get-FileHash -LiteralPath $pdf).Hash -ne $row.PdfSha256 -or
        (Get-FileHash -LiteralPath $referencePdf).Hash -ne $referenceHash) { throw ('Comparison input changed: ' + $id) }
    if ($job.IdentityHash -and (Get-FileHash -LiteralPath $job.Identity).Hash -ne $job.IdentityHash -or
        $job.ExportLogHash -and (Get-FileHash -LiteralPath $job.ExportLog).Hash -ne $job.ExportLogHash) {
        throw ('Comparison reference provenance changed: ' + $id)
    }
    if ($reuse.Count) {
        $prior = $reuse[0].Row
        $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;
            CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;
            ReferenceSourceDirectory=$source;ReferenceIdentitySha256=$job.IdentityHash;ReferenceExportLogSha256=$job.ExportLogHash;
            Failures=$prior.Failures;Metrics=$prior.Metrics;Thresholds=$prior.Thresholds;
            ContentSeverity=$row.Severity;ReusedComparison=$reuse[0].Path;
            ReusedComparisonSha256=(Get-FileHash $reuse[0].Path).Hash}
        Write-Output ($id + ': reused identical PDFs and threshold sources')
        continue
    }
    $candidateRaster = Get-PortfolioRaster $pdf (Join-Path $run 'portfolio-raster')
    $referenceRaster = Get-PortfolioRaster $referencePdf (Join-Path $source 'portfolio-raster')
    dotnet $VisualDiffPath $referenceRaster $candidateRaster $diff *> (Join-Path $run 'office-diff.log')
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
    $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;ReferenceSourceDirectory=$source;ReferenceIdentitySha256=$job.IdentityHash;ReferenceExportLogSha256=$job.ExportLogHash;Failures=@($failures | Sort-Object -Unique);Metrics=$metrics;Thresholds=$limits;ContentSeverity=$row.Severity}
    if ((Get-FileHash -LiteralPath $pdf).Hash -ne $row.PdfSha256 -or
        (Get-FileHash -LiteralPath $referencePdf).Hash -ne $referenceHash) { throw ('Comparison input changed: ' + $id) }
    Write-Output ($id + ': ' + ($failures | Sort-Object -Unique) -join ', ')
}
[ordered]@{
    CandidateLibrarySha256=$outcomes.LibrarySha256
    DocxThresholdSourceSha256=$thresholdHashes.docx
    PptxThresholdSourceSha256=$thresholdHashes.pptx
    Scope='Existing numerical visual thresholds retained; no claim of complete content or layout fidelity'
    Rows=$rows
} | ConvertTo-Json -Depth 14 | Set-Content (Join-Path $candidate 'office-comparisons.json') -Encoding utf8
