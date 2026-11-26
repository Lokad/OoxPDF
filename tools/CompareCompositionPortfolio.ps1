# Independent Office comparisons for a measured portfolio; no library builds.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory=$true)] [string] $CandidateDirectory,
    [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] [string[]] $ReferenceDirectory,
    [Parameter(Mandatory=$true)] [string] $VisualDiffPath,
    [string] $ValidationToolsDirectory = $PSScriptRoot,
    [string[]] $ReuseComparisonDirectory = @(),
    [ValidateSet('OfficePdf','OfficeImage')] [string] $ReferenceKind = 'OfficePdf',
    [string] $RasterizerPath,
    [string] $InspectorPath,
    [string] $ContentAuditPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReferenceCache.ps1')
. (Join-Path $PSScriptRoot 'CompositionReferenceIdentity.ps1')
. (Join-Path $PSScriptRoot 'CompositionOutcomes.ps1')
if ($ReferenceKind -eq 'OfficeImage') { . (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1') }
$reportName=if ($ReferenceKind -eq 'OfficeImage') {'office-image-comparisons.json'} else {'office-comparisons.json'}
$candidate = (Resolve-Path -LiteralPath $CandidateDirectory).Path
$references = @($ReferenceDirectory | ForEach-Object { (Resolve-Path -LiteralPath $_).Path } | Select-Object -Unique)
$outcomes = Get-Content -LiteralPath (Join-Path $candidate 'outcomes.json') -Raw | ConvertFrom-Json
if (!$outcomes.Rows) { throw 'A nonempty measured portfolio is required' }
$loadedAudits = @{}
$audit = if ($ContentAuditPath) { Read-PortfolioContentAudit $ContentAuditPath $loadedAudits }
$docxThresholdSource = Join-Path $PSScriptRoot 'CompositionDocxThresholds.json'
$pptxThresholdSource = Join-Path $PSScriptRoot 'CompositionPptxThresholds.json'
$thresholds = @{
    docx=(Get-Content $docxThresholdSource -Raw | ConvertFrom-Json)
    pptx=(Get-Content $pptxThresholdSource -Raw | ConvertFrom-Json)
}
$thresholdHashes = @{
    docx=(Get-FileHash $docxThresholdSource).Hash
    pptx=(Get-FileHash $pptxThresholdSource).Hash
}
$VisualDiffPath = (Resolve-Path -LiteralPath $VisualDiffPath).Path
$inspectorDll = if ($InspectorPath) { (Resolve-Path -LiteralPath $InspectorPath).Path }
    else { Join-Path $ValidationToolsDirectory "Lokad.OoxPdf.PdfInspect/bin/Debug/$Framework/Lokad.OoxPdf.PdfInspect.dll" }
$inspectorHash = if (Test-Path -LiteralPath $inspectorDll -PathType Leaf) { (Get-FileHash -LiteralPath $inspectorDll).Hash }
$inspectedPageCounts = @{}
function Get-PortfolioPdfPageCount([string] $Pdf, [string] $PdfHash) {
    if (!$inspectedPageCounts.ContainsKey($PdfHash)) {
        if (!$inspectorHash) { throw 'PDF inspector is missing. Supply an existing build with -InspectorPath.' }
        $lines = @(dotnet $inspectorDll $Pdf --page-count)
        if ($LASTEXITCODE -ne 0) { throw ('Could not inspect portfolio PDF pages: ' + $Pdf) }
        $pageCount = Get-LogicalPdfPageCount $lines $Pdf
        if ((Get-FileHash -LiteralPath $Pdf).Hash -cne $PdfHash) {
            throw ('Portfolio PDF changed during page inspection: ' + $Pdf)
        }
        $inspectedPageCounts[$PdfHash] = $pageCount
    }
    return $inspectedPageCounts[$PdfHash]
}
$rasterizerDll = if ($RasterizerPath) {(Resolve-Path -LiteralPath $RasterizerPath).Path}
    else {Join-Path $ValidationToolsDirectory "Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework/Lokad.OoxPdf.PdfiumRasterizer.dll"}
function Get-PortfolioRasterizerId {
    $assembly = if (Test-Path -LiteralPath $rasterizerDll -PathType Leaf) { $rasterizerDll }
    $identity = [ordered]@{
        RasterizerId = Get-RasterizerId $assembly
        WrapperSha256 = (Get-FileHash (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1')).Hash
    }
    Get-ShortSha256 ([Text.Encoding]::UTF8.GetBytes(($identity | ConvertTo-Json -Compress))) 32
}
$rasterizerId = Get-PortfolioRasterizerId
# Schema 1 means whole-page VisualDiff defaults, with no region specifications.
# A future change to those invocation semantics must change this schema.
$comparisonPolicy = [ordered]@{
    Schema=1;VisualDiffSha256=(Get-FileHash -LiteralPath $VisualDiffPath).Hash;RasterizerId=$rasterizerId
}
# Domains of the VisualDiff producer, independent of the active thresholds.
$metricRanges = [ordered]@{
    MeanAbsoluteError = @(0, 255)
    ChangedPixelRatioAtThreshold16 = @(0, 1)
    ForegroundRecall = @(0, 1)
    StructuralSimilarity = @(-1, 1)
}
function Get-ComparisonFailures([object[]] $Metrics, $Limits, [int] $CandidatePageCount, [int] $ReferencePageCount, [string] $Id) {
    $incomplete = 'Comparison page metrics incomplete: ' + $Id
    if (!$Metrics.Count) { throw $incomplete }
    $failures = @(); $candidatePages = 0; $referencePages = 0
    for ($index=0; $index -lt $Metrics.Count; $index++) {
        $page = $Metrics[$index]
        foreach ($field in @('Page','ReferenceFile','CandidateFile','DimensionsMatch',
            'MeanAbsoluteError','ChangedPixelRatioAtThreshold16','ForegroundRecall','StructuralSimilarity')) {
            if ($null -eq $page -or $field -notin $page.PSObject.Properties.Name) { throw ($incomplete + ': ' + $field) }
        }
        if (($page.Page -isnot [int] -and $page.Page -isnot [long]) -or $page.Page -ne $index+1 -or
            $page.DimensionsMatch -isnot [bool]) { throw $incomplete }
        $expectedFile = 'page-{0:000}.png' -f ($index+1)
        foreach ($field in @('ReferenceFile','CandidateFile')) {
            if ($null -ne $page.$field -and $page.$field -cne $expectedFile) { throw ($incomplete + ': ' + $field) }
        }
        if ($page.ReferenceFile) {
            $referencePages++
            if ($referencePages -ne $index+1) { throw $incomplete }
        }
        if ($page.CandidateFile) {
            $candidatePages++
            if ($candidatePages -ne $index+1) { throw $incomplete }
        }
        if (!$page.ReferenceFile -and !$page.CandidateFile) { throw $incomplete }
        if (!$page.ReferenceFile -or !$page.CandidateFile) { $failures += 'Pagination'; continue }
        if (!$page.DimensionsMatch) { $failures += 'Dimensions' }
        foreach ($field in $metricRanges.Keys) {
            $value = $page.$field
            # Incompatible dimensions can suppress all pixel metrics.
            if ($null -eq $value) {
                if ($page.DimensionsMatch -and $field -ne 'ForegroundRecall') { throw ($incomplete + ': ' + $field) }
            }
            elseif ($value -isnot [ValueType] -or $value -is [bool] -or ![double]::IsFinite([double]$value)) {
                throw ($incomplete + ': ' + $field)
            }
            elseif ($value -lt $metricRanges[$field][0] -or $value -gt $metricRanges[$field][1]) {
                throw ('Comparison page metric outside its range: ' + $Id + ' / ' + $field)
            }
        }
        if ($null -ne $page.MeanAbsoluteError -and $page.MeanAbsoluteError -gt $Limits.maxMeanAbsoluteError) { $failures += 'MeanAbsoluteError' }
        if ($null -ne $page.ChangedPixelRatioAtThreshold16 -and $page.ChangedPixelRatioAtThreshold16 -gt $Limits.maxChangedPixelRatioAtThreshold16) { $failures += 'ChangedPixels' }
        $foreground = $page.ReferenceForegroundPixels
        if ($null -ne $foreground -and
            (($foreground -isnot [int] -and $foreground -isnot [long]) -or $foreground -lt 0 -or
            (($foreground -eq 0) -ne ($null -eq $page.ForegroundRecall)))) {
            throw ('Comparison page recall disagrees with reference foreground pixels: ' + $Id)
        }
        if ($null -eq $page.ForegroundRecall) {
            # Older metrics can omit the count. A matched page requires evidence
            # of zero reference foreground pixels to leave recall inapplicable.
            if ($page.DimensionsMatch -and $null -eq $foreground) { $failures += 'ForegroundRecallUnverified' }
        }
        elseif ($page.ForegroundRecall -lt $Limits.minForegroundRecall) { $failures += 'ForegroundRecall' }
        if ($null -ne $page.StructuralSimilarity -and $null -ne $Limits.minStructuralSimilarity -and
            $page.StructuralSimilarity -lt $Limits.minStructuralSimilarity) { $failures += 'StructuralSimilarity' }
    }
    if ($candidatePages -ne $CandidatePageCount -or $referencePages -ne $ReferencePageCount) { throw $incomplete }
    return @($failures | Sort-Object -Unique)
}
$historical = @()
# Pixel metrics depend on the bound tools/PDFs/DPI, not the limit-file identity.
# Reapply the active limits below, retaining source hashes in the new report.
$cached = @(foreach ($directory in $ReuseComparisonDirectory) {
    $path = Join-Path (Resolve-Path -LiteralPath $directory).Path $reportName
    $bytes = [IO.File]::ReadAllBytes($path)
    $reader = [IO.StreamReader]::new([IO.MemoryStream]::new($bytes))
    try { $report = ConvertFrom-Json -InputObject ($reader.ReadToEnd()) }
    finally { $reader.Dispose() }
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    if ($report.ComparisonPolicy.Schema -eq $comparisonPolicy.Schema -and
        $report.ComparisonPolicy.VisualDiffSha256 -eq $comparisonPolicy.VisualDiffSha256 -and
        $report.ComparisonPolicy.RasterizerId -eq $comparisonPolicy.RasterizerId) {
        foreach ($prior in $report.Rows) { [pscustomobject]@{Path=$path;Sha256=$hash;Row=$prior} }
    }
    else { $historical += [ordered]@{Path=$path;Sha256=$hash;Reason='Unbound or different comparison policy; historical metrics only'} }
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
        $reference = Read-CompositionReference $source $row.Id $row.Mode $row.InputSha256 $ReferenceKind
        $imagePath = $null; $imageHash = ''; $imagePages = 0; $dpi = 144; $calibration = 'PDF-backed bootstrap'
        if ($ReferenceKind -eq 'OfficeImage') {
            $image = Test-OfficeImageReference $source $row.InputSha256 -ObservedPdfSha256 $reference.Hash
            $imagePath = Join-Path $source 'reference-image-metadata.json'
            $imageHash = (Get-FileHash -LiteralPath $imagePath).Hash
            if ($imageHash -ne $reference.Pin.ImageMetadataSha256) { throw ('Image reference metadata differs: ' + $id) }
            $imagePages = @($image.Capture.Pages).Count
            $dpi = [int]$image.Dpi; $calibration = $image.CalibrationStatus
        }
        [pscustomobject]@{Source=$source;Pdf=$reference.Pdf;Hash=$reference.Hash;
            Identity=$reference.Identity;IdentityHash=$reference.IdentityHash;
            ExportLog=$reference.ExportLog;ExportLogHash=$reference.ExportLogHash;
            ImagePath=$imagePath;ImageHash=$imageHash;ImagePages=$imagePages;Dpi=$dpi;Calibration=$calibration}
    })
    if (!$sources.Count) { throw ('Reference missing: ' + $id) }
    if (@($sources.Hash | Select-Object -Unique).Count -ne 1) { throw ('Conflicting references: ' + $id) }
    if (@($sources.ImageHash | Select-Object -Unique).Count -ne 1) { throw ('Conflicting image references: '+$id) }
    $selected = $sources[0]
    $source = $selected.Source; $referencePdf = $selected.Pdf; $referenceHash = $selected.Hash
    $witnesses = Get-PortfolioContentWitnesses $row $outcomes $audit $loadedAudits $referenceHash
    if ($witnesses.Audit -and (Get-FileHash -LiteralPath (Join-Path $run 'inspect/text-operations.json')).Hash -cne $row.TextOperationsSha256) {
        throw ('Portfolio content audit text inspection differs: ' + $id)
    }
    $pageInspection = Get-VerifiedPortfolioPageInspection (Join-Path $run 'inspect.log') $row
    $candidatePageCount = if ($pageInspection) { $pageInspection.PageCount }
        else { Get-PortfolioPdfPageCount $pdf $row.PdfSha256 }
    $referencePageCount = if ($ReferenceKind -eq 'OfficeImage') { $selected.ImagePages }
        else { Get-PortfolioPdfPageCount $referencePdf $referenceHash }
    $reuse = @($cached | Where-Object {
        $_.Row.Id -eq $row.Id -and $_.Row.Mode -eq $row.Mode -and
        $_.Row.CandidatePdfSha256 -eq $row.PdfSha256 -and $_.Row.ReferencePdfSha256 -eq $referenceHash -and
        [string]$_.Row.ReferenceImageMetadataSha256 -eq $selected.ImageHash -and
        $_.Row.ReferenceKind -eq $ReferenceKind -and $_.Row.RasterDpi -eq $selected.Dpi
    } | Select-Object -First 1)
    $limits = if ($row.Mode -eq 'presentation') {$thresholds.pptx} else {$thresholds.docx}
    $reuseFailures = @(if ($reuse.Count) {
        Get-ComparisonFailures @($reuse[0].Row.Metrics) $limits $candidatePageCount $referencePageCount $id
    })
    $diff = Join-Path $run $(if ($ReferenceKind -eq 'OfficeImage') {'office-image-diff'} else {'office-diff'})
    if (!$reuse.Count -and (Test-Path -LiteralPath $diff)) { throw ('Comparison already exists: ' + $id) }
    [pscustomobject]@{Row=$row;Id=$id;Run=$run;Source=$source;Pdf=$pdf;
        ReferencePdf=$referencePdf;ReferenceHash=$referenceHash;Identity=$selected.Identity;IdentityHash=$selected.IdentityHash;
        ExportLog=$selected.ExportLog;ExportLogHash=$selected.ExportLogHash;Reuse=$reuse;ReuseFailures=$reuseFailures;Limits=$limits;Diff=$diff;
        CandidatePageCount=$candidatePageCount;ReferencePageCount=$referencePageCount;ContentWitnesses=$witnesses;
        ImagePath=$selected.ImagePath;ImageHash=$selected.ImageHash;Dpi=$selected.Dpi;Calibration=$selected.Calibration}
})
# Match the reference cache's PDF/DPI/tool/page-hash convention. Legacy raster
# folders stay untouched: only derivatives completed by this command may be reused.
function Get-PortfolioRaster([string] $Pdf, [string] $Folder, [int] $Dpi=144) {
    $pdfHash = (Get-FileHash -LiteralPath $Pdf).Hash
    $identityPath = Join-Path $Folder 'raster-identity.json'
    $reuseRaster = Test-Path -LiteralPath $Folder
    if ($reuseRaster) {
        if (!(Test-Path -LiteralPath $identityPath -PathType Leaf)) { throw ('Unpinned or incomplete portfolio raster: ' + $Folder) }
        $pin = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
        if ($pin.PdfSha256 -ne $pdfHash -or $pin.Dpi -ne $Dpi -or $pin.RasterizerId -ne $rasterizerId) {
            throw ('Portfolio raster identity differs: ' + $Folder)
        }
    }
    else {
        $binding=@{}
        if ($RasterizerPath) {$binding.RasterizerPath=$rasterizerDll}
        & (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1') -Framework $Framework -InputPdf $Pdf -OutputDirectory $Folder -Dpi $Dpi -SkipBuild @binding *> ($Folder+'.log')
        # The wrapper throws on native failure; Start-Process does not set LASTEXITCODE.
    }
    try { $files = Get-ReferenceRasterPages $Folder }
    catch { throw ('Portfolio raster page set is incomplete: ' + $Folder) }
    if (!$files.Count) { throw ('Portfolio raster page set is incomplete: ' + $Folder) }
    $pages = @(foreach ($file in $files) {
        [ordered]@{Name=$file.Name;Sha256=(Get-FileHash -LiteralPath $file.FullName).Hash}
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
        @{PdfSha256=$pdfHash;Dpi=$Dpi;RasterizerId=$rasterizerId;Pages=$pages} | ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath $identityPath -Encoding utf8
    }
    return $Folder
}
$rows = @()
$usedReports = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
foreach ($job in $jobs) {
    $row = $job.Row; $id = $job.Id; $run = $job.Run; $source = $job.Source
    $pdf = $job.Pdf; $referencePdf = $job.ReferencePdf; $referenceHash = $job.ReferenceHash
    $reuse = @($job.Reuse); $diff = $job.Diff
    $contentSeverity = Get-PortfolioContentSeverity $row.MissingExpectedText $row.IgnoredEquationDiagnostics $job.ContentWitnesses
    if ((Get-FileHash -LiteralPath $pdf).Hash -ne $row.PdfSha256 -or
        (Get-FileHash -LiteralPath $referencePdf).Hash -ne $referenceHash) { throw ('Comparison input changed: ' + $id) }
    if ((Get-FileHash -LiteralPath $job.Identity).Hash -ne $job.IdentityHash -or
        $job.ExportLogHash -and (Get-FileHash -LiteralPath $job.ExportLog).Hash -ne $job.ExportLogHash) {
        throw ('Comparison reference provenance changed: ' + $id)
    }
    if ($reuse.Count) {
        $prior = $reuse[0]
        $usedReports[$prior.Path] = $prior.Sha256
        $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;
            CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;
            ReferenceSourceDirectory=$source;ReferenceIdentitySha256=$job.IdentityHash;ReferenceExportLogSha256=$job.ExportLogHash;
            ReferenceKind=$ReferenceKind;ReferenceImageMetadataSha256=$job.ImageHash;ReferenceCalibrationStatus=$job.Calibration;
            RasterDpi=$job.Dpi;
            Failures=$job.ReuseFailures;Metrics=$prior.Row.Metrics;Thresholds=$job.Limits;
            ContentSeverity=$contentSeverity;ContentWitnesses=$job.ContentWitnesses;ReusedComparison=$prior.Path;
            ReusedComparisonSha256=$prior.Sha256}
        Write-Output ("$id — $contentSeverity. Reused pixel metrics with current thresholds.")
        continue
    }
    $rasterFolder=if ($job.Dpi -eq 144) {'portfolio-raster'} else {'portfolio-raster-dpi'+$job.Dpi}
    $candidateRaster = Get-PortfolioRaster $pdf (Join-Path $run $rasterFolder) $job.Dpi
    $referenceRaster = if ($ReferenceKind -eq 'OfficeImage') {$source}
        else {Get-PortfolioRaster $referencePdf (Join-Path $source 'portfolio-raster')}
    dotnet $VisualDiffPath $referenceRaster $candidateRaster $diff *> (Join-Path $run 'office-diff.log')
    if ($LASTEXITCODE -ne 0) { throw ('Visual comparison failed: ' + $id) }
    $metrics = @(Get-Content (Join-Path $diff 'metrics.json') -Raw | ConvertFrom-Json)
    $limits = $job.Limits
    $failures = @(Get-ComparisonFailures $metrics $limits $job.CandidatePageCount $job.ReferencePageCount $id)
    $rows += [ordered]@{Id=$row.Id;Mode=$row.Mode;HeldOut=$row.HeldOut;CandidatePdfSha256=$row.PdfSha256;ReferencePdfSha256=$referenceHash;ReferenceSourceDirectory=$source;ReferenceIdentitySha256=$job.IdentityHash;ReferenceExportLogSha256=$job.ExportLogHash;ReferenceKind=$ReferenceKind;ReferenceImageMetadataSha256=$job.ImageHash;ReferenceCalibrationStatus=$job.Calibration;RasterDpi=$job.Dpi;Failures=@($failures | Sort-Object -Unique);Metrics=$metrics;Thresholds=$limits;ContentSeverity=$contentSeverity;ContentWitnesses=$job.ContentWitnesses}
    if ($ReferenceKind -eq 'OfficeImage') {
        if ((Get-FileHash $job.ImagePath).Hash -ne $job.ImageHash) { throw ('Image reference metadata changed: '+$id) }
        $null=Test-OfficeImageReference $source $row.InputSha256
    }
    if ((Get-FileHash -LiteralPath $pdf).Hash -ne $row.PdfSha256 -or
        (Get-FileHash -LiteralPath $referencePdf).Hash -ne $referenceHash) { throw ('Comparison input changed: ' + $id) }
    if ((Get-FileHash -LiteralPath $job.Identity).Hash -ne $job.IdentityHash -or
        $job.ExportLogHash -and (Get-FileHash -LiteralPath $job.ExportLog).Hash -ne $job.ExportLogHash) {
        throw ('Comparison reference provenance changed: ' + $id)
    }
    $visualVerdict = if ($failures.Count) { ($failures | Sort-Object -Unique) -join ', ' } else { 'Visual thresholds met' }
    Write-Output ("$id — $contentSeverity. $visualVerdict.")
}
if ((Get-FileHash -LiteralPath $VisualDiffPath).Hash -ne $comparisonPolicy.VisualDiffSha256 -or
    (Get-PortfolioRasterizerId) -ne $comparisonPolicy.RasterizerId -or
    ($inspectedPageCounts.Count -and (Get-FileHash -LiteralPath $inspectorDll).Hash -ne $inspectorHash)) {
    throw 'Comparison tools changed during the batch; retained attempts are not a qualified report'
}
foreach ($prior in $usedReports.GetEnumerator()) {
    if (!(Test-Path -LiteralPath $prior.Key -PathType Leaf) -or
        (Get-FileHash -LiteralPath $prior.Key).Hash -cne $prior.Value) {
        throw ('Reused comparison report changed: ' + $prior.Key)
    }
}
[ordered]@{
    CandidateLibrarySha256=$outcomes.LibrarySha256
    ReferenceKind=$ReferenceKind
    ComparisonPolicy=$comparisonPolicy
    PageCountInspectorSha256=if ($inspectedPageCounts.Count) { $inspectorHash }
    HistoricalComparisonReports=$historical
    DocxThresholdSource='tools/CompositionDocxThresholds.json'
    PptxThresholdSource='tools/CompositionPptxThresholds.json'
    DocxThresholdSourceSha256=$thresholdHashes.docx
    PptxThresholdSourceSha256=$thresholdHashes.pptx
    Scope='Visual metrics and separate content findings. Review conversion diagnostics in the measured outcomes.'
    Rows=$rows
} | ConvertTo-Json -Depth 14 | Set-Content (Join-Path $candidate $reportName) -Encoding utf8
