# Independent Office comparisons for a measured portfolio; no library builds.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
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
function Get-ComparisonFailures([object[]] $Metrics, $Limits, $Outcome, [string] $Id) {
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
        foreach ($field in @('MeanAbsoluteError','ChangedPixelRatioAtThreshold16','ForegroundRecall','StructuralSimilarity')) {
            $value = $page.$field
            # VisualDiff emits null pixel metrics for incompatible dimensions,
            # and null recall when the reference has no foreground pixels.
            if ($null -eq $value) {
                if ($page.DimensionsMatch -and $field -ne 'ForegroundRecall') { throw ($incomplete + ': ' + $field) }
            }
            elseif ($value -isnot [ValueType] -or $value -is [bool] -or ![double]::IsFinite([double]$value)) {
                throw ($incomplete + ': ' + $field)
            }
        }
        if ($null -ne $page.MeanAbsoluteError -and $page.MeanAbsoluteError -gt $Limits.maxMeanAbsoluteError) { $failures += 'MeanAbsoluteError' }
        if ($null -ne $page.ChangedPixelRatioAtThreshold16 -and $page.ChangedPixelRatioAtThreshold16 -gt $Limits.maxChangedPixelRatioAtThreshold16) { $failures += 'ChangedPixels' }
        if ($null -ne $page.ForegroundRecall -and $page.ForegroundRecall -lt $Limits.minForegroundRecall) { $failures += 'ForegroundRecall' }
        if ($null -ne $page.StructuralSimilarity -and $Limits.minStructuralSimilarity -and $page.StructuralSimilarity -lt $Limits.minStructuralSimilarity) { $failures += 'StructuralSimilarity' }
    }
    if (!$candidatePages -or !$referencePages -or
        ($null -ne $Outcome.InspectedPages -and $candidatePages -ne $Outcome.InspectedPages)) { throw $incomplete }
    return @($failures | Sort-Object -Unique)
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
    $limits = if ($row.Mode -eq 'presentation') {$thresholds.pptx} else {$thresholds.docx}
    $reuseFailures = @(if ($reuse.Count) {
        Get-ComparisonFailures -Metrics @($reuse[0].Row.Metrics) -Limits $limits -Outcome $row -Id $id
    })
    $diff = Join-Path $run 'office-diff'
    if (!$reuse.Count -and (Test-Path -LiteralPath $diff)) { throw ('Comparison already exists: ' + $id) }
    [pscustomobject]@{Row=$row;Id=$id;Run=$run;Source=$source;Pdf=$pdf;
        ReferencePdf=$referencePdf;ReferenceHash=$referenceHash;Identity=$selected.Identity;IdentityHash=$selected.IdentityHash;
        ExportLog=$selected.ExportLog;ExportLogHash=$selected.ExportLogHash;Reuse=$reuse;ReuseFailures=$reuseFailures;Limits=$limits;Diff=$diff}
})
# Match the reference cache's PDF/DPI/tool/page-hash convention. Legacy raster
# folders stay untouched: only derivatives completed by this command may be reused.
$rasterizerDll = Join-Path $ValidationToolsDirectory "Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework/Lokad.OoxPdf.PdfiumRasterizer.dll"
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
        & (Join-Path $ValidationToolsDirectory 'RasterizePdf.ps1') -Framework $Framework -InputPdf $Pdf -OutputDirectory $Folder -Dpi 144 -SkipBuild *> ($Folder+'.log')
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
            Failures=$job.ReuseFailures;Metrics=$prior.Metrics;Thresholds=$job.Limits;
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
    $limits = $job.Limits
    $failures = @(Get-ComparisonFailures -Metrics $metrics -Limits $limits -Outcome $row -Id $id)
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
