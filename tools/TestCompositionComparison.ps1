# Exercise comparison preflight and identity reuse without rendering or building.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $ComparisonScript = (Join-Path $PSScriptRoot 'CompareCompositionPortfolio.ps1')
)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/comparison-test-' + [Guid]::NewGuid().ToString('N'))
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
$tools = Join-Path $repo 'tools'
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$script = Join-Path $tools 'CompareCompositionPortfolio.ps1'
Copy-Item -LiteralPath $ComparisonScript -Destination $script
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ReferenceCache.ps1') -Destination $tools
@'
param($InputPdf, $OutputDirectory, $Dpi, [switch]$SkipBuild)
'Unexpected raster stage' | Set-Content -LiteralPath ($OutputDirectory + '.called')
throw 'Preflight allowed a rendering stage'
'@ | Set-Content -LiteralPath (Join-Path $tools 'RasterizePdf.ps1')
$thresholdNames = @('docx-math-normal-fraction-inline-arial-normal', 'pptx-ladder-11-composite-chart-port')
foreach ($name in $thresholdNames) {
    $relative = 'visual-cases/cases/' + $name + '/case.json'
    $target = Join-Path $repo $relative
    New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path (Split-Path $PSScriptRoot) $relative) -Destination $target
}
function New-Case([string] $Name, [string] $Fault) {
    $directory = Join-Path $OutputDirectory $Name
    $candidate = Join-Path $directory 'candidate'
    $reference = Join-Path $directory 'reference'
    $rows = @(foreach ($id in @('first', 'second')) {
        $folder = $id + '-final-preserve'
        $run = Join-Path $candidate $folder
        $source = Join-Path $(if ($Fault.StartsWith('multi-', [StringComparison]::Ordinal) -and $id -eq 'second') {
            Join-Path $directory 'reference-second'
        } else { $reference }) $folder
        New-Item -ItemType Directory -Path $run, $source -Force | Out-Null
        $pdf = Join-Path $run 'candidate.pdf'
        ('Fake candidate ' + $id) | Set-Content -LiteralPath $pdf
        if ($Fault -ne 'missing-reference' -or $id -ne 'second') {
            ('Fake reference ' + $id) | Set-Content -LiteralPath (Join-Path $source 'reference.pdf')
        }
        $referencePdf = Join-Path $source 'reference.pdf'
        if ((Test-Path -LiteralPath $referencePdf) -and $Fault -ne 'legacy' -and
            !($Fault -eq 'multi-missing-pin' -and $id -eq 'second')) {
            $log = Join-Path $source 'export.log'
            ('Fake pinned export ' + $id) | Set-Content -LiteralPath $log
            $pin = @{Id=$id;Mode='final-preserve';InputSha256=('0' * 64);
                PdfSha256=(Get-FileHash -LiteralPath $referencePdf).Hash;ExportLogSha256=(Get-FileHash -LiteralPath $log).Hash}
            if ($id -eq 'second') {
                if ($Fault -eq 'multi-source') { $pin.InputSha256 = '1' * 64 }
                if ($Fault -eq 'multi-mode') { $pin.Mode = 'all-word-compatible' }
                if ($Fault -eq 'multi-reference-hash') { $pin.PdfSha256 = '1' * 64 }
                if ($Fault -eq 'multi-log') { 'Changed after pinning' | Add-Content -LiteralPath $log }
            }
            $pin | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $source 'source-identity.json')
        }
        [pscustomobject]@{Id=$id;Mode='final-preserve';HeldOut=$false;
            InputSha256=('0' * 64);PdfSha256=(Get-FileHash -LiteralPath $pdf).Hash;InspectedPages=1;Severity='Test fixture'}
    })
    if ($Fault -eq 'candidate-hash') { $rows[1].PdfSha256 = '0' * 64 }
    if ($Fault -eq 'existing-comparison') {
        $diff = Join-Path $candidate 'second-final-preserve/office-diff'
        New-Item -ItemType Directory -Path $diff | Out-Null
        'Retain this evidence' | Set-Content -LiteralPath (Join-Path $diff 'sentinel.txt')
    }
    if ($Fault -eq 'duplicate') { $rows = @($rows[0], $rows[0]) }
    @{LibrarySha256=('1' * 64);Rows=$rows} | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath (Join-Path $candidate 'outcomes.json')
    return @{Directory=$directory;CandidateDirectory=$candidate;ReferenceDirectory=$reference;SecondReferenceDirectory=(Join-Path $directory 'reference-second');Rows=$rows}
}
function Snapshot([string] $Directory) {
    return @(Get-ChildItem -LiteralPath $Directory -Recurse -File | Sort-Object FullName |
        ForEach-Object { $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName).Hash }) -join "`n"
}
$passed = @()
foreach ($fault in @('missing-reference', 'candidate-hash', 'existing-comparison', 'duplicate')) {
    $case = New-Case $fault $fault
    $before = Snapshot $case.Directory
    $rejected = $false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null
    }
    catch {
        $expected = switch ($fault) {
            'missing-reference' { 'second-final-preserve' }
            'candidate-hash' { 'Candidate identity differs: second-final-preserve' }
            'existing-comparison' { 'Comparison already exists: second-final-preserve' }
            'duplicate' { 'Duplicate comparison outcome: first-final-preserve' }
        }
        if (!$_.Exception.Message.Contains($expected)) { throw }
        $rejected = $true
    }
    if (!$rejected -or (Snapshot $case.Directory) -ne $before) {
        throw ('Preflight must reject without stage mutations: ' + $fault)
    }
    $passed += $fault
}
function New-ReuseCache($case) {
    $cache = Join-Path $case.Directory 'prior'
    New-Item -ItemType Directory -Path $cache | Out-Null
    $priorRows = @($case.Rows | ForEach-Object {
        @{Id=$_.Id;Mode=$_.Mode;CandidatePdfSha256=$_.PdfSha256;
            ReferencePdfSha256=(Get-FileHash -LiteralPath $(
                $relative = $_.Id+'-'+$_.Mode+'/reference.pdf'
                $path = Join-Path $case.ReferenceDirectory $relative
                if (Test-Path -LiteralPath $path) { $path } else { Join-Path $case.SecondReferenceDirectory $relative }
            )).Hash;
            Failures=@('ForegroundRecall');Metrics=@(@{
                Page=1;ReferenceFile='page-001.png';CandidateFile='page-001.png';DimensionsMatch=$true;
                MeanAbsoluteError=0;ChangedPixelRatioAtThreshold16=0;ForegroundRecall=.5;StructuralSimilarity=1
            });Thresholds=@{minForegroundRecall=.99}}
    })
    @{DocxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[0]+'/case.json'))).Hash;
        PptxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[1]+'/case.json'))).Hash;
        Rows=$priorRows} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $cache 'office-comparisons.json')
    return $cache
}
$case = New-Case 'identity-reuse' 'legacy'
$cache = New-ReuseCache $case
& $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
    -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
$report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or $_.Failures.Count -ne 1 -or $_.Failures[0] -ne 'ForegroundRecall'
}).Count -or @(Get-ChildItem -LiteralPath $case.Directory -Recurse -Filter '*.called').Count) {
    throw 'Identical comparisons must reuse retained failures without rendering'
}
$passed += 'identity-reuse'
Copy-Item (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1') -Destination $tools
function New-ImageCase([string] $Name) {
    $case=New-Case $Name 'image'
    foreach ($row in $case.Rows) {
        $source=Join-Path $case.ReferenceDirectory ($row.Id+'-'+$row.Mode)
        $png=[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a9WQAAAAASUVORK5CYII=')
        [IO.File]::WriteAllBytes((Join-Path $source 'page-001.png'),$png)
        @{Schema=1;ReferenceKind='OfficeImage';InputSha256=$row.InputSha256;OfficeBuild='control';Dpi=144;
            CalibrationStatus='Uncalibrated';SecondaryPdfSha256=(Get-FileHash (Join-Path $source 'reference.pdf')).Hash;
            Capture=@{Method='managed control';Pages=@(@{Page=1;File='page-001.png';Width=1;Height=1;
                WidthPoints=0.5;HeightPoints=0.5;Sha256=(Get-FileHash (Join-Path $source 'page-001.png')).Hash})}} |
            ConvertTo-Json -Depth 8 | Set-Content (Join-Path $source 'reference-image-metadata.json')
        $identity=Join-Path $source 'source-identity.json'
        $pin=Get-Content $identity -Raw | ConvertFrom-Json -AsHashtable
        $pin.ReferenceKind='OfficeImage';$pin.ImageMetadataSha256=(Get-FileHash (Join-Path $source 'reference-image-metadata.json')).Hash
        $pin | ConvertTo-Json | Set-Content $identity
    }
    $cache=New-ReuseCache $case
    $prior=Get-Content (Join-Path $cache 'office-comparisons.json') -Raw | ConvertFrom-Json -AsHashtable
    foreach ($row in $prior.Rows) {
        $row.ReferenceImageMetadataSha256=(Get-FileHash (Join-Path $case.ReferenceDirectory ($row.Id+'-'+$row.Mode+'/reference-image-metadata.json'))).Hash
    }
    $prior | ConvertTo-Json -Depth 9 | Set-Content (Join-Path $cache 'office-image-comparisons.json')
    $case.Cache=$cache
    return $case
}
$case=New-ImageCase 'image-identity-reuse'
& $script -ReferenceKind OfficeImage -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $case.Cache | Out-Null
$report=Get-Content (Join-Path $case.CandidateDirectory 'office-image-comparisons.json') -Raw | ConvertFrom-Json
if ($report.ReferenceKind -ne 'OfficeImage' -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or !$_.ReferenceImageMetadataSha256 -or $_.Failures -notcontains 'ForegroundRecall' -or
    $_.ReferenceCalibrationStatus -ne 'Uncalibrated'
}).Count -or @(Get-ChildItem $case.Directory -Recurse -Filter '*.called').Count) {throw 'Image reuse lost provenance, failures or calibration status'}
$passed+='image-identity-reuse'
foreach ($fault in @('image-as-pdf','changed-image-metadata')) {
    $case=New-ImageCase $fault
    $arguments=@{}
    if ($fault -eq 'changed-image-metadata') {
        $arguments.ReferenceKind='OfficeImage'
        $path=Join-Path $case.ReferenceDirectory 'second-final-preserve/reference-image-metadata.json'
        [IO.File]::AppendAllText($path,' ')
    }
    $before=Snapshot $case.Directory;$rejected=$false
    try {& $script @arguments -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null}
    catch {if ($_.Exception.Message -notmatch 'Reference kind differs|Image reference metadata differs') {throw};$rejected=$true}
    if (!$rejected -or (Snapshot $case.Directory) -cne $before) {throw 'Image reference mismatch mutated comparison stages'}
    $passed+=$fault
}
foreach ($fault in @('empty-metrics','missing-metric','null-error','nonfinite-metric','boolean-metric',
    'missing-page-index','duplicate-page','truncated-pages','wrong-file','nonboolean-dimensions','neither-file')) {
    $case = New-Case $fault 'legacy'
    $cache = New-ReuseCache $case
    $path = Join-Path $cache 'office-comparisons.json'
    $prior = Get-Content $path -Raw | ConvertFrom-Json -AsHashtable
    $metric = $prior.Rows[1].Metrics[0]
    switch ($fault) {
        'empty-metrics' { $prior.Rows[1].Metrics=@() }
        'missing-metric' { $metric.Remove('ForegroundRecall') }
        'null-error' { $metric.MeanAbsoluteError=$null }
        'nonfinite-metric' { $metric.ForegroundRecall=[double]::NaN }
        'boolean-metric' { $metric.MeanAbsoluteError=$false }
        'missing-page-index' { $metric.Remove('Page') }
        'duplicate-page' { $prior.Rows[1].Metrics=@($metric,$metric) }
        'truncated-pages' {
            $outcomesPath=Join-Path $case.CandidateDirectory 'outcomes.json'
            $outcomes=Get-Content $outcomesPath -Raw | ConvertFrom-Json
            $outcomes.Rows[1].InspectedPages=2
            $outcomes | ConvertTo-Json -Depth 6 | Set-Content $outcomesPath
        }
        'wrong-file' { $metric.ReferenceFile='page-002.png' }
        'nonboolean-dimensions' { $metric.DimensionsMatch='true' }
        'neither-file' { $metric.ReferenceFile=$null;$metric.CandidateFile=$null }
    }
    $prior.Rows[1].Failures=@()
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    $before=Snapshot $case.Directory;$rejected=$false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
    }
    catch {
        if ($_.Exception.Message -notlike 'Comparison page metrics incomplete: second-final-preserve*') { throw }
        $rejected=$true
    }
    if (!$rejected -or (Snapshot $case.Directory) -ne $before) {
        throw ('Incomplete metrics must refuse before mutations: '+$fault)
    }
    $passed+=$fault
}
foreach ($control in @('cleared-summary','obsolete-failure','blank-reference','dimension-mismatch',
    'near-dimensions-with-recall-failure','reference-extra-page','candidate-extra-page','threshold-boundary')) {
    $case=New-Case $control 'legacy'
    $cache=New-ReuseCache $case
    $path=Join-Path $cache 'office-comparisons.json'
    $prior=Get-Content $path -Raw | ConvertFrom-Json -AsHashtable
    $metric=$prior.Rows[1].Metrics[0]
    $expectedFailures=@()
    $prior.Rows[1].Failures=@()
    $prior.Rows[1].Thresholds=@{minForegroundRecall=0}
    switch ($control) {
        'cleared-summary' { $expectedFailures=@('ForegroundRecall') }
        'obsolete-failure' { $prior.Rows[1].Failures=@('MeanAbsoluteError');$metric.ForegroundRecall=1 }
        'blank-reference' { $metric.ForegroundRecall=$null }
        'dimension-mismatch' {
            $metric.DimensionsMatch=$false
            foreach ($field in @('MeanAbsoluteError','ChangedPixelRatioAtThreshold16','ForegroundRecall','StructuralSimilarity')) { $metric[$field]=$null }
            $expectedFailures=@('Dimensions')
        }
        'near-dimensions-with-recall-failure' { $metric.DimensionsMatch=$false;$expectedFailures=@('Dimensions','ForegroundRecall') }
        {$_ -in @('reference-extra-page','candidate-extra-page')} {
            $metric.ForegroundRecall=1
            $extra=@{Page=2;DimensionsMatch=$false;ReferenceFile='page-002.png';CandidateFile='page-002.png';
                MeanAbsoluteError=$null;ChangedPixelRatioAtThreshold16=$null;ForegroundRecall=$null;StructuralSimilarity=$null}
            if ($control -eq 'reference-extra-page') { $extra.CandidateFile=$null }
            else {
                $extra.ReferenceFile=$null
                $outcomesPath=Join-Path $case.CandidateDirectory 'outcomes.json'
                $outcomes=Get-Content $outcomesPath -Raw | ConvertFrom-Json
                $outcomes.Rows[1].InspectedPages=2
                $outcomes | ConvertTo-Json -Depth 6 | Set-Content $outcomesPath
            }
            $prior.Rows[1].Metrics=@($metric,$extra)
            $expectedFailures=@('Pagination')
        }
        'threshold-boundary' {
            $limits=Get-Content (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[0]+'/case.json')) -Raw | ConvertFrom-Json
            $metric.MeanAbsoluteError=$limits.expected.maxMeanAbsoluteError
            $metric.ChangedPixelRatioAtThreshold16=$limits.expected.maxChangedPixelRatioAtThreshold16
            $metric.ForegroundRecall=$limits.expected.minForegroundRecall
            $metric.StructuralSimilarity=$limits.expected.minStructuralSimilarity
        }
    }
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    $before=Snapshot $case.ReferenceDirectory
    & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
        -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
    $result=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    $actual=$result.Rows[1]
    if ((@($actual.Failures) -join ',') -cne ($expectedFailures -join ',') -or !$actual.ReusedComparison -or
        $actual.Thresholds.minForegroundRecall -ne .99 -or (Snapshot $case.ReferenceDirectory) -ne $before -or
        @(Get-ChildItem -LiteralPath $case.Directory -Recurse -Filter '*.called').Count) {
        throw ('Reuse must derive current failures without rendering: '+$control)
    }
    $passed+=$control
}
foreach ($fault in @('multi-missing-pin','multi-source','multi-mode','multi-reference-hash','multi-log','multi-conflict')) {
    $case = New-Case $fault $fault
    if ($fault -eq 'multi-conflict') {
        $source = Join-Path $case.ReferenceDirectory 'first-final-preserve'
        $target = Join-Path $case.SecondReferenceDirectory 'first-final-preserve'
        Copy-Item -LiteralPath $source -Destination $target -Recurse
        'Different independent export' | Set-Content -LiteralPath (Join-Path $target 'reference.pdf')
        $identity = Join-Path $target 'source-identity.json'
        $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json
        $pin.PdfSha256 = (Get-FileHash -LiteralPath (Join-Path $target 'reference.pdf')).Hash
        $pin | ConvertTo-Json | Set-Content -LiteralPath $identity
    }
    $before = Snapshot $case.Directory; $rejected = $false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory @($case.ReferenceDirectory,$case.SecondReferenceDirectory) `
            -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null
    }
    catch {
        $expected = switch ($fault) {
            'multi-missing-pin' { 'Reference identity missing: second-final-preserve' }
            'multi-source' { 'Reference identity differs: second-final-preserve' }
            'multi-mode' { 'Reference identity differs: second-final-preserve' }
            'multi-reference-hash' { 'Reference identity differs: second-final-preserve' }
            'multi-log' { 'Reference export provenance differs: second-final-preserve' }
            'multi-conflict' { 'Conflicting references: first-final-preserve' }
        }
        if (!$_.Exception.Message.Contains($expected)) { throw }
        $rejected = $true
    }
    if (!$rejected -or (Snapshot $case.Directory) -ne $before) {
        throw ('Multi-cache preflight must reject without stage mutations: ' + $fault)
    }
    $passed += $fault
}
$case = New-Case 'multi-reuse' 'multi-split'
$cache = New-ReuseCache $case
# Identical duplicates are allowed; differing exports above are never chosen by order.
Copy-Item -LiteralPath (Join-Path $case.ReferenceDirectory 'first-final-preserve') `
    -Destination (Join-Path $case.SecondReferenceDirectory 'first-final-preserve') -Recurse
$referenceBefore = @((Snapshot $case.ReferenceDirectory),(Snapshot $case.SecondReferenceDirectory)) -join "`n"
& $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory @($case.ReferenceDirectory,$case.SecondReferenceDirectory) `
    -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
$report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
    !$_.ReferenceIdentitySha256 -or !$_.ReferenceExportLogSha256 -or !$_.ReusedComparison -or
    $_.Failures.Count -ne 1 -or $_.Failures[0] -ne 'ForegroundRecall'
}).Count -or $report.Rows[0].ReferenceSourceDirectory -ne (Join-Path $case.ReferenceDirectory 'first-final-preserve') -or
    $report.Rows[1].ReferenceSourceDirectory -ne (Join-Path $case.SecondReferenceDirectory 'second-final-preserve') -or
    (@((Snapshot $case.ReferenceDirectory),(Snapshot $case.SecondReferenceDirectory)) -join "`n") -ne $referenceBefore -or
    @(Get-ChildItem -LiteralPath $case.Directory -Recurse -Filter '*.called').Count) {
    throw 'Split reference caches must reuse the pinned failures and provenance without copies or rendering'
}
$passed += 'multi-reuse-identical-duplicates'
# Exercise the real raster-identity branches with tiny owned PDF/page stand-ins.
@'
param($InputPdf, $OutputDirectory, $Dpi, [switch]$SkipBuild)
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$count=if ([IO.File]::ReadAllText($InputPdf).Contains('two pages')) {2} else {1}
for ($i=1;$i -le $count;$i++) { [IO.File]::WriteAllText((Join-Path $OutputDirectory ('page-{0:000}.png' -f $i)),('owned page '+$i)) }
$global:LASTEXITCODE=0
'@ | Set-Content -LiteralPath (Join-Path $tools 'RasterizePdf.ps1')
$comparisonProbeState=@{DiffCalls=0}
function dotnet {
    $comparisonProbeState.DiffCalls++
    $reference=$args[1];$candidate=$args[2];$diff=$args[3]
    New-Item -ItemType Directory -Path $diff | Out-Null
    $names=@(@((Get-ChildItem $reference -Filter 'page-*.png').Name)+@((Get-ChildItem $candidate -Filter 'page-*.png').Name) | Sort-Object -Unique)
    $metrics=@(foreach ($name in $names) {
        @{Page=[int]([IO.Path]::GetFileNameWithoutExtension($name).Substring(5));
          ReferenceFile=$(if (Test-Path (Join-Path $reference $name)) {$name} else {$null});
          CandidateFile=$(if (Test-Path (Join-Path $candidate $name)) {$name} else {$null});
          DimensionsMatch=$true;MeanAbsoluteError=0;ChangedPixelRatioAtThreshold16=0;ForegroundRecall=1;StructuralSimilarity=1}
    })
    ConvertTo-Json -InputObject $metrics -Depth 5 | Set-Content (Join-Path $diff 'metrics.json')
    $global:LASTEXITCODE=0
}
try {
    $baseline=New-Case 'complete-raster-baseline' 'legacy'
    & $script -CandidateDirectory $baseline.CandidateDirectory -ReferenceDirectory $baseline.ReferenceDirectory -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null
    foreach ($id in @('first','second')) {
        # Preserve the completed attempt while testing derivative reuse separately
        # from the existing-comparison refusal. All paths belong to this fixture.
        $diff=[IO.Path]::GetFullPath((Join-Path $baseline.CandidateDirectory ($id+'-final-preserve/office-diff')))
        $ownedRoot=[IO.Path]::GetFullPath($OutputDirectory)+[IO.Path]::DirectorySeparatorChar
        if (!$diff.StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture move outside test directory' }
        [IO.Directory]::Move($diff,$diff+'.completed')
    }
    foreach ($fault in @('valid-raster-reuse','stale-pdf','missing-page','changed-page','extra-page','wrong-dpi','wrong-tool','missing-pin')) {
        $directory=Join-Path $OutputDirectory $fault
        Copy-Item -LiteralPath $baseline.Directory -Destination $directory -Recurse
        $candidate=Join-Path $directory 'candidate';$reference=Join-Path $directory 'reference'
        $run=Join-Path $candidate 'first-final-preserve';$raster=Join-Path $run 'portfolio-raster'
        $pinPath=Join-Path $raster 'raster-identity.json';$pin=Get-Content $pinPath -Raw | ConvertFrom-Json
        switch ($fault) {
            'stale-pdf' {
                Add-Content (Join-Path $run 'candidate.pdf') 'two pages'
                $outcomes=Get-Content (Join-Path $candidate 'outcomes.json') -Raw | ConvertFrom-Json
                $outcomes.Rows[0].PdfSha256=(Get-FileHash (Join-Path $run 'candidate.pdf')).Hash
                $outcomes | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $candidate 'outcomes.json')
            }
            'missing-page' { Remove-Item -LiteralPath (Join-Path $raster 'page-001.png') }
            'changed-page' { Add-Content (Join-Path $raster 'page-001.png') 'changed' }
            'extra-page' { Set-Content (Join-Path $raster 'page-002.png') 'extra' }
            'wrong-dpi' { $pin.Dpi=72;$pin | ConvertTo-Json -Depth 5 | Set-Content $pinPath }
            'wrong-tool' { $pin.RasterizerId='wrong';$pin | ConvertTo-Json -Depth 5 | Set-Content $pinPath }
            'missing-pin' { Remove-Item -LiteralPath $pinPath }
        }
        $before=Snapshot $directory;$beforeCalls=$comparisonProbeState.DiffCalls;$rejected=$false
        try { & $script -CandidateDirectory $candidate -ReferenceDirectory $reference -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null }
        catch {
            if ($fault -eq 'valid-raster-reuse' -or $_.Exception.Message -notmatch '[Pp]ortfolio raster') { throw }
            $rejected=$true
        }
        if ($fault -eq 'valid-raster-reuse') {
            if ($rejected -or $comparisonProbeState.DiffCalls -ne $beforeCalls+2) { throw 'Valid pinned raster was not reused' }
        }
        elseif (!$rejected -or $comparisonProbeState.DiffCalls -ne $beforeCalls -or (Snapshot $directory) -ne $before) {
            throw ('Invalid raster reached comparison or mutated retained evidence: '+$fault)
        }
        $passed+=$fault
    }
    $case=New-Case 'legacy-raster-two-page-control' 'legacy'
    $run=Join-Path $case.CandidateDirectory 'first-final-preserve'
    Add-Content (Join-Path $run 'candidate.pdf') 'two pages'
    $outcomes=Get-Content (Join-Path $case.CandidateDirectory 'outcomes.json') -Raw | ConvertFrom-Json
    $outcomes.Rows[0].PdfSha256=(Get-FileHash (Join-Path $run 'candidate.pdf')).Hash
    $outcomes.Rows[0].InspectedPages=2
    $outcomes | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $case.CandidateDirectory 'outcomes.json')
    New-Item -ItemType Directory -Path (Join-Path $run 'raster') | Out-Null
    Set-Content (Join-Path $run 'raster/page-001.png') 'stale one-page legacy raster'
    & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath 'unused.dll' -ValidationToolsDirectory $tools | Out-Null
    $result=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    if ($result.Rows[0].Failures -notcontains 'Pagination') { throw 'Stale legacy raster concealed the second page' }
    $passed+='legacy-raster-cannot-conceal-pagination'
}
finally { Remove-Item Function:\dotnet }
@{Passed=$passed;LibraryBuilds=0;Conversions=0;NativeRasterizations=0;Scope='Fake PDF/page inputs; real preflight, reuse and comparison branches'} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
Write-Output ($passed.Count.ToString() + ' comparison orchestration checks passed; no build or native rendering.')
