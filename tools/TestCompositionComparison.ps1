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
foreach ($name in @('CompositionReferenceIdentity.ps1','CompositionOutcomes.ps1','PdfTextGeometry.ps1','PathHelpers.ps1',
    'CompositionDocxThresholds.json','CompositionPptxThresholds.json')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $tools
}
. (Join-Path $tools 'ReferenceCache.ps1')
$visualDiff = Join-Path $tools 'comparison-engine.dll'
Set-Content -LiteralPath $visualDiff -Value 'Synthetic comparison engine A; never executed'
$inspectorDirectory = Join-Path $tools 'Lokad.OoxPdf.PdfInspect/bin/Debug/net11.0'
New-Item -ItemType Directory -Path $inspectorDirectory | Out-Null
$inspector = Join-Path $inspectorDirectory 'Lokad.OoxPdf.PdfInspect.dll'
Set-Content -LiteralPath $inspector -Value 'Synthetic PDF inspector; never executed'
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
function New-Case([string] $Name, [string] $Fault, [string] $Mode='final-preserve', [switch] $ExtraReferencePage) {
    $directory = Join-Path $OutputDirectory $Name
    $candidate = Join-Path $directory 'candidate'
    $reference = Join-Path $directory 'reference'
    $rows = @(foreach ($id in @('first', 'second')) {
        $folder = $id + '-' + $Mode
        $run = Join-Path $candidate $folder
        $source = Join-Path $(if ($Fault.StartsWith('multi-', [StringComparison]::Ordinal) -and $id -eq 'second') {
            Join-Path $directory 'reference-second'
        } else { $reference }) $folder
        New-Item -ItemType Directory -Path $run, $source -Force | Out-Null
        $pdf = Join-Path $run 'candidate.pdf'
        ('Fake candidate ' + $id) | Set-Content -LiteralPath $pdf
        if ($Fault -ne 'missing-reference' -or $id -ne 'second') {
            ('Fake reference ' + $id) | Set-Content -LiteralPath (Join-Path $source 'reference.pdf')
            if ($ExtraReferencePage -and $id -eq 'second') {
                Add-Content -LiteralPath (Join-Path $source 'reference.pdf') 'two pages'
            }
        }
        $referencePdf = Join-Path $source 'reference.pdf'
        if ((Test-Path -LiteralPath $referencePdf) -and
            !($Fault -in @('bare-reference','multi-missing-pin') -and $id -eq 'second')) {
            $log = Join-Path $source 'export.log'
            $settings = if ($Mode -eq 'presentation') { 'Fake PowerPoint export' }
                elseif ($Mode -eq 'all-word-compatible' -and !($Fault -eq 'wrong-review-view' -and $id -eq 'second')) {
                    'ExportSettings: Word/16.0 ShowRevisions=True RevisionsView=0 MarkupMode=2 RevisionsMode=2 ShowComments=True ShowFormatChanges=True ShowInsertionsAndDeletions=True RejectAll=False'
                } else {
                    'ExportSettings: Word/16.0 ShowRevisions=False RevisionsView=0 MarkupMode=0 RevisionsMode=0 ShowComments=False ShowFormatChanges=False ShowInsertionsAndDeletions=False RejectAll=False'
                }
            Set-Content -LiteralPath $log -Value $settings
            $pin = @{Id=$id;Mode=$Mode;InputSha256=('0' * 64);
                PdfSha256=(Get-FileHash -LiteralPath $referencePdf).Hash}
            if ($Fault -ne 'legacy') { $pin.ExportLogSha256 = (Get-FileHash -LiteralPath $log).Hash }
            if ($id -eq 'second') {
                if ($Fault -eq 'multi-source') { $pin.InputSha256 = '1' * 64 }
                if ($Fault -eq 'multi-mode') { $pin.Mode = 'all-word-compatible' }
                if ($Fault -eq 'multi-reference-hash') { $pin.PdfSha256 = '1' * 64 }
                if ($Fault -eq 'multi-log') { 'Changed after pinning' | Add-Content -LiteralPath $log }
                if ($Fault -eq 'missing-view-log') { Remove-Item -LiteralPath $log }
            }
            $pin | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $source 'source-identity.json')
        }
        [pscustomobject]@{Id=$id;Mode=$Mode;HeldOut=$false;
            InputSha256=('0' * 64);PdfSha256=(Get-FileHash -LiteralPath $pdf).Hash;InspectedPages=1;Severity='Test fixture'}
    })
    if ($Fault -eq 'candidate-hash') { $rows[1].PdfSha256 = '0' * 64 }
    if ($Fault -eq 'existing-comparison') {
        $diff = Join-Path $candidate 'second-final-preserve/office-diff'
        New-Item -ItemType Directory -Path $diff | Out-Null
        'Retain this evidence' | Set-Content -LiteralPath (Join-Path $diff 'sentinel.txt')
    }
    if ($Fault -eq 'duplicate') { $rows = @($rows[0], $rows[0]) }
    if ($Fault -eq 'empty') { $rows = @() }
    @{LibrarySha256=('1' * 64);Rows=$rows} | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath (Join-Path $candidate 'outcomes.json')
    return @{Directory=$directory;CandidateDirectory=$candidate;ReferenceDirectory=$reference;SecondReferenceDirectory=(Join-Path $directory 'reference-second');Rows=$rows}
}
function Snapshot([string] $Directory) {
    return @(Get-ChildItem -LiteralPath $Directory -Recurse -File | Sort-Object FullName |
        ForEach-Object { $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName).Hash }) -join "`n"
}
function Write-ControlPdfInspection([string] $Pdf) {
    $count = if ([IO.File]::ReadAllText($Pdf).Contains('two pages')) { 2 } else { 1 }
    for ($i=1; $i -le $count; $i++) { Write-Output "$i 0 obj: page" }
    $global:LASTEXITCODE = 0
}
function dotnet {
    if ($args[0] -ne $inspector -or $args[2] -ne '--text-only') { throw 'Unexpected tool invocation during comparison preflight' }
    Write-ControlPdfInspection $args[1]
}
$passed = @()
foreach ($fault in @('empty', 'missing-reference', 'candidate-hash', 'existing-comparison', 'duplicate')) {
    $case = New-Case $fault $fault
    $before = Snapshot $case.Directory
    $rejected = $false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null
    }
    catch {
        $expected = switch ($fault) {
            'empty' { 'A nonempty measured portfolio is required' }
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
        @{Id=$_.Id;Mode=$_.Mode;CandidatePdfSha256=$_.PdfSha256;ReferenceKind='OfficePdf';RasterDpi=144;
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
    # Historical wire-format reports bind the former visual-case owners. Their
    # pixel metrics stay reusable after relocation; active limits decide failures.
    @{DocxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[0]+'/case.json'))).Hash;
        PptxThresholdSourceSha256=(Get-FileHash (Join-Path $repo ('visual-cases/cases/'+$thresholdNames[1]+'/case.json'))).Hash;
        ComparisonPolicy=(Test-ComparisonPolicyFixture);Rows=$priorRows} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $cache 'office-comparisons.json')
    return $cache
}
function Test-ComparisonPolicyFixture {
    # A valid wire-format fixture; the checks below independently change its tools,
    # policy and DPI rather than accepting a cached pass/fail summary.
    $identity=[ordered]@{RasterizerId=(Get-RasterizerId);
        WrapperSha256=(Get-FileHash (Join-Path $tools 'RasterizePdf.ps1')).Hash}
    @{Schema=1;VisualDiffSha256=(Get-FileHash $visualDiff).Hash;
        RasterizerId=(Get-ShortSha256 ([Text.Encoding]::UTF8.GetBytes(($identity | ConvertTo-Json -Compress))) 32)}
}
$case = New-Case 'identity-reuse' 'legacy'
$cache = New-ReuseCache $case
& $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
    -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
$report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or $_.Failures.Count -ne 1 -or $_.Failures[0] -ne 'ForegroundRecall' -or
    $_.ContentSeverity -cne 'Content evidence incomplete'
}).Count -or @(Get-ChildItem -LiteralPath $case.Directory -Recurse -Filter '*.called').Count) {
    throw 'Identical comparisons must reuse retained failures without rendering'
}
$passed += 'identity-reuse'
$case = New-Case 'reused-content-severity' 'legacy'
$cache = New-ReuseCache $case
$priorPath = Join-Path $cache 'office-comparisons.json'
$prior = Get-Content $priorPath -Raw | ConvertFrom-Json -AsHashtable
foreach ($row in $prior.Rows) {
    $row.Metrics[0].ForegroundRecall = 1
    $row.ContentSeverity = 'Needs geometry and visual review'
}
$prior | ConvertTo-Json -Depth 10 | Set-Content $priorPath
$outcomePath = Join-Path $case.CandidateDirectory 'outcomes.json'
$outcomes = Get-Content $outcomePath -Raw | ConvertFrom-Json -AsHashtable
$outcomes.Rows[0].MissingExpectedText = @('Required phrase')
$outcomes.Rows[0].Severity = 'Needs geometry and visual review'
$outcomes.Rows[1].MissingExpectedText = @()
$outcomes.Rows[1].IgnoredEquationDiagnostics = 1
$outcomes.Rows[1].Remove('Severity')
$outcomes | ConvertTo-Json -Depth 6 | Set-Content $outcomePath
& $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
    -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
$report = Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or $_.Failures.Count -or $_.ContentSeverity -cne 'Content blocker'
}).Count) { throw 'Reused visual passes must retain missing-text and ignored-equation blockers' }
$passed += 'reused-content-severity'
foreach ($fault in @('wrong-review-view', 'bare-reference', 'missing-view-log')) {
    $case = New-Case $fault $fault 'all-word-compatible'
    $cache = New-ReuseCache $case
    $before = Snapshot $case.Directory
    $expected = switch ($fault) {
        'wrong-review-view' { 'Reference Office view differs: second-all-word-compatible / ShowRevisions' }
        'bare-reference' { 'Reference identity missing: second-all-word-compatible' }
        'missing-view-log' { 'Reference Office view is unverified: second-all-word-compatible' }
    }
    $rejected = $false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
    }
    catch { if ($_.Exception.Message -cne $expected) { throw }; $rejected = $true }
    if (!$rejected -or (Snapshot $case.Directory) -cne $before) {
        throw ('Reference provenance must be verified before metrics reuse: ' + $fault)
    }
    $passed += $fault
}
if (@($report.Rows | Where-Object { !$_.ReferenceExportLogSha256 }).Count) {
    throw 'Legacy Word references must publish the verified export log hash'
}
foreach ($control in @('pptx-boundary','pptx-mae','pptx-recall','pptx-changed-pixels')) {
    $case=New-Case $control 'legacy' 'presentation'
    $outcomePath = Join-Path $case.CandidateDirectory 'outcomes.json'
    if ($control -eq 'pptx-boundary') {
        $measured = Get-Content -LiteralPath $outcomePath -Raw | ConvertFrom-Json -AsHashtable
        $measured.Rows[1].Diagnostics = @(@{Id='PPTX_UNSUPPORTED_CALLOUT';Severity='Warning'})
        $measured | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outcomePath
    }
    $outcomeHash = (Get-FileHash -LiteralPath $outcomePath).Hash
    $cache=New-ReuseCache $case
    $path=Join-Path $cache 'office-comparisons.json'
    $prior=Get-Content $path -Raw | ConvertFrom-Json -AsHashtable
    foreach ($row in $prior.Rows) {
        $metric=$row.Metrics[0]
        $metric.MeanAbsoluteError=1.2;$metric.ForegroundRecall=.90
        $metric.ChangedPixelRatioAtThreshold16=.025;$metric.StructuralSimilarity=.4
        switch ($control) {
            'pptx-mae' {$metric.MeanAbsoluteError=1.21}
            'pptx-recall' {$metric.ForegroundRecall=.899}
            'pptx-changed-pixels' {$metric.ChangedPixelRatioAtThreshold16=.026}
        }
    }
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
        -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
    $report=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    $expected=switch ($control) {
        'pptx-boundary' {''} 'pptx-mae' {'MeanAbsoluteError'}
        'pptx-recall' {'ForegroundRecall'} 'pptx-changed-pixels' {'ChangedPixels'}
    }
    if (@($report.Rows | Where-Object {
        !$_.ReusedComparison -or ($_.Failures -join ',') -cne $expected -or
        $_.Thresholds.minForegroundRecall -ne .90 -or $_.Thresholds.maxMeanAbsoluteError -ne 1.2 -or
        (($_.Thresholds.PSObject.Properties.Name | Sort-Object) -join ',') -cne 'maxChangedPixelRatioAtThreshold16,maxMeanAbsoluteError,minForegroundRecall'
    }).Count -or (Get-FileHash -LiteralPath $outcomePath).Hash -ne $outcomeHash -or
        @(Get-ChildItem $case.Directory -Recurse -Filter '*.called').Count) {
        throw ('PPTX must retain its own limits without rerendering: '+$control)
    }
    $passed+=$control
}
foreach ($control in @('unrelated-case-edited','policy-bytes-changed','stricter-policy','zero-ssim-policy')) {
    $case=New-Case $control 'legacy'
    $cache=New-ReuseCache $case
    $path=Join-Path $cache 'office-comparisons.json'
    $prior=Get-Content $path -Raw | ConvertFrom-Json -AsHashtable
    foreach ($row in $prior.Rows) { $row.Metrics[0].ForegroundRecall=.995; $row.Failures=@() }
    if ($control -eq 'zero-ssim-policy') {
        $prior.Rows[0].Metrics[0].StructuralSimilarity=-.1
        $prior.Rows[1].Metrics[0].StructuralSimilarity=0
    }
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    $policyPath=Join-Path $tools 'CompositionDocxThresholds.json'
    $casePath=Join-Path $repo ('visual-cases/cases/'+$thresholdNames[0]+'/case.json')
    $policyBytes=[IO.File]::ReadAllBytes($policyPath);$caseBytes=[IO.File]::ReadAllBytes($casePath)
    try {
        switch ($control) {
            'unrelated-case-edited' {
                $unrelated=Get-Content $casePath -Raw | ConvertFrom-Json
                $unrelated.expected.minForegroundRecall=1
                $unrelated | ConvertTo-Json -Depth 8 | Set-Content $casePath
            }
            'policy-bytes-changed' { [IO.File]::AppendAllText($policyPath,"`n") }
            'stricter-policy' {
                $limits=Get-Content $policyPath -Raw | ConvertFrom-Json
                $limits.minForegroundRecall=.999
                $limits | ConvertTo-Json | Set-Content $policyPath
            }
            'zero-ssim-policy' {
                $limits=Get-Content $policyPath -Raw | ConvertFrom-Json
                $limits.minStructuralSimilarity=0
                $limits | ConvertTo-Json | Set-Content $policyPath
            }
        }
        $before=Snapshot $case.ReferenceDirectory
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
        $report=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
        $expectedRecall=if ($control -eq 'stricter-policy') {.999} else {.99}
        $expectedSsim=if ($control -eq 'zero-ssim-policy') {0} else {.985}
        foreach ($row in $report.Rows) {
            $expectedFailures=if ($control -eq 'stricter-policy') {'ForegroundRecall'}
                elseif ($control -eq 'zero-ssim-policy' -and $row.Id -eq 'first') {'StructuralSimilarity'} else {''}
            if (!$row.ReusedComparison -or ($row.Failures -join ',') -cne $expectedFailures -or
                $row.Thresholds.minForegroundRecall -ne $expectedRecall -or
                $row.Thresholds.minStructuralSimilarity -ne $expectedSsim -or
                (($row.Thresholds.PSObject.Properties.Name | Sort-Object) -join ',') -cne 'maxChangedPixelRatioAtThreshold16,maxMeanAbsoluteError,minForegroundRecall,minStructuralSimilarity') {
                throw ('Incorrect active threshold ownership: '+$control)
            }
        }
        if ($report.DocxThresholdSourceSha256 -ne (Get-FileHash $policyPath).Hash -or
            (Snapshot $case.ReferenceDirectory) -cne $before -or
            @(Get-ChildItem $case.Directory -Recurse -Filter '*.called').Count) { throw ('Policy update rerendered or lost provenance: '+$control) }
    }
    finally {
        [IO.File]::WriteAllBytes($policyPath,$policyBytes);[IO.File]::WriteAllBytes($casePath,$caseBytes)
    }
    $passed+=$control
}
Copy-Item (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1') -Destination $tools
function New-ImageCase([string] $Name, [switch] $ExtraReferencePage) {
    $case=New-Case $Name 'image'
    foreach ($row in $case.Rows) {
        $source=Join-Path $case.ReferenceDirectory ($row.Id+'-'+$row.Mode)
        $png=[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a9WQAAAAASUVORK5CYII=')
        $count = if ($ExtraReferencePage -and $row.Id -eq 'second') { 2 } else { 1 }
        $pages = @(for ($index=1; $index -le $count; $index++) {
            $name = 'page-{0:000}.png' -f $index
            $path = Join-Path $source $name
            [IO.File]::WriteAllBytes($path,$png)
            @{Page=$index;File=$name;Width=1;Height=1;WidthPoints=0.5;HeightPoints=0.5;Sha256=(Get-FileHash $path).Hash}
        })
        @{Schema=1;ReferenceKind='OfficeImage';InputSha256=$row.InputSha256;OfficeBuild='control';Dpi=144;
            CalibrationStatus='Uncalibrated';SecondaryPdfSha256=(Get-FileHash (Join-Path $source 'reference.pdf')).Hash;
            Capture=@{Method='managed control';Pages=$pages}} |
            ConvertTo-Json -Depth 8 | Set-Content (Join-Path $source 'reference-image-metadata.json')
        $identity=Join-Path $source 'source-identity.json'
        $pin=Get-Content $identity -Raw | ConvertFrom-Json -AsHashtable
        $pin.ReferenceKind='OfficeImage';$pin.ImageMetadataSha256=(Get-FileHash (Join-Path $source 'reference-image-metadata.json')).Hash
        $pin | ConvertTo-Json | Set-Content $identity
    }
    $cache=New-ReuseCache $case
    $prior=Get-Content (Join-Path $cache 'office-comparisons.json') -Raw | ConvertFrom-Json -AsHashtable
    foreach ($row in $prior.Rows) {
        $row.ReferenceKind='OfficeImage'
        $row.ReferenceImageMetadataSha256=(Get-FileHash (Join-Path $case.ReferenceDirectory ($row.Id+'-'+$row.Mode+'/reference-image-metadata.json'))).Hash
    }
    $prior | ConvertTo-Json -Depth 9 | Set-Content (Join-Path $cache 'office-image-comparisons.json')
    $case.Cache=$cache
    return $case
}
$case=New-ImageCase 'image-identity-reuse'
& $script -ReferenceKind OfficeImage -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $case.Cache | Out-Null
$report=Get-Content (Join-Path $case.CandidateDirectory 'office-image-comparisons.json') -Raw | ConvertFrom-Json
if ($report.ReferenceKind -ne 'OfficeImage' -or @($report.Rows | Where-Object {
    !$_.ReusedComparison -or !$_.ReferenceImageMetadataSha256 -or $_.Failures -notcontains 'ForegroundRecall' -or
    $_.ReferenceCalibrationStatus -ne 'Uncalibrated'
}).Count -or @(Get-ChildItem $case.Directory -Recurse -Filter '*.called').Count) {throw 'Image reuse lost provenance, failures or calibration status'}
$passed+='image-identity-reuse'
$case = New-ImageCase 'image-reference-extra-page' -ExtraReferencePage
$priorPath = Join-Path $case.Cache 'office-image-comparisons.json'
$prior = Get-Content -LiteralPath $priorPath -Raw | ConvertFrom-Json -AsHashtable
$metric = $prior.Rows[1].Metrics[0]
$metric.ForegroundRecall = 1
$prior.Rows[1].Metrics = @($metric, @{
    Page=2;ReferenceFile='page-002.png';CandidateFile=$null;DimensionsMatch=$false
    MeanAbsoluteError=$null;ChangedPixelRatioAtThreshold16=$null;ForegroundRecall=$null;StructuralSimilarity=$null
})
$imageArguments = @{ReferenceKind='OfficeImage';CandidateDirectory=$case.CandidateDirectory;ReferenceDirectory=$case.ReferenceDirectory
    VisualDiffPath=$visualDiff;ValidationToolsDirectory=$tools;ReuseComparisonDirectory=$case.Cache}
$prior | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $priorPath
& $script @imageArguments | Out-Null
$report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-image-comparisons.json') -Raw | ConvertFrom-Json
if (($report.Rows[1].Failures -join ',') -cne 'Pagination') { throw 'Complete reuse must retain the extra reference image failure' }
$passed += 'image-reference-extra-page'
$prior.Rows[1].Metrics = @($metric)
$prior | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $priorPath
$before = Snapshot $case.Directory
$rejected = $false
try { & $script @imageArguments | Out-Null }
catch {
    if ($_.Exception.Message -cne 'Comparison page metrics incomplete: second-final-preserve') { throw }
    $rejected = $true
}
if (!$rejected -or (Snapshot $case.Directory) -cne $before) { throw 'Truncated image metrics must refuse before publication' }
$passed += 'image-reference-truncated-pages'
foreach ($fault in @('image-as-pdf','changed-image-metadata')) {
    $case=New-ImageCase $fault
    $arguments=@{}
    if ($fault -eq 'changed-image-metadata') {
        $arguments.ReferenceKind='OfficeImage'
        $path=Join-Path $case.ReferenceDirectory 'second-final-preserve/reference-image-metadata.json'
        [IO.File]::AppendAllText($path,' ')
    }
    $before=Snapshot $case.Directory;$rejected=$false
    try {& $script @arguments -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null}
    catch {if ($_.Exception.Message -notmatch 'Reference kind differs|Image reference metadata differs') {throw};$rejected=$true}
    if (!$rejected -or (Snapshot $case.Directory) -cne $before) {throw 'Image reference mismatch mutated comparison stages'}
    $passed+=$fault
}
$outOfRangeMetrics = [ordered]@{
    'negative-error' = @('MeanAbsoluteError', -4)
    'error-above-byte-range' = @('MeanAbsoluteError', 256)
    'negative-changed-ratio' = @('ChangedPixelRatioAtThreshold16', -0.2)
    'changed-ratio-above-one' = @('ChangedPixelRatioAtThreshold16', 1.01)
    'negative-recall' = @('ForegroundRecall', -0.01)
    'recall-above-one' = @('ForegroundRecall', 1.5)
    'similarity-below-minus-one' = @('StructuralSimilarity', -1.01)
    'similarity-above-one' = @('StructuralSimilarity', 1.01)
}
foreach ($fault in (@('empty-metrics','missing-metric','null-error','nonfinite-metric','boolean-metric',
    'missing-page-index','duplicate-page','truncated-pages','truncated-reference-pages','wrong-file','nonboolean-dimensions','neither-file',
    'recall-contradicts-foreground') + @($outOfRangeMetrics.Keys))) {
    $case = New-Case $fault 'legacy' -ExtraReferencePage:($fault -eq 'truncated-reference-pages')
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
        'truncated-reference-pages' { $metric.ForegroundRecall=1 }
        'wrong-file' { $metric.ReferenceFile='page-002.png' }
        'nonboolean-dimensions' { $metric.DimensionsMatch='true' }
        'neither-file' { $metric.ReferenceFile=$null;$metric.CandidateFile=$null }
        'recall-contradicts-foreground' { $metric.ReferenceForegroundPixels=1; $metric.ForegroundRecall=$null }
        default { $field, $value = $outOfRangeMetrics[$fault]; $metric[$field] = $value }
    }
    $prior.Rows[1].Failures=@()
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    $before=Snapshot $case.Directory;$rejected=$false
    try {
        & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
            -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
    }
    catch {
        $expectedError = if ($outOfRangeMetrics.Contains($fault)) {
            'Comparison page metric outside its range: second-final-preserve / ' + $outOfRangeMetrics[$fault][0]
        } elseif ($fault -eq 'recall-contradicts-foreground') {
            'Comparison page recall disagrees with reference foreground pixels: second-final-preserve'
        } else { 'Comparison page metrics incomplete: second-final-preserve*' }
        if ($_.Exception.Message -notlike $expectedError) { throw }
        $rejected=$true
    }
    if (!$rejected -or (Snapshot $case.Directory) -ne $before) {
        throw ('Incomplete metrics must refuse before mutations: '+$fault)
    }
    $passed+=$fault
}
foreach ($control in @('cleared-summary','obsolete-failure','blank-reference','unverified-recall','zero-recall','dimension-mismatch',
    'near-dimensions-with-recall-failure','reference-extra-page','candidate-extra-page','legacy-missing-page-count','threshold-boundary',
    'metric-domain-minima','metric-domain-maxima')) {
    $case=New-Case $control 'legacy' -ExtraReferencePage:($control -eq 'reference-extra-page')
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
        'blank-reference' { $metric.ForegroundRecall=$null; $metric.ReferenceForegroundPixels=0 }
        'unverified-recall' { $metric.ForegroundRecall=$null; $expectedFailures=@('ForegroundRecallUnverified') }
        'zero-recall' { $metric.ForegroundRecall=0; $metric.ReferenceForegroundPixels=10; $expectedFailures=@('ForegroundRecall') }
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
        'legacy-missing-page-count' {
            $outcomesPath = Join-Path $case.CandidateDirectory 'outcomes.json'
            $outcomes = Get-Content -LiteralPath $outcomesPath -Raw | ConvertFrom-Json -AsHashtable
            $outcomes.Rows[1].Remove('InspectedPages')
            $outcomes | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $outcomesPath
            $expectedFailures=@('ForegroundRecall')
        }
        'threshold-boundary' {
            $metric.MeanAbsoluteError=.2
            $metric.ChangedPixelRatioAtThreshold16=.004
            $metric.ForegroundRecall=.99
            $metric.StructuralSimilarity=.985
        }
        'metric-domain-minima' {
            $metric.ForegroundRecall=0
            $metric.StructuralSimilarity=-1
            $expectedFailures=@('ForegroundRecall','StructuralSimilarity')
        }
        'metric-domain-maxima' {
            $metric.MeanAbsoluteError=255
            $metric.ChangedPixelRatioAtThreshold16=1
            $metric.ForegroundRecall=1
            $expectedFailures=@('ChangedPixels','MeanAbsoluteError')
        }
    }
    $prior | ConvertTo-Json -Depth 10 | Set-Content $path
    $before=Snapshot $case.ReferenceDirectory
    & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
        -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
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
            -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null
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
    -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
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
param($InputPdf, $OutputDirectory, $Dpi, [switch]$SkipBuild, $RasterizerPath)
if ($RasterizerPath -and !(Test-Path -LiteralPath $RasterizerPath -PathType Leaf)) {
    throw 'Mock rasterizer did not receive its bound assembly'
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$count=if ([IO.File]::ReadAllText($InputPdf).Contains('two pages')) {2} else {1}
for ($i=1;$i -le $count;$i++) { [IO.File]::WriteAllText((Join-Path $OutputDirectory ('page-{0:000}.png' -f $i)),('owned page '+$i)) }
$global:LASTEXITCODE=0
'@ | Set-Content -LiteralPath (Join-Path $tools 'RasterizePdf.ps1')
$comparisonProbeState=@{DiffCalls=0;AfterDiff=$null;BlankReference=$false}
function dotnet {
    if ($args[0] -eq $inspector) {
        Write-ControlPdfInspection $args[1]
        return
    }
    $comparisonProbeState.DiffCalls++
    $reference=$args[1];$candidate=$args[2];$diff=$args[3]
    New-Item -ItemType Directory -Path $diff | Out-Null
    $names=@(@((Get-ChildItem $reference -Filter 'page-*.png').Name)+@((Get-ChildItem $candidate -Filter 'page-*.png').Name) | Sort-Object -Unique)
    $metrics=@(foreach ($name in $names) {
        @{Page=[int]([IO.Path]::GetFileNameWithoutExtension($name).Substring(5));
          ReferenceFile=$(if (Test-Path (Join-Path $reference $name)) {$name} else {$null});
          CandidateFile=$(if (Test-Path (Join-Path $candidate $name)) {$name} else {$null});
          DimensionsMatch=$true;MeanAbsoluteError=0;ChangedPixelRatioAtThreshold16=0;StructuralSimilarity=1
          ForegroundRecall=$(if (!$comparisonProbeState.BlankReference) {1})
          ReferenceForegroundPixels=$(if ($comparisonProbeState.BlankReference) {0} else {10})}
    })
    ConvertTo-Json -InputObject $metrics -Depth 5 | Set-Content (Join-Path $diff 'metrics.json')
    if ($comparisonProbeState.AfterDiff) { & $comparisonProbeState.AfterDiff }
    $global:LASTEXITCODE=0
}
try {
    foreach ($control in @('matching-policy','legacy-policy','different-schema','changed-engine',
        'changed-wrapper','changed-rasterizer','different-dpi','missing-dpi','different-kind')) {
        $case=New-Case ('metric-policy-'+$control) 'legacy'
        $cache=New-ReuseCache $case
        $path=Join-Path $cache 'office-comparisons.json'
        $prior=Get-Content $path -Raw | ConvertFrom-Json -AsHashtable
        $engineBytes=[IO.File]::ReadAllBytes($visualDiff)
        $wrapper=Join-Path $tools 'RasterizePdf.ps1';$wrapperBytes=[IO.File]::ReadAllBytes($wrapper)
        $rasterizer=Join-Path $tools 'Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/net11.0/Lokad.OoxPdf.PdfiumRasterizer.dll'
        $expectedCalls=2;$expectedHistorical=1
        switch ($control) {
            'matching-policy' { $expectedCalls=0;$expectedHistorical=0 }
            'legacy-policy' { $prior.Remove('ComparisonPolicy') }
            'different-schema' { $prior.ComparisonPolicy.Schema=0 }
            'changed-engine' { [IO.File]::AppendAllText($visualDiff,'different engine bytes') }
            'changed-wrapper' { [IO.File]::AppendAllText($wrapper,"`n# different raster wrapper bytes") }
            'changed-rasterizer' {
                New-Item -ItemType Directory -Path (Split-Path $rasterizer) -Force | Out-Null
                [IO.File]::WriteAllText($rasterizer,'different rasterizer bytes; never executed')
            }
            'different-dpi' { $prior.Rows[0].RasterDpi=72;$expectedCalls=1;$expectedHistorical=0 }
            'missing-dpi' { $prior.Rows[0].Remove('RasterDpi');$expectedCalls=1;$expectedHistorical=0 }
            'different-kind' { $prior.Rows[0].ReferenceKind='OfficeImage';$expectedCalls=1;$expectedHistorical=0 }
        }
        $prior | ConvertTo-Json -Depth 10 | Set-Content $path
        $priorHash=(Get-FileHash $path).Hash;$beforeCalls=$comparisonProbeState.DiffCalls
        try {
            & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
                -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
            $report=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
            if ($comparisonProbeState.DiffCalls -ne $beforeCalls+$expectedCalls -or
                @($report.Rows | Where-Object ReusedComparison).Count -ne 2-$expectedCalls -or
                @($report.HistoricalComparisonReports).Count -ne $expectedHistorical -or
                $report.ComparisonPolicy.Schema -ne 1 -or
                $report.ComparisonPolicy.VisualDiffSha256 -ne (Get-FileHash $visualDiff).Hash -or
                @($report.Rows | Where-Object {$_.RasterDpi -ne 144 -or $_.ReferenceKind -ne 'OfficePdf'}).Count -or
                @($report.Rows | Where-Object {$_.ReusedComparison -and $_.Failures -notcontains 'ForegroundRecall'}).Count -or
                (Get-FileHash $path).Hash -ne $priorHash) { throw ('Incorrect metric-policy reuse: '+$control) }
            if ($expectedHistorical -and $report.HistoricalComparisonReports[0].Sha256 -ne $priorHash) {
                throw 'Historical metrics must retain their exact report identity'
            }
        }
        finally {
            [IO.File]::WriteAllBytes($visualDiff,$engineBytes)
            [IO.File]::WriteAllBytes($wrapper,$wrapperBytes)
            if ($control -eq 'changed-rasterizer') { Remove-Item -LiteralPath $rasterizer }
        }
        $passed+=('metric-policy-'+$control)
    }
    foreach ($control in @('engine','wrapper','native')) {
        $case=New-Case ('tool-change-during-'+$control) 'legacy'
        $binding=@{}
        $tool=switch ($control) {
            'engine' { $visualDiff }
            'wrapper' { Join-Path $tools 'RasterizePdf.ps1' }
            'native' {
                $alternate=Join-Path $case.Directory 'another checkout/tools'
                $runtime=Join-Path $alternate 'Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/net11.0'
                $native=Join-Path $alternate 'vendor/pdfium/win-x64/bin'
                New-Item -ItemType Directory -Path $runtime,$native -Force | Out-Null
                $binding.RasterizerPath=Join-Path $runtime 'Lokad.OoxPdf.PdfiumRasterizer.dll'
                [IO.File]::WriteAllText($binding.RasterizerPath, 'Owned fake rasterizer')
                $nativeFile=Join-Path $native 'pdfium.dll'
                [IO.File]::WriteAllText($nativeFile, 'Native dependency stand-in, never loaded')
                $nativeFile
            }
        }
        $original=[IO.File]::ReadAllBytes($tool);$rejected=$false
        $comparisonProbeState.AfterDiff={ [IO.File]::AppendAllText($tool,"`n# changed during comparison") }
        try {
            & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory `
                -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools @binding | Out-Null
        }
        catch { if ($_.Exception.Message -notlike 'Comparison tools changed during the batch*') {throw};$rejected=$true }
        finally { $comparisonProbeState.AfterDiff=$null;[IO.File]::WriteAllBytes($tool,$original) }
        if (!$rejected -or (Test-Path (Join-Path $case.CandidateDirectory 'office-comparisons.json'))) {
            throw 'A changed tool cannot publish qualified metrics'
        }
        $passed+=('tool-change-during-'+$control)
    }
    $baseline=New-Case 'complete-raster-baseline' 'legacy'
    $outcomePath = Join-Path $baseline.CandidateDirectory 'outcomes.json'
    $outcomes = Get-Content $outcomePath -Raw | ConvertFrom-Json -AsHashtable
    foreach ($row in $outcomes.Rows) {
        $row.MissingExpectedText = @()
        $row.IgnoredEquationDiagnostics = 0
        $row.ExpectedSourceEquations = 0
        $row.ExpectedSourcePictures = 0
        $row.Severity = 'Content blocker'
    }
    $outcomes.Rows[0].MissingExpectedText = @('Required phrase')
    $outcomes.Rows[0].Severity = 'Needs geometry and visual review'
    $outcomes | ConvertTo-Json -Depth 6 | Set-Content $outcomePath
    $comparisonProbeState.BlankReference=$true
    try {
        & $script -CandidateDirectory $baseline.CandidateDirectory -ReferenceDirectory $baseline.ReferenceDirectory -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null
    }
    finally { $comparisonProbeState.BlankReference=$false }
    $report=Get-Content (Join-Path $baseline.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    if ($report.Rows.Count -ne 2 -or @($report.Rows | Where-Object {
        $_.Failures.Count -or $_.ReusedComparison -or $_.Metrics[0].ReferenceForegroundPixels -ne 0 -or $null -ne $_.Metrics[0].ForegroundRecall
    }).Count) { throw 'Fresh zero-foreground metrics must pass with inapplicable recall' }
    if ($report.Rows[0].ContentSeverity -cne 'Content blocker' -or
        $report.Rows[1].ContentSeverity -cne 'Needs geometry and visual review') {
        throw 'Fresh comparison content severity must follow the observations'
    }
    $passed+='fresh-blank-reference'
    foreach ($id in @('first','second')) {
        # Preserve the completed attempt while testing derivative reuse separately
        # from the existing-comparison refusal. All paths belong to this fixture.
        $diff=[IO.Path]::GetFullPath((Join-Path $baseline.CandidateDirectory ($id+'-final-preserve/office-diff')))
        $ownedRoot=[IO.Path]::GetFullPath($OutputDirectory)+[IO.Path]::DirectorySeparatorChar
        if (!$diff.StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture move outside test directory' }
        [IO.Directory]::Move($diff,$diff+'.completed')
    }
    $directory=Join-Path $OutputDirectory 'changed-engine-reuses-pinned-rasters'
    Copy-Item -LiteralPath $baseline.Directory -Destination $directory -Recurse
    $candidate=Join-Path $directory 'candidate';$reference=Join-Path $directory 'reference'
    $cache=Join-Path $directory 'prior';New-Item -ItemType Directory -Path $cache | Out-Null
    Copy-Item -LiteralPath (Join-Path $candidate 'office-comparisons.json') -Destination $cache
    $rasters=@(foreach ($parent in @($candidate,$reference)) {
        foreach ($id in @('first','second')) {Join-Path $parent ($id+'-final-preserve/portfolio-raster')}
    })
    $beforeRasters=@($rasters | ForEach-Object {Snapshot $_}) -join "`n"
    $engineBytes=[IO.File]::ReadAllBytes($visualDiff);$beforeCalls=$comparisonProbeState.DiffCalls
    try {
        [IO.File]::AppendAllText($visualDiff,'new metric implementation, same rasterization')
        & $script -CandidateDirectory $candidate -ReferenceDirectory $reference -VisualDiffPath $visualDiff `
            -ValidationToolsDirectory $tools -ReuseComparisonDirectory $cache | Out-Null
        $report=Get-Content (Join-Path $candidate 'office-comparisons.json') -Raw | ConvertFrom-Json
        if ($comparisonProbeState.DiffCalls -ne $beforeCalls+2 -or @($report.Rows | Where-Object ReusedComparison).Count -or
            (@($rasters | ForEach-Object {Snapshot $_}) -join "`n") -cne $beforeRasters) {
            throw 'A metric-only tool change must recompare existing pinned rasters'
        }
    }
    finally {[IO.File]::WriteAllBytes($visualDiff,$engineBytes)}
    $passed+='changed-engine-reuses-pinned-rasters'
    foreach ($fault in @('valid-raster-reuse','stale-pdf','missing-page','changed-page','extra-page',
        'wrong-page-name','missing-middle-page','wrong-dpi','wrong-tool','missing-pin')) {
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
            { $_ -in @('wrong-page-name','missing-middle-page') } {
                $name=if ($fault -eq 'wrong-page-name') {'page-label.png'} else {'page-003.png'}
                $page=Join-Path $raster $name
                Set-Content -LiteralPath $page 'Correctly pinned invalid page sequence'
                $pin.Pages=@($pin.Pages)+@{Name=$name;Sha256=(Get-FileHash -LiteralPath $page).Hash}
                $pin | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $pinPath
            }
            'wrong-dpi' { $pin.Dpi=72;$pin | ConvertTo-Json -Depth 5 | Set-Content $pinPath }
            'wrong-tool' { $pin.RasterizerId='wrong';$pin | ConvertTo-Json -Depth 5 | Set-Content $pinPath }
            'missing-pin' { Remove-Item -LiteralPath $pinPath }
        }
        $before=Snapshot $directory;$beforeCalls=$comparisonProbeState.DiffCalls;$rejected=$false
        try { & $script -CandidateDirectory $candidate -ReferenceDirectory $reference -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null }
        catch {
            if ($fault -eq 'valid-raster-reuse' -or $_.Exception.Message -notmatch '[Pp]ortfolio raster') { throw }
            if ($fault -in @('wrong-page-name','missing-middle-page') -and
                $_.Exception.Message -notlike 'Portfolio raster page set is incomplete:*') { throw }
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
    & $script -CandidateDirectory $case.CandidateDirectory -ReferenceDirectory $case.ReferenceDirectory -VisualDiffPath $visualDiff -ValidationToolsDirectory $tools | Out-Null
    $result=Get-Content (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    if ($result.Rows[0].Failures -notcontains 'Pagination') { throw 'Stale legacy raster concealed the second page' }
    $passed+='legacy-raster-cannot-conceal-pagination'

    $case = New-Case 'audited-content' 'legacy'
    $measuredPath = Join-Path $case.CandidateDirectory 'outcomes.json'
    $measured = Get-Content -LiteralPath $measuredPath -Raw | ConvertFrom-Json -AsHashtable
    $measured.InspectorSha256 = 'D' * 64
    $auditRows = @(foreach ($row in $measured.Rows) {
        $key = $row.Id + '-' + $row.Mode
        $textPath = Join-Path $case.CandidateDirectory "$key/inspect/text-operations.json"
        New-Item -ItemType Directory -Path (Split-Path $textPath) | Out-Null
        '[]' | Set-Content -LiteralPath $textPath
        $row.TextOperationsSha256 = (Get-FileHash -LiteralPath $textPath).Hash
        $row.MissingExpectedText = @()
        $row.IgnoredEquationDiagnostics = 0
        $row.ExpectedSourceEquations = 1
        $row.ExpectedSourcePictures = 1
        @{Id=$row.Id;Mode=$row.Mode;InputSha256=$row.InputSha256;CandidatePdfSha256=$row.PdfSha256;
          CandidateOperationsSha256=$row.TextOperationsSha256;ExpectedSourceEquations=1;ExpectedSourcePictures=1;
          IgnoredEquationDiagnostics=0;UnresolvedSourceFeatures=@();ExpectedSimpleMath=@{'x+y'=1};
          MissingSimpleMath=if ($row.Id -eq 'first') {@{'x+y'=1}} else {@{}};
          PictureInk=@{ExpectedSourcePictures=1;Status=if ($row.Id -eq 'first') {'Picture witnesses present'} else {'Missing picture ink'}};
          ReferencePdfSha256=(Get-FileHash -LiteralPath (Join-Path $case.ReferenceDirectory "$key/reference.pdf")).Hash}
    })
    $measured | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $measuredPath
    $auditPath = Join-Path $OutputDirectory 'content-audit.json'
    $content = @{Schema='ooxpdf-portfolio-content-audit/1';ToolSha256=('E'*64);ManifestSha256=('F'*64);
        LibrarySha256=$measured.LibrarySha256;InspectorSha256=$measured.InspectorSha256;Rows=$auditRows}
    $content | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    $arguments = @{CandidateDirectory=$case.CandidateDirectory;ReferenceDirectory=$case.ReferenceDirectory;
        VisualDiffPath=$visualDiff;ValidationToolsDirectory=$tools;ContentAuditPath=$auditPath;
        ReuseComparisonDirectory=$baseline.CandidateDirectory}
    $beforeCalls = $comparisonProbeState.DiffCalls
    & $script @arguments | Out-Null
    $report = Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json
    if ($comparisonProbeState.DiffCalls -ne $beforeCalls -or @($report.Rows | Where-Object {
        !$_.ReusedComparison -or $_.Failures.Count -or $_.ContentSeverity -cne 'Content blocker'
    }).Count -or $report.Rows[0].ContentWitnesses.Equations -cne 'Missing' -or
        $report.Rows[1].ContentWitnesses.Pictures -cne 'Missing') {
        throw 'Reused passing metrics concealed a bound content omission'
    }
    $passed += 'reused-raster-retains-audited-content-blockers'
    $auditRows[0].UnresolvedSourceFeatures = @(@{Part='word/footer1.xml';Feature='fldSimple'})
    $auditRows[0].ReferenceConfirmedSimpleMath = @{'x+y'=1}
    $auditRows[0].MissingReferenceConfirmedSimpleMath = @{'x+y'=1}
    $content | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    & $script @arguments | Out-Null
    $mixed = (Get-Content -LiteralPath (Join-Path $case.CandidateDirectory 'office-comparisons.json') -Raw | ConvertFrom-Json).Rows[0]
    if ($comparisonProbeState.DiffCalls -ne $beforeCalls -or !$mixed.ReusedComparison -or $mixed.Failures.Count -or
        $mixed.ContentSeverity -cne 'Content blocker' -or $mixed.ContentWitnesses.Equations -cne 'Missing') {
        throw 'Unresolved field content concealed a reference-confirmed equation omission'
    }
    $passed += 'unresolved-field-retains-confirmed-equation-blocker'
    $auditRows[1].ReferencePdfSha256 = '0' * 64
    $content | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    $failure = $null
    try { & $script @arguments | Out-Null } catch { $failure = $_.Exception.Message }
    if ($failure -cne 'Portfolio content audit identity differs: second/final-preserve' -or
        $comparisonProbeState.DiffCalls -ne $beforeCalls) { throw 'Audit accepted an unrelated reference PDF' }
    $passed += 'content-audit-reference-refusal'
}
finally { Remove-Item Function:\dotnet }
@{Passed=$passed;LibraryBuilds=0;Conversions=0;NativeRasterizations=0;Scope='Fake PDF/page inputs; real preflight, reuse and comparison branches'} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
Write-Output ($passed.Count.ToString() + ' comparison orchestration checks passed; no build or native rendering.')
