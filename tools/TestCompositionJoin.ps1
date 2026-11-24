# Check ownership, source identities and complete preflight without a build/conversion.
[CmdletBinding()]
param([string] $OutputDirectory, [string] $JoinScript = (Join-Path $PSScriptRoot 'JoinCompositionPortfolio.ps1'))
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) { $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/composition-join-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root = (Resolve-Path -LiteralPath $OutputDirectory).Path
$manifest = Join-Path $root 'manifest.json'
$cases = @(@{Id='one';Input='one.docx';Sha256=('A'*64);ExpectedText=@('retained raw flag');ExpectedEquationCount=0},
           @{Id='two';Input='two.docx';Sha256=('B'*64);ExpectedText=@('retained raw flag');ExpectedEquationCount=0})
$cases |
    ConvertTo-Json | Set-Content -LiteralPath $manifest
$sources = @()
$rejections = @()
$sourceOnlyReports = [ordered]@{
    'inspect/font-resources.json' = '[{"PageNumber":87,"ResourceName":"F1"}]'
    'inspect/graphics-operations.json' = '[{"PageNumber":87,"Operator":"f"}]'
}
foreach ($id in @('one','two')) {
    $source = Join-Path $root $id
    $run = Join-Path $source ($id+'-all-word-compatible')
    foreach ($folder in @('inspect','raster','office-diff')) { New-Item -ItemType Directory -Path (Join-Path $run $folder) -Force | Out-Null }
    [IO.File]::WriteAllText((Join-Path $run 'candidate.pdf'), 'fake PDF '+$id)
    [IO.File]::WriteAllText((Join-Path $run 'inspect.log'), 'one page')
    [IO.File]::WriteAllText((Join-Path $run 'inspect/text-operations.json'), '[]')
    foreach ($name in $sourceOnlyReports.Keys) {
        [IO.File]::WriteAllText((Join-Path $run $name), $sourceOnlyReports[$name])
    }
    [IO.File]::WriteAllText((Join-Path $run 'diagnostics.json'), '[]')
    [IO.File]::WriteAllText((Join-Path $run 'raster/page-001.png'), 'raster remains at its original path')
    [IO.File]::WriteAllText((Join-Path $run 'office-diff/metrics.json'), 'metrics remain at their original path')
    [ordered]@{LibrarySha256=('C'*64);CliSha256=('D'*64);InspectorSha256=('E'*64);Rows=@(
        @{Id=$id;Mode='all-word-compatible';InputSha256=$(if ($id -eq 'one') {'A'*64} else {'B'*64});
          PdfSha256=(Get-FileHash (Join-Path $run 'candidate.pdf')).Hash;
          TextOperationsSha256=(Get-FileHash (Join-Path $run 'inspect/text-operations.json')).Hash;
          MissingExpectedText=@('retained raw flag');ExpectedSourceEquations=0;Diagnostics=@();IgnoredEquationDiagnostics=0;
          Severity='Needs geometry and visual review'}
    )} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $source 'outcomes.json')
    $sources += $source
}
function Join([string] $Name, [string[]] $Inputs=$sources, [string] $CaseManifest=$manifest) {
    & $JoinScript -SourceDirectory $Inputs -Manifest $CaseManifest -OutputDirectory (Join-Path $root $Name) -DocxPreset Review | Out-Null
}
function Reject([string] $Name, [string] $Expected, [string[]] $Inputs=$sources, [string] $CaseManifest=$manifest) {
    try { Join $Name $Inputs $CaseManifest }
    catch {
        if (!$_.Exception.Message.Contains($Expected)) { throw }
        if (Test-Path -LiteralPath (Join-Path $root $Name)) { throw 'Rejected join published a partial output' }
        $script:rejections += $Name
        return
    }
    throw ('Expected rejection: '+$Name)
}
function dotnet { throw 'A join must not build or convert' }
try {
    Join 'joined'
    $joined = Get-Content -LiteralPath (Join-Path $root 'joined/outcomes.json') -Raw | ConvertFrom-Json
    if ($joined.NewConversions -ne 0 -or $joined.Rows.Count -ne 2 -or $joined.SourceReports.Count -ne 2) { throw 'Joined outcomes differ' }
    foreach ($row in $joined.Rows) {
        $run = Join-Path $root ('joined/'+$row.Id+'-'+$row.Mode)
        if (Test-Path -LiteralPath (Join-Path $run 'raster')) { throw 'A raster tree was copied' }
        if (Test-Path -LiteralPath (Join-Path $run 'office-diff')) { throw 'A comparison tree was copied' }
        if (!$row.JoinedFrom.ReportSha256 -or $row.MissingExpectedText[0] -ne 'retained raw flag') { throw 'Provenance or original classification lost' }
        foreach ($name in $sourceOnlyReports.Keys) {
            if ((Test-Path -LiteralPath (Join-Path $run $name)) -or
                [IO.File]::ReadAllText((Join-Path $row.JoinedFrom.RunDirectory $name)) -cne $sourceOnlyReports[$name]) {
                throw 'Auxiliary inspection reports must remain with their source'
            }
        }
        if ($row.Severity -cne 'Content blocker') { throw 'Joined severity must follow the verified missing text' }
        if ($row.TextOperationsSha256 -cne (Get-FileHash (Join-Path $run 'inspect/text-operations.json')).Hash) {
            throw 'Joined text inspection is not pinned'
        }
    }
    $original = Join-Path $sources[0] 'one-all-word-compatible/candidate.pdf'
    $hash = (Get-FileHash -LiteralPath $original).Hash
    [IO.File]::WriteAllText((Join-Path $root 'joined/one-all-word-compatible/candidate.pdf'), 'changed owned copy')
    if ((Get-FileHash -LiteralPath $original).Hash -ne $hash) { throw 'Joined PDF aliases its source' }
    Reject 'missing' 'Measured outcome is missing' @($sources[0])
    $report = Join-Path $sources[1] 'outcomes.json'
    $bytes = [IO.File]::ReadAllBytes($report)
    $textReport = Join-Path $sources[1] 'two-all-word-compatible/inspect/text-operations.json'
    $textBytes = [IO.File]::ReadAllBytes($textReport)
    [IO.File]::WriteAllText($textReport, '[{"DecodedText":"Other PDF"}]')
    Reject 'stale-text-inspection' 'Candidate text inspection differs'
    [IO.File]::WriteAllBytes($textReport, $textBytes)
    $legacy = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json -AsHashtable
    $legacy.Rows[0].Remove('TextOperationsSha256')
    $legacy | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'legacy-text-inspection' 'reinspect the candidate PDF'
    [IO.File]::WriteAllBytes($report, $bytes)
    $sidecar = Join-Path $sources[1] 'two-all-word-compatible/diagnostics.json'
    [IO.File]::WriteAllText($sidecar, '[{"Id":"DOCX_UNSUPPORTED_EQUATION"}]')
    Reject 'stale-diagnostics' 'Reused diagnostics differ from the retained outcome'
    [IO.File]::WriteAllText($sidecar, ' [ ] ')
    $changedManifest = Join-Path $root 'changed-manifest.json'
    $caseOverride = $cases[1].Clone()
    $caseOverride.ExpectedByMode = @{'all-word-compatible'=@{ExpectedText=@('new required phrase')}}
    @($cases[0],$caseOverride) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $changedManifest
    Reject 'changed-text-expectations' 'Content expectations differ' $sources $changedManifest
    $caseOverride.ExpectedByMode = @{'all-word-compatible'=@{ExpectedEquationCount=7}}
    @($cases[0],$caseOverride) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $changedManifest
    Reject 'changed-equation-expectation' 'Content expectations differ' $sources $changedManifest
    $subsetManifest = Join-Path $root 'review-subset-manifest.json'
    $caseOverride.ExpectedText = @('different Final phrase')
    $caseOverride.ExpectedByMode = @{'all-word-compatible'=@{ExpectedText=@('retained raw flag')}}
    $caseOverride | ConvertTo-Json -AsArray -Depth 6 | Set-Content -LiteralPath $subsetManifest
    Join 'review-subset' $sources $subsetManifest
    $subset = Get-Content -LiteralPath (Join-Path $root 'review-subset/outcomes.json') -Raw | ConvertFrom-Json
    if ($subset.Rows.Count -ne 1 -or $subset.Rows[0].Id -cne 'two' -or $subset.ManifestSha256 -cne (Get-FileHash -LiteralPath $subsetManifest).Hash) {
        throw 'The requested Review expectation or subset identity differs'
    }
    $changed = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json
    $changed.LibrarySha256 = 'F'*64
    $changed | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'mixed-runtime' 'one frozen runtime'
    [IO.File]::WriteAllBytes($report, $bytes)
    $changed.Rows[0].InputSha256 = '0'*64
    $changed.LibrarySha256 = 'C'*64
    $changed | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'input-mismatch' 'Input identity differs'
    [IO.File]::WriteAllBytes($report, $bytes)
    $pdf = Join-Path $sources[1] 'two-all-word-compatible/candidate.pdf'
    $pdfBytes = [IO.File]::ReadAllBytes($pdf)
    [IO.File]::WriteAllText($pdf, 'bad bytes')
    Reject 'pdf-mismatch' 'Candidate identity differs'
    [IO.File]::WriteAllBytes($pdf, $pdfBytes)
    $badManifest = Join-Path $root 'bad-manifest.json'
    @(@{Id='../escape';Input='one.docx';Sha256=('A'*64)}) | ConvertTo-Json -AsArray | Set-Content -LiteralPath $badManifest
    Reject 'unsafe-id' 'Invalid or duplicate' $sources $badManifest
    $duplicate = Join-Path $root 'conflicting-source'
    Copy-Item -LiteralPath $sources[1] -Destination $duplicate -Recurse
    # Equal PDF hashes cannot resolve competing classifications or provenance.
    Reject 'identical-distinct-reports' 'Conflicting duplicate outcome' @($sources[0],$sources[1],$duplicate)
    $duplicateReport = Join-Path $duplicate 'outcomes.json'
    $duplicateData = Get-Content -LiteralPath $duplicateReport -Raw | ConvertFrom-Json -AsHashtable
    $duplicateData.Rows[0].Diagnostics = @(@{Id='DOCX_UNSUPPORTED_EQUATION'})
    $duplicateData.Rows[0].IgnoredEquationDiagnostics = 1
    $duplicateData.Rows[0].Severity = 'Content blocker'
    [IO.File]::WriteAllText((Join-Path $duplicate 'two-all-word-compatible/diagnostics.json'), '[{"Id":"DOCX_UNSUPPORTED_EQUATION"}]')
    $duplicateData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $duplicateReport
    Reject 'equal-pdf-clear-first' 'Conflicting duplicate outcome' @($sources[0],$sources[1],$duplicate)
    Reject 'equal-pdf-blocker-first' 'Conflicting duplicate outcome' @($sources[0],$duplicate,$sources[1])
    # Duplicate rows within one report are equally ambiguous.
    $sameReport = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json -AsHashtable
    $sameReport.Rows = @($sameReport.Rows) + @($duplicateData.Rows[0])
    $sameReport | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report
    Reject 'same-report-conflicting-rows' 'Conflicting duplicate outcome'
    [IO.File]::WriteAllBytes($report, $bytes)
    $duplicatePdf = Join-Path $duplicate 'two-all-word-compatible/candidate.pdf'
    [IO.File]::WriteAllText($duplicatePdf, 'different validly pinned fake PDF')
    $duplicateReport = Join-Path $duplicate 'outcomes.json'
    $duplicateData = Get-Content -LiteralPath $duplicateReport -Raw | ConvertFrom-Json
    $duplicateData.Rows[0].PdfSha256 = (Get-FileHash -LiteralPath $duplicatePdf).Hash
    $duplicateData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $duplicateReport
    Reject 'conflicting' 'Conflicting duplicate outcome' @($sources[0],$sources[1],$duplicate)
    # Repeating the same report selection is reused without repeating measurements.
    Join 'deduplicated' @($sources[0],$sources[0],$sources[1])
    $deduplicated = Get-Content -LiteralPath (Join-Path $root 'deduplicated/outcomes.json') -Raw | ConvertFrom-Json
    if ($deduplicated.Rows.Count -ne 2 -or $deduplicated.SourceReports.Count -ne 2) { throw 'Duplicate source selection was retained twice' }
    # Final and review outcomes remain separate, while presentation modes are
    # retained independently of the DOCX preset.
    foreach ($source in $sources) {
        $dataPath = Join-Path $source 'outcomes.json'
        $data = Get-Content -LiteralPath $dataPath -Raw | ConvertFrom-Json -AsHashtable
        $final = $data.Rows[0].Clone()
        $final.Mode = 'final-preserve'
        Copy-Item -LiteralPath (Join-Path $source ($final.Id+'-all-word-compatible')) `
            -Destination (Join-Path $source ($final.Id+'-final-preserve')) -Recurse
        $data.Rows = @($data.Rows) + @($final)
        $data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $dataPath
    }
    & $JoinScript -SourceDirectory $sources -Manifest $manifest -OutputDirectory (Join-Path $root 'both') | Out-Null
    $both = Get-Content -LiteralPath (Join-Path $root 'both/outcomes.json') -Raw | ConvertFrom-Json
    if ($both.Rows.Count -ne 4) { throw 'DOCX presets were conflated' }
    & $JoinScript -SourceDirectory $sources -Manifest $manifest -OutputDirectory (Join-Path $root 'final') -DocxPreset Final | Out-Null
    $finalOnly = Get-Content -LiteralPath (Join-Path $root 'final/outcomes.json') -Raw | ConvertFrom-Json
    if ($finalOnly.Rows.Count -ne 2 -or @($finalOnly.Rows | Where-Object Mode -NE 'final-preserve').Count) { throw 'Final selection differs' }
    $presentationManifest = Join-Path $root 'presentation-manifest.json'
    $presentationCase = $cases[0].Clone()
    $presentationCase.Input = 'one.pptx'
    $presentationCase | ConvertTo-Json -AsArray | Set-Content -LiteralPath $presentationManifest
    $presentationSource = Join-Path $root 'presentation-source'
    Copy-Item -LiteralPath $sources[0] -Destination $presentationSource -Recurse
    $presentationData = Get-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json') -Raw | ConvertFrom-Json -AsHashtable
    $presentationData.Rows = @($presentationData.Rows[0])
    $presentationData.Rows[0].Mode = 'presentation'
    Copy-Item -LiteralPath (Join-Path $presentationSource 'one-all-word-compatible') -Destination (Join-Path $presentationSource 'one-presentation') -Recurse
    $presentationData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json')
    Join 'presentation' @($presentationSource) $presentationManifest
    $presentation = Get-Content -LiteralPath (Join-Path $root 'presentation/outcomes.json') -Raw | ConvertFrom-Json
    if ($presentation.Rows.Count -ne 1 -or $presentation.Rows[0].Mode -ne 'presentation') { throw 'Presentation selection differs' }
    $docxData = Get-Content -LiteralPath (Join-Path $sources[0] 'outcomes.json') -Raw | ConvertFrom-Json -AsHashtable
    $docxData.Rows += $presentationData.Rows[0]
    $docxData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json')
    Reject 'pptx-mode-on-docx' 'Outcome mode does not match the input format' @($presentationSource,$sources[1])
    $docxData.Rows = @($presentationData.Rows[0],$docxData.Rows[1])
    $docxData | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $presentationSource 'outcomes.json')
    Reject 'unselected-docx-mode-on-pptx' 'Outcome mode does not match the input format' @($presentationSource) $presentationManifest

    $auditManifest = Join-Path $root 'audit-manifest.json'
    $auditCase = $cases[0].Clone()
    $auditCase.ExpectedText = @()
    $auditCase.ExpectedEquationCount = 1
    $auditCase.ExpectedInlineImageCount = 1
    ConvertTo-Json -InputObject @($auditCase) | Set-Content -LiteralPath $auditManifest
    $reportPath = Join-Path $sources[0] 'outcomes.json'
    $measured = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json -AsHashtable
    $row = $measured.Rows[0]
    $row.MissingExpectedText = @()
    $row.ExpectedSourceEquations = 1
    $row.ExpectedSourcePictures = 1
    $measured | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath
    $auditPath = Join-Path $root 'content-audit.json'
    @{Schema='ooxpdf-portfolio-content-audit/1';ToolSha256=('F'*64);ManifestSha256=(Get-FileHash $auditManifest).Hash;
      LibrarySha256=$measured.LibrarySha256;InspectorSha256=$measured.InspectorSha256;Rows=@(
        @{Id=$row.Id;Mode=$row.Mode;InputSha256=$row.InputSha256;CandidatePdfSha256=$row.PdfSha256;
          CandidateOperationsSha256=$row.TextOperationsSha256;ExpectedSourceEquations=1;ExpectedSourcePictures=1;
          IgnoredEquationDiagnostics=0;ReferencePdfSha256=('9'*64);UnresolvedSourceFeatures=@();
          ExpectedSimpleMath=@{'x+y'=1};MissingSimpleMath=@{};PictureInk=@{ExpectedSourcePictures=1;Status='Picture witnesses present'}}
    )} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    $arguments = @{SourceDirectory=$sources[0];Manifest=$auditManifest;OutputDirectory=(Join-Path $root 'audited');
        DocxPreset='Review';ContentAuditPath=$auditPath}
    & $JoinScript @arguments | Out-Null
    Join 'audited-again' @((Join-Path $root 'audited')) $auditManifest
    $retained = (Get-Content -LiteralPath (Join-Path $root 'audited-again/outcomes.json') -Raw | ConvertFrom-Json).Rows[0]
    if ($retained.Severity -cne 'Needs geometry and visual review' -or
        $retained.ContentWitnesses.Equations -cne 'Witnesses present' -or $retained.ContentWitnesses.Pictures -cne 'Witnesses present' -or
        $retained.ContentWitnesses.Audit.Sha256 -cne (Get-FileHash -LiteralPath $auditPath).Hash) {
        throw 'Joined audit did not retain the identified positive witnesses'
    }
    Add-Content -LiteralPath $auditPath ' '
    Reject 'changed-content-audit' 'Retained portfolio content audit changed' @((Join-Path $root 'audited')) $auditManifest
    $pending = Get-Content -LiteralPath $auditPath -Raw | ConvertFrom-Json -AsHashtable
    $pending.Rows[0].UnresolvedSourceFeatures = @(@{Part='word/document.xml';Feature='vanish'})
    $pending.Rows[0].MissingSimpleMath = @{'x+y'=1}
    $pending | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    $arguments.OutputDirectory = Join-Path $root 'audit-unverified'
    & $JoinScript @arguments | Out-Null
    $unverified = (Get-Content -LiteralPath (Join-Path $arguments.OutputDirectory 'outcomes.json') -Raw | ConvertFrom-Json).Rows[0]
    if ($unverified.ContentWitnesses.Equations -cne 'Unverified' -or $unverified.Severity -cne 'Content evidence incomplete') {
        throw 'The audit claimed a visible omission in unresolved source content'
    }
    $pending.Rows[0].ReferenceConfirmedSimpleMath = @{'x+y'=1}
    $pending.Rows[0].MissingReferenceConfirmedSimpleMath = @{'x+y'=1}
    $pending | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath
    $arguments.OutputDirectory = Join-Path $root 'audit-mixed'
    & $JoinScript @arguments | Out-Null
    $missing = (Get-Content -LiteralPath (Join-Path $arguments.OutputDirectory 'outcomes.json') -Raw | ConvertFrom-Json).Rows[0]
    if ($missing.ContentWitnesses.Equations -cne 'Missing' -or $missing.Severity -cne 'Content blocker') {
        throw 'The join suppressed a reference-confirmed omission beside unresolved source content'
    }
}
finally { Remove-Item Function:\dotnet }
[ordered]@{Rejections=$rejections;Builds=0;Conversions=0;RunnerInvocations=0} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'evidence.json')
Write-Output 'Composition outcome join ownership, identity, preflight and copy-avoidance checks passed.'
