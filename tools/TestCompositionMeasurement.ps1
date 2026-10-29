# Prove a late invalid input prevents all conversion work, without running .NET.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $MeasurementScript = (Join-Path $PSScriptRoot 'MeasureCompositionPortfolio.ps1')
)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/measurement-test-' + [Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
$tools = Join-Path $repo 'tools'
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$script = Join-Path $tools 'MeasureCompositionPortfolio.ps1'
Copy-Item -LiteralPath $MeasurementScript -Destination $script
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1') -Destination (Join-Path $tools 'PdfTextGeometry.ps1')
$cli = Join-Path $repo 'candidate.dll'
$inspector = Join-Path $repo 'inspector.dll'
foreach ($file in @($cli, $inspector, (Join-Path $repo 'Lokad.OoxPdf.dll'))) {
    'Sentinel runtime; never execute' | Set-Content -LiteralPath $file
}
$calls = [Collections.Generic.List[object]]::new()
$inspectionOnly = $false
$inspectionOperations = @([pscustomobject]@{PageNumber=1;Font='Body';DecodedText='Retained source.'})
$mutation = [pscustomobject]@{CopyPath=$null;InspectionSource=$null;InspectionDestination=$false}
function Copy-Item {
    param([string] $LiteralPath, [string] $Destination)
    if ($mutation.CopyPath -and [IO.Path]::GetFullPath($LiteralPath) -eq [IO.Path]::GetFullPath($mutation.CopyPath)) {
        [IO.File]::WriteAllText($LiteralPath, '[]')
        $mutation.CopyPath = $null
    }
    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination
}
function dotnet {
    $calls.Add($args)
    if ($inspectionOnly) {
        if ($args[0] -ne $inspector) { throw 'Reinspection reached the converter' }
        $changed = '[{"Id":"DOCX_UNSUPPORTED_EQUATION","Severity":1,"Message":"Late change."}]'
        if ($mutation.InspectionSource) {
            [IO.File]::WriteAllText($mutation.InspectionSource, $changed)
            $mutation.InspectionSource = $null
        }
        if ($mutation.InspectionDestination) {
            [IO.File]::WriteAllText((Join-Path (Split-Path $args[2]) 'diagnostics.json'), $changed)
            $mutation.InspectionDestination = $false
        }
        New-Item -ItemType Directory -Path $args[2] | Out-Null
        $inspectionOperations | ConvertTo-Json -AsArray |
            Set-Content -LiteralPath (Join-Path $args[2] 'text-operations.json')
        $global:LASTEXITCODE = 0
        foreach ($page in @($inspectionOperations.PageNumber | Sort-Object -Unique)) {
            Write-Output ($page.ToString() + ' 0 obj: page')
        }
        return
    }
    throw 'Sentinel conversion reached'
}
$first = Join-Path $repo 'first.docx'
$second = Join-Path $repo 'second.docx'
'First input' | Set-Content -LiteralPath $first
$firstHash = (Get-FileHash -LiteralPath $first).Hash
foreach ($fault in @('missing', 'hash', 'valid')) {
    $case = Join-Path $OutputDirectory $fault
    New-Item -ItemType Directory -Path $case | Out-Null
    $manifest = Join-Path $case 'manifest.json'
    if ($fault -ne 'missing') { 'Second input' | Set-Content -LiteralPath $second }
    $secondHash = if ($fault -eq 'valid') { (Get-FileHash -LiteralPath $second).Hash } else { '0' * 64 }
    @(@{Id='first';Input='first.docx';Sha256=$firstHash},
      @{Id='second';Input='second.docx';Sha256=$secondHash}) |
        ConvertTo-Json | Set-Content -LiteralPath $manifest
    $output = Join-Path $case 'candidate'
    $calls.Clear()
    $failure = $null
    try {
        & $script -CliPath $cli -InspectorPath $inspector -Manifest $manifest -OutputDirectory $output
    }
    catch { $failure = $_ }
    if ($fault -eq 'valid') {
        if ($calls.Count -ne 1 -or "$failure" -notlike '*Sentinel conversion reached*') {
            throw 'Valid inputs must reach the converter after preflight'
        }
    }
    else {
        if (!$failure -or $calls.Count -ne 0 -or (Test-Path -LiteralPath $output)) {
            throw ('Late invalid input started conversion or published output: ' + $fault)
        }
        if ($fault -eq 'hash' -and "$failure" -notlike '*Portfolio input differs: second*') {
            throw 'Hash mismatch must identify its actual input'
        }
    }
}
$inspectionOnly = $true
$preflightRejects = $lateRejects = $validReuses = 0
foreach ($fault in @('runtime', 'input', 'pdf', 'missing', 'duplicate',
    'diagnostics-cleared', 'diagnostics-replaced', 'diagnostics-added', 'diagnostics-malformed',
    'diagnostics-null', 'diagnostics-object', 'missing-provenance', 'null-provenance', 'scalar-provenance',
    'late-copy', 'late-destination', 'late-source', 'late-later-source', 'equivalent', 'valid')) {
    $case = Join-Path $OutputDirectory ('reuse-' + $fault)
    $prior = Join-Path $case 'prior'
    New-Item -ItemType Directory -Path $prior -Force | Out-Null
    $manifest = Join-Path $case 'manifest.json'
    $documents = @(@{Id='first';Input='first.docx';Sha256=$firstHash;ExpectedText=@('Retained source.');ExpectedEquationCount=0},
        @{Id='second';Input='second.docx';Sha256=(Get-FileHash -LiteralPath $second).Hash;ExpectedText=@('Retained source.');ExpectedEquationCount=0})
    $documents | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifest
    $priorRows = @(foreach ($document in $documents) {
        foreach ($mode in @('final-preserve','all-word-compatible')) {
            $key = $document.Id + '-' + $mode
            $run = Join-Path $prior $key
            New-Item -ItemType Directory -Path $run | Out-Null
            ('Pinned PDF ' + $key) | Set-Content -LiteralPath (Join-Path $run 'candidate.pdf')
            $diagnostics = @()
            if ($document.Id -ne 'first') {
                $diagnostics = @([ordered]@{Id='DOCX_APPROXIMATED_EQUATION';Severity=1;Message='Retained warning.'})
            }
            ConvertTo-Json -InputObject $diagnostics -Depth 4 | Set-Content -LiteralPath (Join-Path $run 'diagnostics.json')
            if ($fault -ne 'missing' -or $key -ne 'second-all-word-compatible') {
                'Retained original conversion log' | Set-Content -LiteralPath (Join-Path $run 'convert.log')
            }
            @{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;
                PdfSha256=(Get-FileHash -LiteralPath (Join-Path $run 'candidate.pdf')).Hash;ColdProcessElapsedMs=123d;Diagnostics=$diagnostics}
        }
    })
    if ($fault -eq 'input') { $priorRows[-1].InputSha256 = '0' * 64 }
    if ($fault -eq 'pdf') { $priorRows[-1].PdfSha256 = '0' * 64 }
    if ($fault -eq 'duplicate') { $priorRows += $priorRows[0] }
    $lastDiagnostics = Join-Path $prior 'second-all-word-compatible/diagnostics.json'
    switch ($fault) {
        'diagnostics-cleared' { [IO.File]::WriteAllText($lastDiagnostics, '[]') }
        'diagnostics-replaced' { [IO.File]::WriteAllText($lastDiagnostics, '[{"Id":"DOCX_APPROXIMATED_EQUATION","Severity":1,"Message":"Different warning."}]') }
        'diagnostics-added' { [IO.File]::WriteAllText($lastDiagnostics, '[{"Id":"DOCX_APPROXIMATED_EQUATION","Severity":1,"Message":"Retained warning."},{"Id":"DOCX_UNSUPPORTED_EQUATION","Severity":1,"Message":"Stale warning."}]') }
        'diagnostics-malformed' { [IO.File]::WriteAllText($lastDiagnostics, '[{') }
        'diagnostics-null' { [IO.File]::WriteAllText($lastDiagnostics, 'null') }
        'diagnostics-object' { [IO.File]::WriteAllText($lastDiagnostics, '{"Id":"DOCX_APPROXIMATED_EQUATION","Severity":1,"Message":"Retained warning."}') }
        'missing-provenance' { $priorRows[-1].Remove('Diagnostics') }
        'null-provenance' { $priorRows[-1].Diagnostics = $null }
        'scalar-provenance' { $priorRows[-1].Diagnostics = $priorRows[-1].Diagnostics[0] }
        'equivalent' {
            # Whitespace, a UTF-8 BOM and property ordering do not change diagnostics.
            [IO.File]::WriteAllText($lastDiagnostics, '[ { "Message" : "Retained warning.", "Severity" : 1, "Id" : "DOCX_APPROXIMATED_EQUATION" } ]', [Text.UTF8Encoding]::new($true))
        }
    }
    @{LibrarySha256=$(if ($fault -eq 'runtime') {'0' * 64} else {(Get-FileHash -LiteralPath (Join-Path $repo 'Lokad.OoxPdf.dll')).Hash});
        CliSha256=(Get-FileHash -LiteralPath $cli).Hash;Rows=$priorRows} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $prior 'outcomes.json')
    $priorHash = (Get-FileHash -LiteralPath (Join-Path $prior 'outcomes.json')).Hash
    $calls.Clear()
    $mutation.CopyPath = $mutation.InspectionSource = $null
    $mutation.InspectionDestination = $false
    switch ($fault) {
        'late-copy' { $mutation.CopyPath = $lastDiagnostics }
        'late-destination' { $mutation.InspectionDestination = $true }
        'late-source' { $mutation.InspectionSource = Join-Path $prior 'first-final-preserve/diagnostics.json' }
        'late-later-source' { $mutation.InspectionSource = $lastDiagnostics }
    }
    $output = Join-Path $case 'candidate'
    $failure = $null
    try { & $script -CliPath $cli -InspectorPath $inspector -Manifest $manifest -OutputDirectory $output -ReuseCandidateDirectory $prior | Out-Null }
    catch { $failure = $_ }
    if ($fault.StartsWith('late-')) {
        $expectedCalls = if ($fault -in @('late-copy','late-later-source')) { 3 } else { 1 }
        if (!$failure -or "$failure" -notlike '*Reused diagnostics changed*' -or $calls.Count -ne $expectedCalls -or
            (Test-Path -LiteralPath (Join-Path $output 'outcomes.json'))) {
            throw ('Late diagnostic mutation was not refused before a completed report: ' + $fault + '; calls=' + $calls.Count + '; failure=' + $failure)
        }
        $lateRejects++
        continue
    }
    if ($fault -notin @('valid','equivalent')) {
        if (!$failure -or $calls.Count -ne 0 -or (Test-Path -LiteralPath $output)) {
            throw ('Invalid reuse started work or published output: ' + $fault)
        }
        $preflightRejects++
        continue
    }
    if ($failure) { throw $failure }
    $report = Get-Content -LiteralPath (Join-Path $output 'outcomes.json') -Raw | ConvertFrom-Json
    if ($calls.Count -ne 4 -or $report.Conversions -ne 0 -or $report.ReusedCandidates -ne 4 -or
        $report.ReusedOutcomesSha256 -ne $priorHash -or
        (Get-FileHash -LiteralPath (Join-Path $prior 'outcomes.json')).Hash -ne $priorHash) {
        throw 'Successful reinspection must preserve all candidates and original provenance without conversion'
    }
    foreach ($row in $report.Rows) {
        if ($row.ColdProcessElapsedMs -ne 123d -or $row.MissingExpectedText.Count -ne 0 -or
            $row.PdfSha256 -ne (Get-FileHash -LiteralPath (Join-Path $prior ($row.Id + '-' + $row.Mode + '/candidate.pdf'))).Hash) {
            throw 'Reinspection changed candidate identity/timing or lost content'
        }
        $expectedDiagnostics = if ($row.Id -eq 'first') { 0 } else { 1 }
        if ($row.Diagnostics.Count -ne $expectedDiagnostics -or $row.IgnoredEquationDiagnostics -ne 0 -or
            $row.Severity -ne 'Needs geometry and visual review' -or
            ($expectedDiagnostics -and $row.Diagnostics[0].Message -ne 'Retained warning.')) {
            throw 'Reinspection changed retained diagnostics or content classification'
        }
    }
    $validReuses++
}
$textCases = @(
    @{Id='punctuation';Pages=@(1,1,1);Text=@('Public first','-','page heading.');Expected='Public first-page heading.';Missing=$false},
    @{Id='word-fragments';Pages=@(1,1,1);Text=@('Ret','ained ','source.');Expected='Retained source.';Missing=$false},
    @{Id='emitted-word-space';Pages=@(1,1,1);Text=@('Retained',' ','source.');Expected='Retained source.';Missing=$false},
    @{Id='absent-word-space';Pages=@(1,1);Text=@('Retained','source.');Expected='Retained source.';Missing=$true},
    @{Id='page-boundary';Pages=@(1,2);Text=@('Retained','source.');Expected='Retainedsource.';Missing=$true},
    @{Id='numeric-page-order';Pages=@(10,2);Text=@('last','first');Expected='first last';Missing=$false},
    @{Id='decoded-unicode';Pages=@(1,1,1);Text=@('of',([string][char]0xFB01),'ce');Expected='office';Missing=$false},
    @{Id='missing-word';Pages=@(1,1);Text=@('Retained ','.');Expected='Retained source.';Missing=$true}
)
foreach ($textCase in $textCases) {
    $inspectionOperations = @(for ($index=0; $index -lt $textCase.Text.Count; $index++) {
        [pscustomobject]@{PageNumber=$textCase.Pages[$index];Font='Body';DecodedText=$textCase.Text[$index]}
    })
    $case = Join-Path $OutputDirectory ('text-' + $textCase.Id)
    New-Item -ItemType Directory -Path $case | Out-Null
    $manifest = Join-Path $case 'manifest.json'
    @(@{Id='first';Input='first.docx';Sha256=$firstHash;ExpectedText=@($textCase.Expected);ExpectedEquationCount=0}) |
        ConvertTo-Json -Depth 4 -AsArray | Set-Content -LiteralPath $manifest
    $calls.Clear()
    $output = Join-Path $case 'candidate'
    & $script -CliPath $cli -InspectorPath $inspector -Manifest $manifest -OutputDirectory $output -ReuseCandidateDirectory $prior | Out-Null
    $report = Get-Content -LiteralPath (Join-Path $output 'outcomes.json') -Raw | ConvertFrom-Json
    if ($calls.Count -ne 2 -or $report.Conversions -ne 0 -or $report.ReusedCandidates -ne 2) {
        throw ('Text-screen reinspection started conversion or lost candidates: ' + $textCase.Id)
    }
    foreach ($row in $report.Rows) {
        $expectedMissing = if ($textCase.Missing) {1} else {0}
        if ($row.MissingExpectedText.Count -ne $expectedMissing -or
            ($textCase.Missing -and $row.MissingExpectedText[0] -cne $textCase.Expected)) {
            throw ('Decoded operation assembly changed a source expectation: ' + $textCase.Id)
        }
        if (($row.Severity -eq 'Content blocker') -ne $textCase.Missing) {
            throw ('Text-screen classification differs: ' + $textCase.Id)
        }
    }
}
Write-Output "Measurement/reinspection passes: $preflightRejects invalid caches reject before output/work, $lateRejects late changes reject before report publication, $validReuses valid diagnostic batches and $($textCases.Count) text-assembly controls reuse without conversion. No .NET processes or library builds."
