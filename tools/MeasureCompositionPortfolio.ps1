# Convert or reinspect a pinned portfolio with an existing immutable CLI; never builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CliPath,
    [Parameter(Mandatory=$true)] [string] $InspectorPath,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both',
    [string] $Manifest = (Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json'),
    [switch] $IncludeRawStreams,
    [string] $ReuseCandidateDirectory
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$cli = (Resolve-Path -LiteralPath $CliPath).Path
$inspector = (Resolve-Path -LiteralPath $InspectorPath).Path
$library = Join-Path (Split-Path $cli) 'Lokad.OoxPdf.dll'
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
$cliHash = (Get-FileHash -LiteralPath $cli).Hash
$inspectorHash = (Get-FileHash -LiteralPath $inspector).Hash
if (Test-Path -LiteralPath $OutputDirectory) { throw 'A fresh portfolio output directory is required' }
# Resolve every input before any conversion or output publication. A late stale
# path must not invalidate an otherwise completed batch of expensive conversions.
$jobs = @(foreach ($document in @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json)) {
    $inputPath = (Resolve-Path -LiteralPath (Join-Path $repo $document.Input)).Path
    if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) { throw ('Portfolio input differs: ' + $document.Id) }
    [pscustomobject]@{Document=$document;InputPath=$inputPath}
})
function Get-PortfolioModes($Document) {
    if (!$Document.Input.EndsWith('.docx', [StringComparison]::OrdinalIgnoreCase)) { return @('presentation') }
    if ($DocxPreset -eq 'Final') { return @('final-preserve') }
    if ($DocxPreset -eq 'Review') { return @('all-word-compatible') }
    return @('final-preserve','all-word-compatible')
}
$priorByKey = @{}
$priorPath = $null
$priorHash = $null
if ($ReuseCandidateDirectory) {
    $reuse = (Resolve-Path -LiteralPath $ReuseCandidateDirectory).Path
    $priorPath = Join-Path $reuse 'outcomes.json'
    $priorHash = (Get-FileHash -LiteralPath $priorPath).Hash
    $prior = Get-Content -LiteralPath $priorPath -Raw | ConvertFrom-Json
    if ($prior.LibrarySha256 -ne $libraryHash -or $prior.CliSha256 -ne $cliHash) {
        throw 'Reused portfolio production runtime differs'
    }
    foreach ($row in $prior.Rows) {
        $key = $row.Id + '-' + $row.Mode
        if ($priorByKey.ContainsKey($key)) { throw ('Duplicate reused outcome: ' + $key) }
        $priorByKey[$key] = $row
    }
    # Verify the complete selected batch before copying PDFs or inspecting any.
    foreach ($job in $jobs) {
        foreach ($mode in @(Get-PortfolioModes $job.Document)) {
            $key = $job.Document.Id + '-' + $mode
            $row = $priorByKey[$key]
            if (!$row -or $row.InputSha256 -ne $job.Document.Sha256) { throw ('Reused input differs: ' + $key) }
            $source = Join-Path $reuse $key
            foreach ($name in @('candidate.pdf','diagnostics.json','convert.log')) {
                if (!(Test-Path -LiteralPath (Join-Path $source $name) -PathType Leaf)) {
                    throw ('Reused candidate asset missing: ' + $key + '/' + $name)
                }
            }
            if ((Get-FileHash -LiteralPath (Join-Path $source 'candidate.pdf')).Hash -ne $row.PdfSha256) {
                throw ('Reused PDF differs: ' + $key)
            }
        }
    }
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$rows = @()
$conversions = 0
$reusedCandidates = 0
function Normalize-Text([string] $Text) {
    [regex]::Replace($Text.Normalize([Text.NormalizationForm]::FormKC), '\s+', ' ').Trim()
}
foreach ($job in $jobs) {
    $document = $job.Document
    $inputPath = $job.InputPath
    # Retain identity checks in case a source changes after preflight.
    if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) { throw ('Portfolio input differs: ' + $document.Id) }
    $modes = @(Get-PortfolioModes $document)
    foreach ($mode in $modes) {
        $modeExpected = if ($document.ExpectedByMode) { $document.ExpectedByMode.PSObject.Properties[$mode].Value } else { $null }
        $expectedText = if ($modeExpected -and $modeExpected.PSObject.Properties['ExpectedText']) { @($modeExpected.ExpectedText) } else { @($document.ExpectedText) }
        $expectedEquations = if ($modeExpected -and $modeExpected.PSObject.Properties['ExpectedEquationCount']) { $modeExpected.ExpectedEquationCount } else { $document.ExpectedEquationCount }
        $run = Join-Path $output ($document.Id + '-' + $mode)
        New-Item -ItemType Directory -Path $run | Out-Null
        $pdf = Join-Path $run 'candidate.pdf'
        $diagnosticsPath = Join-Path $run 'diagnostics.json'
        if ($ReuseCandidateDirectory) {
            $key = $document.Id + '-' + $mode
            $priorRow = $priorByKey[$key]
            $source = Join-Path $reuse $key
            foreach ($name in @('candidate.pdf','diagnostics.json','convert.log')) {
                Copy-Item -LiteralPath (Join-Path $source $name) -Destination (Join-Path $run $name)
            }
            if ((Get-FileHash -LiteralPath $pdf).Hash -ne $priorRow.PdfSha256 -or
                (Get-FileHash -LiteralPath $priorPath).Hash -ne $priorHash) { throw 'Reused candidate changed during inspection' }
            $elapsedMs = $priorRow.ColdProcessElapsedMs
            $reusedCandidates++
        }
        else {
            $arguments = @('convert',$inputPath,$pdf,'--diagnostics',$diagnosticsPath)
            if ($mode -eq 'all-word-compatible') { $arguments += @('--docx-markup','all','--docx-markup-geometry','word-compatible') }
            $timer = [Diagnostics.Stopwatch]::StartNew()
            dotnet $cli @arguments *> (Join-Path $run 'convert.log')
            $exitCode = $LASTEXITCODE
            $timer.Stop()
            if ($exitCode -ne 0) { throw ('Portfolio conversion failed: ' + $document.Id + '/' + $mode) }
            $elapsedMs = $timer.Elapsed.TotalMilliseconds
            $conversions++
        }
        $inspect = Join-Path $run 'inspect'
        # Text/graphics summaries suffice for routine content and geometry gates.
        # Dump decoded font/image streams only for a focused byte-level audit.
        $inspectArguments = @($pdf, $inspect)
        if (!$IncludeRawStreams) { $inspectArguments += '--text-only' }
        dotnet $inspector @inspectArguments *> (Join-Path $run 'inspect.log')
        if ($LASTEXITCODE -ne 0) { throw ('Portfolio inspection failed: ' + $document.Id) }
        # The inspector omits its text report when a PDF contains no text.
        # Page objects still identify image-only and blank pages independently.
        $textReport = Join-Path $inspect 'text-operations.json'
        $operations = if (Test-Path -LiteralPath $textReport) { @(Get-Content -LiteralPath $textReport -Raw | ConvertFrom-Json) } else { @() }
        $pageCount = @(Get-Content -LiteralPath (Join-Path $run 'inspect.log') | Where-Object { $_ -match '^\d+ \d+ obj: page$' }).Count
        if ($pageCount -lt 1) { throw ('No inspected PDF pages: ' + $document.Id) }
        $text = Normalize-Text (($operations.DecodedText) -join ' ')
        $missing = @($expectedText | Where-Object { !$text.Contains((Normalize-Text $_), [StringComparison]::Ordinal) })
        $diagnostics = @(Get-Content -LiteralPath $diagnosticsPath -Raw | ConvertFrom-Json)
        $ignoredMath = @($diagnostics | Where-Object { $_.Id -eq 'DOCX_UNSUPPORTED_EQUATION' }).Count
        $rows += [ordered]@{
            Id=$document.Id;Mode=$mode;HeldOut=$document.HeldOut;InputSha256=$document.Sha256
            PdfSha256=(Get-FileHash -LiteralPath $pdf).Hash;ColdProcessElapsedMs=$elapsedMs
            MissingExpectedText=$missing;IgnoredEquationDiagnostics=$ignoredMath
            ExpectedSourceEquations=$expectedEquations
            MathematicalTextOperations=@($operations | Where-Object Font -Like 'Eq*').Count
            InspectedPages=$pageCount
            Diagnostics=$diagnostics;VisualStatus='Pending independent Office comparison; no visual pass claimed'
            Severity=if ($missing.Count -or $ignoredMath) {'Content blocker'} else {'Needs geometry and visual review'}
        }
        if ((Get-FileHash -LiteralPath $library).Hash -ne $libraryHash -or
            (Get-FileHash -LiteralPath $cli).Hash -ne $cliHash -or
            (Get-FileHash -LiteralPath $inspector).Hash -ne $inspectorHash) { throw 'A measured runtime changed' }
    }
}
[ordered]@{
    LibrarySha256=$libraryHash;CliSha256=$cliHash;InspectorSha256=$inspectorHash
    ManifestSha256=(Get-FileHash -LiteralPath $Manifest).Hash;Rebuilt=$false;DocxPreset=$DocxPreset
    InspectionMode=if ($IncludeRawStreams) {'SummariesAndRawStreams'} else {'SummariesOnly'}
    Conversions=$conversions;ReusedCandidates=$reusedCandidates
    ReusedOutcomes=$priorPath;ReusedOutcomesSha256=$priorHash
    MeasurementScope=if ($ReuseCandidateDirectory) {'Original cold CLI timings retained; reinspection performs no conversions'}
        else {'Cold CLI process wall time includes startup; allocation and stage baselines are separate'}
    Rows=$rows
} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $output 'outcomes.json') -Encoding utf8
Write-Output "$($rows.Count) pinned document/mode outcomes recorded; content and visual status remain separate."
