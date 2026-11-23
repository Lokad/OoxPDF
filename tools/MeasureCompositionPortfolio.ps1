# Convert or reinspect a pinned portfolio with an existing immutable CLI; never builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CliPath,
    [Parameter(Mandatory=$true)] [string] $InspectorPath,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both',
    [string] $Manifest = (Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json'),
    [switch] $IncludeRawStreams,
    [string] $ReuseCandidateDirectory,
    [string] $ContentAuditPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'CompositionOutcomes.ps1')
$measurementScriptHash = (Get-FileHash -LiteralPath $PSCommandPath).Hash
$outcomeHelperHash = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'CompositionOutcomes.ps1')).Hash
$textAssemblyHash = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')).Hash
$repo = Split-Path -Parent $PSScriptRoot
$cli = (Resolve-Path -LiteralPath $CliPath).Path
$inspector = (Resolve-Path -LiteralPath $InspectorPath).Path
$library = Join-Path (Split-Path $cli) 'Lokad.OoxPdf.dll'
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
$cliHash = (Get-FileHash -LiteralPath $cli).Hash
$inspectorHash = (Get-FileHash -LiteralPath $inspector).Hash
$measurementIdentity = @{LibrarySha256=$libraryHash; InspectorSha256=$inspectorHash}
$loadedAudits = @{}
$audit = if ($ContentAuditPath) { Read-PortfolioContentAudit $ContentAuditPath $loadedAudits }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'A fresh portfolio output directory is required' }
# Resolve every input before any conversion or output publication. A late stale
# path must not invalidate an otherwise completed batch of expensive conversions.
$ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$manifestBytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Manifest).Path)
$documents = ConvertFrom-Json -AsHashtable -InputObject ([Text.Encoding]::UTF8.GetString($manifestBytes).TrimStart([char]0xFEFF))
if (!$documents) { throw 'An explicit nonempty manifest is required' }
$jobs = @(foreach ($document in @($documents)) {
    if (!$ids.Add($document.Id)) { throw ('Duplicate portfolio ID: ' + $document.Id) }
    $inputPath = (Resolve-Path -LiteralPath (Join-Path $repo $document.Input)).Path
    if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) { throw ('Portfolio input differs: ' + $document.Id) }
    [pscustomobject]@{Document=$document;InputPath=$inputPath}
})
$priorByKey = @{}
$priorDiagnosticsHashes = @{}
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
        foreach ($mode in @(Get-PortfolioModes $job.Document $DocxPreset)) {
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
            $priorDiagnosticsHashes[$key] = Get-VerifiedDiagnosticsHash -Path (Join-Path $source 'diagnostics.json') -Diagnostics $row.Diagnostics -Id $key
        }
    }
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$rows = @()
$conversions = 0
$reusedCandidates = 0
foreach ($job in $jobs) {
    $document = $job.Document
    $inputPath = $job.InputPath
    if ($ReuseCandidateDirectory -and (Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) {
        throw ('Portfolio input differs: ' + $document.Id)
    }
    $modes = @(Get-PortfolioModes $document $DocxPreset)
    foreach ($mode in $modes) {
        $expectations = Get-PortfolioContentExpectations $document $mode
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
            if ((Get-FileHash -LiteralPath $diagnosticsPath).Hash -ne $priorDiagnosticsHashes[$key] -or
                (Get-FileHash -LiteralPath (Join-Path $source 'diagnostics.json')).Hash -ne $priorDiagnosticsHashes[$key]) {
                throw ('Reused diagnostics changed during copying: ' + $key)
            }
            if ((Get-FileHash -LiteralPath $priorPath).Hash -ne $priorHash) { throw 'Reused outcome report changed' }
            $elapsedMs = $priorRow.ColdProcessElapsedMs
            $reusedCandidates++
        }
        else {
            $arguments = @('convert',$inputPath,$pdf,'--diagnostics',$diagnosticsPath)
            if ($mode -eq 'all-word-compatible') { $arguments += @('--docx-markup','all','--docx-markup-geometry','word-compatible') }
            # Each mode opens the source independently. Check it on both sides of that conversion.
            if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) {
                throw ('Portfolio input differs: ' + $document.Id + '/' + $mode)
            }
            $timer = [Diagnostics.Stopwatch]::StartNew()
            dotnet $cli @arguments *> (Join-Path $run 'convert.log')
            $exitCode = $LASTEXITCODE
            $timer.Stop()
            if ($exitCode -ne 0) { throw ('Portfolio conversion failed: ' + $document.Id + '/' + $mode) }
            if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) {
                throw ('Portfolio input changed during conversion: ' + $document.Id + '/' + $mode)
            }
            $elapsedMs = $timer.Elapsed.TotalMilliseconds
            $conversions++
        }
        $pdfHash = (Get-FileHash -LiteralPath $pdf).Hash
        if ($ReuseCandidateDirectory -and $pdfHash -ne $priorRow.PdfSha256) {
            throw ('Reused PDF differs: ' + $key)
        }
        $inspect = Join-Path $run 'inspect'
        # Keep text/graphics summaries for phrase checks and the independent content audit.
        $inspectArguments = @($pdf, $inspect)
        if (!$IncludeRawStreams) { $inspectArguments += '--text-only' }
        dotnet $inspector @inspectArguments *> (Join-Path $run 'inspect.log')
        if ($LASTEXITCODE -ne 0) { throw ('Portfolio inspection failed: ' + $document.Id) }
        $textReport = Join-Path $inspect 'text-operations.json'
        if (!(Test-Path -LiteralPath $textReport -PathType Leaf)) {
            throw ('Portfolio inspection did not publish text operations: ' + $document.Id + '/' + $mode)
        }
        $textInspection = Read-PortfolioTextOperations $textReport ($document.Id + '/' + $mode)
        if ((Get-FileHash -LiteralPath $pdf).Hash -ne $pdfHash) {
            throw ('Candidate PDF changed during inspection: ' + $document.Id + '/' + $mode)
        }
        # Page objects count image-only and blank pages even with an empty text report.
        $pageCount = @(Get-Content -LiteralPath (Join-Path $run 'inspect.log') | Where-Object { $_ -match '^\d+ \d+ obj: page$' }).Count
        if ($pageCount -lt 1) { throw ('No inspected PDF pages: ' + $document.Id) }
        $missing = @(Get-MissingPortfolioTextPhrases $textInspection.Operations $expectations.Text $mode)
        if ($ReuseCandidateDirectory) {
            if ((Get-FileHash -LiteralPath $diagnosticsPath).Hash -ne $priorDiagnosticsHashes[$key] -or
                (Get-FileHash -LiteralPath (Join-Path $source 'diagnostics.json')).Hash -ne $priorDiagnosticsHashes[$key] -or
                (Get-FileHash -LiteralPath $priorPath).Hash -ne $priorHash) {
                throw ('Reused diagnostics changed during inspection: ' + $key)
            }
            $diagnostics = @($priorRow.Diagnostics)
        }
        else { $diagnostics = @(Get-Content -LiteralPath $diagnosticsPath -Raw | ConvertFrom-Json) }
        $ignoredMath = Get-PortfolioIgnoredEquationCount $diagnostics
        $row = [ordered]@{
            Id=$document.Id;Mode=$mode;HeldOut=$document.HeldOut;InputSha256=$document.Sha256
            PdfSha256=$pdfHash;TextOperationsSha256=$textInspection.Hash;ColdProcessElapsedMs=$elapsedMs
            MissingExpectedText=$missing;IgnoredEquationDiagnostics=$ignoredMath
            ExpectedSourceEquations=$expectations.Equations
            ExpectedSourcePictures=$expectations.Pictures
            MathematicalTextOperations=@($textInspection.Operations | Where-Object Font -Like 'Eq*').Count
            InspectedPages=$pageCount
            Diagnostics=$diagnostics;VisualStatus='Pending independent Office comparison; no visual pass claimed'
        }
        if ($null -ne $expectations.FieldResults) { $row.ExpectedFieldResults = $expectations.FieldResults }
        $row.ContentWitnesses = Get-PortfolioContentWitnesses $row $measurementIdentity $audit $loadedAudits
        $row.Severity = Get-PortfolioContentSeverity $missing $ignoredMath $row.ContentWitnesses
        $rows += $row
        if ((Get-FileHash -LiteralPath $library).Hash -ne $libraryHash -or
            (Get-FileHash -LiteralPath $cli).Hash -ne $cliHash -or
            (Get-FileHash -LiteralPath $inspector).Hash -ne $inspectorHash) { throw 'A measured runtime changed' }
    }
}
[ordered]@{
    MeasurementScriptSha256=$measurementScriptHash;OutcomeHelperSha256=$outcomeHelperHash;TextAssemblyHelperSha256=$textAssemblyHash
    LibrarySha256=$libraryHash;CliSha256=$cliHash;InspectorSha256=$inspectorHash
    ManifestSha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($manifestBytes));Rebuilt=$false;DocxPreset=$DocxPreset
    InspectionMode=if ($IncludeRawStreams) {'SummariesAndRawStreams'} else {'SummariesOnly'}
    Conversions=$conversions;ReusedCandidates=$reusedCandidates
    ReusedOutcomes=$priorPath;ReusedOutcomesSha256=$priorHash
    MeasurementScope=if ($ReuseCandidateDirectory) {'Original cold CLI timings retained; reinspection performs no conversions'}
        else {'Cold CLI process wall time includes startup; allocation and stage baselines are separate'}
    Rows=$rows
} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $output 'outcomes.json') -Encoding utf8
Write-Output "$($rows.Count) pinned document/mode outcomes recorded; content and visual status remain separate."
