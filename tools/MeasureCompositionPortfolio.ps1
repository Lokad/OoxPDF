# Convert the pinned portfolio with an existing immutable CLI; never builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CliPath,
    [Parameter(Mandatory=$true)] [string] $InspectorPath,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both',
    [string] $Manifest = (Join-Path $PSScriptRoot '../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json'),
    [switch] $IncludeRawStreams
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
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$rows = @()
function Normalize-Text([string] $Text) {
    [regex]::Replace($Text.Normalize([Text.NormalizationForm]::FormKC), '\s+', ' ').Trim()
}
foreach ($document in @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json)) {
    $inputPath = (Resolve-Path -LiteralPath (Join-Path $repo $document.Input)).Path
    if ((Get-FileHash -LiteralPath $inputPath).Hash -ne $document.Sha256) { throw ('Portfolio input differs: ' + $document.Id) }
    $isDocx = $inputPath.EndsWith('.docx', [StringComparison]::OrdinalIgnoreCase)
    $modes = if ($isDocx) { @('final-preserve','all-word-compatible') } else { @('presentation') }
    if ($isDocx -and $DocxPreset -ne 'Both') { $modes = if ($DocxPreset -eq 'Final') { @('final-preserve') } else { @('all-word-compatible') } }
    foreach ($mode in $modes) {
        $modeExpected = if ($document.ExpectedByMode) { $document.ExpectedByMode.PSObject.Properties[$mode].Value } else { $null }
        $expectedText = if ($modeExpected -and $modeExpected.PSObject.Properties['ExpectedText']) { @($modeExpected.ExpectedText) } else { @($document.ExpectedText) }
        $expectedEquations = if ($modeExpected -and $modeExpected.PSObject.Properties['ExpectedEquationCount']) { $modeExpected.ExpectedEquationCount } else { $document.ExpectedEquationCount }
        $run = Join-Path $output ($document.Id + '-' + $mode)
        New-Item -ItemType Directory -Path $run | Out-Null
        $pdf = Join-Path $run 'candidate.pdf'
        $diagnosticsPath = Join-Path $run 'diagnostics.json'
        $arguments = @('convert',$inputPath,$pdf,'--diagnostics',$diagnosticsPath)
        if ($mode -eq 'all-word-compatible') { $arguments += @('--docx-markup','all','--docx-markup-geometry','word-compatible') }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        dotnet $cli @arguments *> (Join-Path $run 'convert.log')
        $exitCode = $LASTEXITCODE
        $timer.Stop()
        if ($exitCode -ne 0) { throw ('Portfolio conversion failed: ' + $document.Id + '/' + $mode) }
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
            PdfSha256=(Get-FileHash -LiteralPath $pdf).Hash;ColdProcessElapsedMs=$timer.Elapsed.TotalMilliseconds
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
    MeasurementScope='Cold CLI process wall time includes startup; allocation and stage baselines are separate'
    Rows=$rows
} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $output 'outcomes.json') -Encoding utf8
Write-Output "$($rows.Count) pinned document/mode outcomes recorded; content and visual status remain separate."
