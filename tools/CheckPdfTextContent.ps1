# Source-text preservation gate (P06 slice 1): converts each case document
# with the CLI, extracts candidate text with the independent PdfInspect
# decoder (own PDF parser, own ToUnicode maps; no Office needed), and checks
# hand-authored expected strings appear in order. Expectations live beside
# the inputs they describe, never generated from candidate output.
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [string] $CasesDir = "",
    [string] $Out = "",
    [string] $CliPath,
    [string] $InspectorPath,
    [switch] $SkipBuild
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($CasesDir)) {
    $CasesDir = Join-Path $repoRoot "tools/text-content-cases"
}
if ([string]::IsNullOrWhiteSpace($Out)) {
    $Out = Join-Path $repoRoot "artifacts/text-content"
}
$caseFiles = @(Get-ChildItem -LiteralPath $CasesDir -File -Filter *.json | Sort-Object Name)
if ($caseFiles.Count -eq 0) { throw "No text-content cases in $CasesDir." }
$cases = foreach ($caseFile in $caseFiles) {
    $case = Get-Content -Raw -LiteralPath $caseFile.FullName | ConvertFrom-Json -NoEnumerate
    if ($case -isnot [pscustomobject] -or $case.expected -isnot [array] -or $case.expected.Count -eq 0) {
        throw "Case '$($caseFile.Name)' requires a nonempty expected array of text strings."
    }
    foreach ($value in $case.expected) {
        if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) {
            throw "Case '$($caseFile.Name)' requires a nonempty expected array of text strings."
        }
    }
    if ($case.input -isnot [string] -or [string]::IsNullOrWhiteSpace($case.input)) {
        throw "Case '$($caseFile.Name)' requires an input path."
    }
    $inputPath = Join-Path $repoRoot $case.input
    if (!(Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "Missing input $($case.input)." }
    [pscustomobject]@{ Name = $caseFile.BaseName; InputPath = $inputPath; Expected = $case.expected }
}
$cliProject = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj"
$cliDll = if ($CliPath) { (Resolve-Path -LiteralPath $CliPath -ErrorAction Stop).Path } else { Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Release/$Framework/Lokad.OoxPdf.Cli.dll" }
$inspectorDll = if ($InspectorPath) { (Resolve-Path -LiteralPath $InspectorPath -ErrorAction Stop).Path } else { Join-Path $repoRoot "tools/Lokad.OoxPdf.PdfInspect/bin/Debug/$Framework/Lokad.OoxPdf.PdfInspect.dll" }
foreach ($binding in @(@{Path=$cliDll;Build=(!$CliPath -and !$SkipBuild)}, @{Path=$inspectorDll;Build=(!$InspectorPath -and !$SkipBuild)})) {
    if (!$binding.Build -and !(Test-Path -LiteralPath $binding.Path -PathType Leaf)) { throw ('Missing prebuilt text-gate binary: '+$binding.Path) }
}
$buildInfoDir = Join-Path $Out "build-info"
New-Item -ItemType Directory -Force -Path $buildInfoDir | Out-Null
if (!$CliPath -and !$SkipBuild) {
    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Framework $Framework -Project $cliProject -OutputDll $cliDll -Description "CLI" -Configuration Release -RecordPath (Join-Path $buildInfoDir "cli.json")
}
if (!$InspectorPath -and !$SkipBuild) {
    & (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Framework $Framework -Project (Join-Path $repoRoot 'tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj') -OutputDll $inspectorDll -Description 'PDF inspect'
}
[ordered]@{CliPath=$cliDll;CliSha256=(Get-FileHash -LiteralPath $cliDll).Hash;
    LibrarySha256=(Get-FileHash -LiteralPath (Join-Path (Split-Path $cliDll) 'Lokad.OoxPdf.dll')).Hash;
    InspectorPath=$inspectorDll;InspectorSha256=(Get-FileHash -LiteralPath $inspectorDll).Hash;
    CliBuildRequested=(!$CliPath -and !$SkipBuild);InspectorBuildRequested=(!$InspectorPath -and !$SkipBuild)} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $buildInfoDir 'runtime.json') -Encoding utf8

$failures = 0
foreach ($case in $cases) {
    $caseDir = Join-Path $Out $case.Name
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null
    $pdf = Join-Path $caseDir "output.pdf"
    $inspectDir = Join-Path $caseDir "inspect"
    Write-Host "[$($case.Name)] converting..."
    & dotnet $cliDll convert $case.InputPath $pdf
    if ($LASTEXITCODE -ne 0) { throw "Conversion failed for $($case.InputPath)." }
    & (Join-Path $repoRoot "tools/InspectPdf.ps1") -Framework $Framework -InputPdf $pdf -OutputDirectory $inspectDir -TextOnly -InspectorPath $inspectorDll -SkipBuild
    $ops = Get-Content -Raw -LiteralPath (Join-Path $inspectDir "text-operations.json") | ConvertFrom-Json
    $fullText = Join-DecodedPdfTextPages $ops
    $cursor = 0
    $missing = 0
    foreach ($expected in $case.Expected) {
        $found = $fullText.IndexOf($expected, $cursor, [StringComparison]::Ordinal)
        if ($found -lt 0) {
            $missing++
            Write-Host "[$($case.Name)] MISSING: $expected"
        }
        else { $cursor = $found + $expected.Length }
    }
    $failures += $missing
    if ($missing -eq 0) { Write-Host "[$($case.Name)] ok" }
}
if ($failures -ne 0) { throw "$failures expected string(s) missing from extracted text." }
Write-Host "Text-content gate passed for $($cases.Count) case(s)."
