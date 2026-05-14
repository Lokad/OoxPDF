param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [switch] $ValidateOnly
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot "private-cases"

function Test-UnderDirectory([string] $Path, [string] $Directory) {
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-GitTracked([string] $Path) {
    $relative = [System.IO.Path]::GetRelativePath($repoRoot, $Path)
    git -C $repoRoot ls-files --error-unmatch -- $relative *> $null
    return $LASTEXITCODE -eq 0
}

function Assert-PrivateUntracked([string] $Path, [string] $Label) {
    if (-not (Test-UnderDirectory $Path $privateRoot)) {
        throw "$Label must be under $privateRoot."
    }

    if (Test-GitTracked $Path) {
        throw "$Label is tracked by git and must not be used as a private case: $Path"
    }
}

$caseFull = (Resolve-Path -LiteralPath $Case).Path
Assert-PrivateUntracked $caseFull "Private case manifest"

$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($manifest.id)) {
    throw "Private case manifest must contain an id."
}

$caseId = [string]$manifest.id
if ($caseId.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $caseId.Contains("/") -or $caseId.Contains("\")) {
    throw "Private case id must be a single filename-safe path segment."
}

if ([string]::IsNullOrWhiteSpace($manifest.input)) {
    throw "Private case manifest must contain an input path."
}

$inputPath = Join-Path $caseDirectory $manifest.input
$inputFull = (Resolve-Path -LiteralPath $inputPath).Path
Assert-PrivateUntracked $inputFull "Private case input"

if ($ValidateOnly) {
    Write-Host "Private case validation passed: $caseFull"
    Write-Host "Private input: $inputFull"
    return
}

$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$runRoot = Join-Path $repoRoot ("artifacts/private-visual/{0}/{1}" -f $caseId, $runId)
$referenceDir = Join-Path $runRoot "reference"
$candidateDir = Join-Path $runRoot "candidate"
$comparisonDir = Join-Path $runRoot "comparison"

New-Item -ItemType Directory -Force -Path $referenceDir, $candidateDir, $comparisonDir | Out-Null

$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
$dpi = if ($manifest.dpi -ne $null) { [int]$manifest.dpi } else { 144 }

dotnet build (Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj") --tl:off --nologo -v minimal
if ($LASTEXITCODE -ne 0) {
    throw "CLI build failed with exit code $LASTEXITCODE."
}

$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
dotnet $cliDll convert $inputFull $candidatePdf --diagnostics $diagnostics
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "RenderReference.ps1") -InputPath $inputFull -OutputDirectory $referenceDir -Dpi $dpi
& (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $candidatePdf -OutputDirectory $candidateDir -Dpi $dpi

dotnet build (Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj") --tl:off --nologo -v minimal
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff build failed with exit code $LASTEXITCODE."
}

$visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/net10.0/Lokad.OoxPdf.VisualDiff.dll"
dotnet $visualDiffDll $referenceDir $candidateDir $comparisonDir
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff failed with exit code $LASTEXITCODE."
}

$assessment = @"
# Private visual assessment: $($manifest.id)

Input: $inputFull
Kind: $($manifest.kind)
Run: $runId
Agent rating: <0-5>

## Summary

<One paragraph summary of visual fidelity. Keep this file private.>

## Pages/slides reviewed

- Page 1: <rating>, <main differences>

## Major defects

1. <Defect, page, suspected feature>

## Diagnostics reviewed

<Note important warnings/errors from diagnostics.json. Do not copy private text into public issues.>

## Public-safe follow-up

<An anonymized implementation target that can be copied into public notes.>
"@

Set-Content -LiteralPath (Join-Path $runRoot "assessment.md") -Value $assessment
Write-Host "Private visual case artifacts: $runRoot"
