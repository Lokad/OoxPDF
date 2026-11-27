# Shared private-mode path guards (T04): Test-UnderDirectory, Test-GitTracked,
# and Assert-PrivateUntracked keep untracked private inputs out of versioned
# trees. Callers must define $repoRoot and $privateRoot before dot-sourcing
# (both resolve in caller scope); every consumer sets them identically.
# Dot-source this file instead of copying the functions into private tools.

function Test-UnderDirectory([string] $Path, [string] $Directory) {
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-GitTracked([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $relative = $fullPath.Substring($fullRoot.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        git -C $repoRoot ls-files --error-unmatch -- $relative *> $null
        return $LASTEXITCODE -eq 0
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Assert-PrivateUntracked([string] $Path, [string] $Label) {
    if (-not (Test-UnderDirectory $Path $privateRoot)) {
        throw "$Label must be under $privateRoot."
    }

    if (Test-GitTracked $Path) {
        throw "$Label is tracked by git and must not be used as a private case: $Path"
    }
}

function Read-PrivateCase([string] $Case) {
    $casePath = (Resolve-Path -LiteralPath $Case).Path
    Assert-PrivateUntracked $casePath 'Private case manifest'

    $manifest = Get-Content -Raw -LiteralPath $casePath | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace($manifest.id)) {
        throw 'Private case manifest must contain an id.'
    }
    $id = [string]$manifest.id
    if ($id.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $id.Contains('/') -or $id.Contains('\')) {
        throw 'Private case id must be a single filename-safe path segment.'
    }
    if ([string]::IsNullOrWhiteSpace($manifest.input)) {
        throw 'Private case manifest must contain an input path.'
    }
    $inputPath = (Resolve-Path -LiteralPath (Join-Path (Split-Path -Parent $casePath) $manifest.input)).Path
    Assert-PrivateUntracked $inputPath 'Private case input'
    return [pscustomobject]@{ CasePath = $casePath; Id = $id; InputPath = $inputPath; Manifest = $manifest }
}

function Resolve-PrivateCaseCli([string] $CliPath, [string] $Framework) {
    if ($CliPath) {
        $assembly = Get-Item -LiteralPath $CliPath
        if ($assembly -isnot [IO.FileInfo] -or $assembly.Extension -ine '.dll') {
            throw 'CliPath must identify an existing CLI assembly.'
        }
        return $assembly.FullName
    }

    $project = Join-Path $repoRoot 'src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj'
    $assemblyPath = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/$Framework/Lokad.OoxPdf.Cli.dll"
    & (Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1') -Project $project -OutputDll $assemblyPath -Description 'CLI' -Framework $Framework | Out-Host
    return $assemblyPath
}
