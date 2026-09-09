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
