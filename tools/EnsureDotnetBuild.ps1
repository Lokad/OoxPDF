# Shared .NET build gate for visual/tool workflows (T03).
#
# Hand-rolled "rebuild if any .cs file is newer than the DLL" checks silently
# select stale builds: they miss root props/targets, confuse copied/touched
# timestamps, and race concurrent writers on shared bin/obj outputs. This
# helper delegates up-to-date tracking to MSBuild itself (dotnet build is a
# no-op when nothing changed), serializes concurrent builds with a
# machine-wide mutex, fails loudly (never falls back to a stale DLL), and
# optionally records source/binary identities for run provenance.

param(
    [Parameter(Mandatory = $true)]
    [string] $Project,

    [Parameter(Mandatory = $true)]
    [string] $OutputDll,

    [Parameter(Mandatory = $true)]
    [string] $Description,

    [string] $Configuration = "Debug",

    [string] $RecordPath
)

$ErrorActionPreference = "Stop"

$projectFull = (Resolve-Path -LiteralPath $Project).Path
$repoRoot = Split-Path -Parent $PSScriptRoot
$mutex = New-Object System.Threading.Mutex($false, "Global\OoxPdfDotnetBuild")
try {
    if (-not $mutex.WaitOne([TimeSpan]::FromMinutes(15))) {
        throw "$Description build unavailable: another build holds the lock past the deadline."
    }

    try {
        & dotnet build $projectFull -c $Configuration --nologo
        if ($LASTEXITCODE -ne 0) {
            throw "$Description build failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        $mutex.ReleaseMutex()
    }
}
finally {
    $mutex.Dispose()
}

if (-not (Test-Path -LiteralPath $OutputDll)) {
    throw "$Description build succeeded but the expected output is missing: $OutputDll. Refusing to use a stale DLL."
}

if (-not [string]::IsNullOrWhiteSpace($RecordPath)) {
    $commit = "unknown"
    $dirty = $null
    try {
        $commit = (& git -C $repoRoot rev-parse HEAD 2>$null | Select-Object -First 1).Trim()
        $dirty = [bool](& git -C $repoRoot status --short 2>$null | Select-Object -First 1)
    }
    catch {
    }

    $rootFiles = [ordered]@{}
    foreach ($name in @("Directory.Build.props", "Directory.Build.rsp")) {
        $path = Join-Path $repoRoot $name
        $rootFiles[$name] = if (Test-Path -LiteralPath $path) {
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        }
        else {
            $null
        }
    }

    [ordered]@{
        Description = $Description
        Project = $projectFull.Substring($repoRoot.Length).TrimStart("\", "/").Replace("\", "/")
        ProjectSha256 = (Get-FileHash -LiteralPath $projectFull -Algorithm SHA256).Hash.ToLowerInvariant()
        Configuration = $Configuration
        RootFiles = $rootFiles
        OutputDll = $OutputDll
        OutputSha256 = (Get-FileHash -LiteralPath $OutputDll -Algorithm SHA256).Hash.ToLowerInvariant()
        GitCommit = $commit
        GitDirty = $dirty
        BuiltAtUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $RecordPath -Encoding UTF8
}
