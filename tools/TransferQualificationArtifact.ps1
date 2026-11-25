# Transfer an already qualified artifact and run only explicit integration smoke jobs.
# Smoke jobs default to ExpectedExitCode 0. Exit 3 also requires DiagnosticsSha256.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $ArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $QualificationEvidence,
    [Parameter(Mandatory=$true)] [string] $SmokeManifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
. (Join-Path $PSScriptRoot 'QualificationBuildIdentity.ps1')
$repo = Split-Path -Parent $PSScriptRoot
$artifact = (Resolve-Path -LiteralPath $ArtifactDirectory).Path
$pin = Get-Content (Join-Path $artifact 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
$framework = if ($pin.Framework) { $pin.Framework } else { 'net10.0' }
if ($framework -notin @('net10.0','net11.0')) { throw 'Unknown frozen target framework' }
$qualification = Get-Content -LiteralPath $QualificationEvidence -Raw | ConvertFrom-Json
if ($qualification.LibrarySha256 -ne $pin.LibrarySha256 -or
    $qualification.TestAssemblySha256 -ne $pin.TestAssemblySha256) { throw 'Qualification identities differ' }
# The transfer also binds the complete test tree and the working tree.
$buildInputs = @(Get-QualificationProductionInputs) + @('tests')
$changes = @(git -C $repo diff --name-only --no-ext-diff $pin.Revision HEAD -- @buildInputs)
if ($LASTEXITCODE -ne 0) { throw 'Frozen source revision is unavailable' }
$workingChanges = @(git -C $repo status --porcelain --untracked-files=normal -- @buildInputs)
if ($LASTEXITCODE -ne 0) { throw 'Cannot verify transfer working-tree inputs' }
if ($changes.Count -or $workingChanges.Count) { throw 'Source/input/build-setting transfer identity differs' }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'A fresh transfer evidence directory is required' }
$jobs = @(Get-Content -LiteralPath $SmokeManifest -Raw | ConvertFrom-Json)
if (!$jobs.Count) { throw 'Select a small nonempty integration smoke set' }
# Check every source and asset before mutating root outputs.
foreach ($name in $pin.Assets.Keys) {
    if ((Get-FileHash -LiteralPath (Join-Path $artifact $name)).Hash -ne $pin.Assets[$name]) { throw 'Frozen asset differs' }
    $directory = Split-Path -Parent (Join-Path $repo $name)
    while ($directory.Length -gt $repo.Length) {
        if ((Test-Path -LiteralPath $directory) -and
            ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Root asset directories must not be linked'
        }
        $directory = Split-Path -Parent $directory
    }
}
foreach ($job in $jobs) {
    if ((Get-FileHash -LiteralPath $job.Input).Hash -ne $job.InputSha256) { throw 'Smoke input differs' }
    $expectedExit = $job.ExpectedExitCode ?? 0
    if (($expectedExit -isnot [int] -and $expectedExit -isnot [long]) -or $expectedExit -notin @(0,3)) {
        throw 'Smoke expected exit must be 0 or 3'
    }
    if ($expectedExit -eq 3 -or $null -ne $job.DiagnosticsSha256) {
        if ($job.DiagnosticsSha256 -isnot [string] -or $job.DiagnosticsSha256 -notmatch '^[0-9a-f]{64}$') {
            throw 'Smoke diagnostics require a SHA-256 pin'
        }
        if ($job.Arguments -contains '--diagnostics') { throw 'Smoke diagnostics path is owned by the transfer' }
    }
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
foreach ($name in $pin.Assets.Keys) {
    $target = Join-Path $repo $name
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    # Replace the entry instead of writing through an existing hard/symbolic link.
    $temporary = $target + '.staging-' + [Guid]::NewGuid().ToString('N')
    $ownsTemporary = $false
    try {
        $source = [IO.File]::OpenRead((Join-Path $artifact $name))
        try {
            $destination = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $ownsTemporary = $true
            try { $source.CopyTo($destination) }
            finally { $destination.Dispose() }
        }
        finally { $source.Dispose() }
        if ((Get-FileHash -LiteralPath $temporary).Hash -ne $pin.Assets[$name]) { throw 'Staged asset differs' }
        [IO.File]::Move($temporary, $target, $true)
        $ownsTemporary = $false
    }
    finally {
        if ($ownsTemporary -and [IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
    if ((Get-FileHash -LiteralPath $target).Hash -ne $pin.Assets[$name]) { throw 'Transferred asset differs' }
}
$cli = Join-Path $repo "src/Lokad.OoxPdf.Cli/bin/Release/$framework/Lokad.OoxPdf.Cli.dll"
$rows = @()
foreach ($job in $jobs) {
    $outputBase = Join-Path $OutputDirectory ($rows.Count.ToString())
    $pdf = $outputBase + '.pdf'
    $arguments = @($job.Arguments)
    if ($job.DiagnosticsSha256) {
        $diagnostics = $outputBase + '.diagnostics.json'
        $arguments += @('--diagnostics', $diagnostics)
    }
    dotnet $cli convert $job.Input $pdf @arguments *> ($pdf+'.log')
    if ($LASTEXITCODE -ne ($job.ExpectedExitCode ?? 0)) { throw 'Transfer smoke exit differs' }
    if (!(Test-Path -LiteralPath $pdf -PathType Leaf) -or (Get-FileHash -LiteralPath $pdf).Hash -ne $job.PdfSha256) {
        throw 'Transfer smoke PDF differs'
    }
    if ($job.DiagnosticsSha256 -and (!(Test-Path -LiteralPath $diagnostics -PathType Leaf) -or
        (Get-FileHash -LiteralPath $diagnostics).Hash -ne $job.DiagnosticsSha256)) { throw 'Transfer smoke diagnostics differ' }
    $rows += $job
}
[ordered]@{FrozenRevision=$pin.Revision;Framework=$framework;LibrarySha256=$pin.LibrarySha256;AssetsVerified=$pin.Assets.Count;Rebuilt=$false;RepeatedSuite=$false;Smoke=$rows} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
Write-Output "$($pin.Assets.Count) assets transferred; $($rows.Count) smoke PDFs match. No build or suite replay."
