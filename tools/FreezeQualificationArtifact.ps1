# Snapshot existing Release outputs and public fixtures. Does not build or test.
[CmdletBinding()]
param([Parameter(Mandatory=$true)] [string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
if (Test-Path -LiteralPath $OutputDirectory) { throw 'An immutable artifact requires a fresh directory' }
$status = @(git -C $repo status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0 -or $status.Count) { throw 'Commit source and public inputs before freezing' }
$revision = git -C $repo rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Cannot pin source revision' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$assets = [ordered]@{}
foreach ($project in @('src/Lokad.OoxPdf','src/Lokad.OoxPdf.Cli','tests/Lokad.OoxPdf.Tests')) {
    $relative = $project + '/bin/Release/net10.0'
    $source = Join-Path $repo $relative
    if (!(Test-Path -LiteralPath $source)) { throw ('Missing Release output: ' + $project) }
    $target = Join-Path $output $relative
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $source -File) {
        $name = $relative + '/' + $file.Name
        $hash = (Get-FileHash -LiteralPath $file.FullName).Hash
        Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $output $name)
        if ((Get-FileHash -LiteralPath (Join-Path $output $name)).Hash -ne $hash) { throw 'Frozen asset differs' }
        $assets[$name] = $hash
    }
}
$inputs = [ordered]@{}
foreach ($name in @(git -C $repo ls-files tests/Lokad.OoxPdf.Tests/Cases)) {
    $source = Join-Path $repo $name
    $target = Join-Path $output $name
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    $hash = (Get-FileHash -LiteralPath $source).Hash
    Copy-Item -LiteralPath $source -Destination $target
    if ((Get-FileHash -LiteralPath $target).Hash -ne $hash) { throw 'Frozen public input differs' }
    $inputs[$name] = $hash
}
[ordered]@{
    Revision=$revision;Rebuilt=$false;Assets=$assets;Inputs=$inputs
    LibrarySha256=(Get-FileHash -LiteralPath (Join-Path $output 'src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll')).Hash
    TestAssemblySha256=(Get-FileHash -LiteralPath (Join-Path $output 'tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/Lokad.OoxPdf.Tests.dll')).Hash
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
Write-Output "Frozen $($assets.Count) assets and $($inputs.Count) public fixture files at $revision."
