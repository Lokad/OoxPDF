# Transfer an already qualified artifact and run only explicit integration smoke jobs.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $ArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $QualificationEvidence,
    [Parameter(Mandatory=$true)] [string] $SmokeManifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$artifact = (Resolve-Path -LiteralPath $ArtifactDirectory).Path
$pin = Get-Content (Join-Path $artifact 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
$qualification = Get-Content -LiteralPath $QualificationEvidence -Raw | ConvertFrom-Json
if ($qualification.LibrarySha256 -ne $pin.LibrarySha256 -or
    $qualification.TestAssemblySha256 -ne $pin.TestAssemblySha256) { throw 'Qualification identities differ' }
$source = @(git -C $repo ls-tree -r $pin.Revision -- src tests)
if ($LASTEXITCODE -ne 0) { throw 'Frozen source revision is unavailable' }
$current = @(git -C $repo ls-tree -r HEAD -- src tests)
if ($LASTEXITCODE -ne 0 -or @(Compare-Object $source $current).Count -or
    @(git -C $repo status --porcelain --untracked-files=normal -- src tests).Count) { throw 'Source/input transfer identity differs' }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'A fresh transfer evidence directory is required' }
$jobs = @(Get-Content -LiteralPath $SmokeManifest -Raw | ConvertFrom-Json)
if (!$jobs.Count) { throw 'Select a small nonempty integration smoke set' }
# Check every source and asset before mutating root outputs.
foreach ($name in $pin.Assets.Keys) {
    if ((Get-FileHash -LiteralPath (Join-Path $artifact $name)).Hash -ne $pin.Assets[$name]) { throw 'Frozen asset differs' }
}
foreach ($job in $jobs) {
    if ((Get-FileHash -LiteralPath $job.Input).Hash -ne $job.InputSha256) { throw 'Smoke input differs' }
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
foreach ($name in $pin.Assets.Keys) {
    $target = Join-Path $repo $name
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    Copy-Item -LiteralPath (Join-Path $artifact $name) -Destination $target -Force
    if ((Get-FileHash -LiteralPath $target).Hash -ne $pin.Assets[$name]) { throw 'Transferred asset differs' }
}
$cli = Join-Path $repo 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.Cli.dll'
$rows = @()
foreach ($job in $jobs) {
    $pdf = Join-Path $OutputDirectory ($rows.Count.ToString() + '.pdf')
    dotnet $cli convert $job.Input $pdf @($job.Arguments) *> ($pdf+'.log')
    if ($LASTEXITCODE -ne 0 -or (Get-FileHash $pdf).Hash -ne $job.PdfSha256) { throw 'Transfer smoke PDF differs' }
    $rows += $job
}
[ordered]@{FrozenRevision=$pin.Revision;LibrarySha256=$pin.LibrarySha256;AssetsVerified=$pin.Assets.Count;Rebuilt=$false;RepeatedSuite=$false;Smoke=$rows} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
Write-Output "$($pin.Assets.Count) assets transferred; $($rows.Count) smoke PDFs match. No build or suite replay."
