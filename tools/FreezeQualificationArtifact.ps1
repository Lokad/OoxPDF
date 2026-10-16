# Snapshot existing Release outputs and public fixtures. Does not build or test.
[CmdletBinding()]
param([Parameter(Mandatory=$true)] [string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
function New-FrozenHardLink([string] $Path, [string] $Target) {
    # The Windows provider needs extended paths in deeply nested worktrees.
    if ($IsWindows) {
        $Path = if ($Path.StartsWith('\\')) { '\\?\UNC\' + $Path.Substring(2) } else { '\\?\' + $Path }
        $Target = if ($Target.StartsWith('\\')) { '\\?\UNC\' + $Target.Substring(2) } else { '\\?\' + $Target }
    }
    New-Item -ItemType HardLink -Path $Path -Target $Target | Out-Null
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'An immutable artifact requires a fresh directory' }
$status = @(git -C $repo status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0 -or $status.Count) { throw 'Commit source and public inputs before freezing' }
# A test-only rebuild can update the library while leaving the CLI copy stale.
# Reject mixed runtime copies before creating an immutable snapshot.
$libraryCopies = @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests') |
    ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $repo ($_ + '/bin/Release/net10.0/Lokad.OoxPdf.dll'))).Hash }
if (@($libraryCopies | Sort-Object -Unique).Count -ne 1) {
    throw 'Release library copies differ; refresh the dependent outputs before freezing'
}
$revision = git -C $repo rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Cannot pin source revision' }
$inputRoot = 'tests/Lokad.OoxPdf.Tests/Cases'
$inputTree = git -C $repo rev-parse ("HEAD:" + $inputRoot)
if ($LASTEXITCODE -ne 0) { throw 'Cannot pin public input tree' }
$commonDirectory = git -C $repo rev-parse --path-format=absolute --git-common-dir
if ($LASTEXITCODE -ne 0) { throw 'Cannot locate shared Git workspace' }
$snapshot = Join-Path (Split-Path $commonDirectory) ('artifacts/qualification-inputs/' + $inputTree)
$inputs = [ordered]@{}
foreach ($name in @(git -C $repo ls-files $inputRoot)) {
    $inputs[$name] = (Get-FileHash -LiteralPath (Join-Path $repo $name)).Hash
}
if ($LASTEXITCODE -ne 0 -or !$inputs.Count) { throw 'Public inputs are unavailable' }
# Publish one complete fixture snapshot per Git tree, shared across worktrees.
# Subsequent freezes still hash every input; reuse never trusts a directory name.
$reusedInputs = Test-Path -LiteralPath $snapshot
if (!$reusedInputs) {
    # Seed only from a published immutable snapshot. A writable checkout must
    # never share file ownership with a frozen input, even if its hash matches.
    $prior = Get-ChildItem -LiteralPath (Split-Path $snapshot) -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^(?:[0-9a-f]{40}|[0-9a-f]{64})$' -and $_.FullName -ne $snapshot } |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    $priorPin = if ($prior) { Get-Content (Join-Path $prior.FullName 'identity.json') -Raw | ConvertFrom-Json -AsHashtable }
    if ($prior -and $priorPin.Tree -ne $prior.Name) { throw 'Prior shared input identity differs' }
    $sharedInputFiles = 0
    $staging = $snapshot + '.staging-' + [Guid]::NewGuid().ToString('N')
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    foreach ($name in $inputs.Keys) {
        $target = Join-Path $staging $name
        New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
        if ($prior -and $priorPin.Inputs[$name] -eq $inputs[$name]) {
            $shared = Join-Path $prior.FullName $name
            if ((Get-FileHash -LiteralPath $shared).Hash -ne $inputs[$name]) { throw 'Prior shared input differs' }
            if (!(Get-Item -LiteralPath $shared).IsReadOnly) { throw 'Prior shared input is writable' }
            New-FrozenHardLink -Path $target -Target $shared
            $sharedInputFiles++
        }
        else { Copy-Item -LiteralPath (Join-Path $repo $name) -Destination $target }
        if ((Get-FileHash -LiteralPath $target).Hash -ne $inputs[$name]) { throw 'Public input changed during snapshot' }
        (Get-Item -LiteralPath $target).IsReadOnly = $true
    }
    [ordered]@{Tree=$inputTree;Inputs=$inputs;SharedInputFiles=$sharedInputFiles} | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath (Join-Path $staging 'identity.json') -Encoding utf8
    # Directory.Move cannot accidentally nest staging inside a concurrent snapshot.
    try { [IO.Directory]::Move($staging, $snapshot) }
    catch { if (!(Test-Path -LiteralPath $snapshot)) { throw } }
}
$inputIdentity = Join-Path $snapshot 'identity.json'
$inputPin = Get-Content -LiteralPath $inputIdentity -Raw | ConvertFrom-Json -AsHashtable
if ($inputPin.Tree -ne $inputTree -or $inputPin.Inputs.Count -ne $inputs.Count -or
    @(Get-ChildItem -LiteralPath (Join-Path $snapshot $inputRoot) -File -Recurse -Force).Count -ne $inputs.Count) {
    throw 'Shared public input identity differs'
}
foreach ($name in $inputs.Keys) {
    if ($inputPin.Inputs[$name] -ne $inputs[$name] -or
        (Get-FileHash -LiteralPath (Join-Path $snapshot $name)).Hash -ne $inputs[$name]) {
        throw 'Shared public input differs'
    }
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$assets = [ordered]@{}
$assetCopies = @{}
$sharedAssetFiles = 0
$sharedAssetBytes = 0L
foreach ($project in @('src/Lokad.OoxPdf','src/Lokad.OoxPdf.Cli','tests/Lokad.OoxPdf.Tests')) {
    $relative = $project + '/bin/Release/net10.0'
    $source = Join-Path $repo $relative
    if (!(Test-Path -LiteralPath $source)) { throw ('Missing Release output: ' + $project) }
    $target = Join-Path $output $relative
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $source -File) {
        $name = $relative + '/' + $file.Name
        $hash = (Get-FileHash -LiteralPath $file.FullName).Hash
        $destination = Join-Path $output $name
        if ($assetCopies.ContainsKey($hash)) {
            # Share only inside this fresh immutable output. Linking a writable
            # build or a prior artifact would give it ownership of these bytes.
            New-FrozenHardLink -Path $destination -Target $assetCopies[$hash]
            $sharedAssetFiles++
            $sharedAssetBytes += $file.Length
        }
        else {
            Copy-Item -LiteralPath $file.FullName -Destination $destination
            $assetCopies[$hash] = $destination
        }
        if ((Get-FileHash -LiteralPath $destination).Hash -ne $hash) { throw 'Frozen asset differs' }
        $assets[$name] = $hash
    }
}
$inputLink = Join-Path $output $inputRoot
New-Item -ItemType Directory -Path (Split-Path $inputLink) -Force | Out-Null
$linkType = if ($IsWindows) { 'Junction' } else { 'SymbolicLink' }
New-Item -ItemType $linkType -Path $inputLink -Target (Join-Path $snapshot $inputRoot) | Out-Null
[ordered]@{
    Revision=$revision;Rebuilt=$false;Assets=$assets;Inputs=$inputs
    InputTree=$inputTree;InputSnapshot=$snapshot;ReusedInputSnapshot=$reusedInputs
    SharedInputFiles=$inputPin.SharedInputFiles
    SharedAssetFiles=$sharedAssetFiles;SharedAssetBytes=$sharedAssetBytes
    InputSnapshotIdentitySha256=(Get-FileHash -LiteralPath $inputIdentity).Hash
    LibrarySha256=(Get-FileHash -LiteralPath (Join-Path $output 'src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll')).Hash
    TestAssemblySha256=(Get-FileHash -LiteralPath (Join-Path $output 'tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/Lokad.OoxPdf.Tests.dll')).Hash
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
Write-Output "Frozen $($assets.Count) assets at $revision; $($inputs.Count) public inputs verified in shared snapshot (reused: $reusedInputs)."
