# Snapshot existing Release outputs and public fixtures. Does not build or test.
[CmdletBinding(DefaultParameterSetName='Single')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='Single')] [string] $OutputDirectory,
    [Parameter(ParameterSetName='Single')] [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory=$true, ParameterSetName='Batch')] [hashtable] $OutputDirectories
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'QualificationBuildIdentity.ps1')
function New-FrozenHardLink([string] $Path, [string] $Target) {
    # The Windows provider needs extended paths in deeply nested worktrees.
    if ($IsWindows) {
        $Path = if ($Path.StartsWith('\\')) { '\\?\UNC\' + $Path.Substring(2) } else { '\\?\' + $Path }
        $Target = if ($Target.StartsWith('\\')) { '\\?\UNC\' + $Target.Substring(2) } else { '\\?\' + $Target }
    }
    New-Item -ItemType HardLink -Path $Path -Target $Target | Out-Null
}
if ($PSCmdlet.ParameterSetName -eq 'Single') {
    $OutputDirectories = @{ $Framework = $OutputDirectory }
}
if (!$OutputDirectories.Count) { throw 'Select at least one framework to freeze' }
$outputs = [ordered]@{}
$pathComparer = if ($IsWindows) { [StringComparer]::OrdinalIgnoreCase } else { [StringComparer]::Ordinal }
$paths = [Collections.Generic.HashSet[string]]::new($pathComparer)
foreach ($targetFramework in @($OutputDirectories.Keys | Sort-Object -Descending)) {
    $path = $OutputDirectories[$targetFramework]
    if ($targetFramework -notin @('net10.0','net11.0') -or $path -isnot [string] -or [string]::IsNullOrWhiteSpace($path)) {
        throw 'Each frozen output requires a supported framework and a directory'
    }
    $targetFramework = $targetFramework.ToLowerInvariant()
    $fullPath = [IO.Path]::TrimEndingDirectorySeparator($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path))
    if ($outputs.Contains($targetFramework) -or !$paths.Add($fullPath) -or (Test-Path -LiteralPath $fullPath)) {
        throw 'Each immutable artifact requires a distinct fresh directory'
    }
    $outputs[$targetFramework] = $fullPath
}
$status = @(git -C $repo status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0 -or $status.Count) { throw 'Commit source and public inputs before freezing' }
$revision = git -C $repo rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Cannot pin source revision' }
# A test-only rebuild can update the library while leaving the CLI copy stale.
# Reject mixed runtime copies before creating an immutable snapshot.
$testProject = 'tests/Lokad.OoxPdf.Tests'
$projects = @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', $testProject, 'tools/Lokad.OoxPdf.DocxInspect')
$buildLibraryHashes = @{}
$compiled = @{}
$verifiedProductionRevisions = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($targetFramework in $outputs.Keys) {
    $libraryCopies = $projects |
        ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $repo ($_ + "/bin/Release/$targetFramework/Lokad.OoxPdf.dll"))).Hash }
    if (@($libraryCopies | Sort-Object -Unique).Count -ne 1) {
        throw 'Release library copies differ; refresh the dependent outputs before freezing'
    }
    $buildLibraryHashes[$targetFramework] = $libraryCopies[0]
    foreach ($project in $projects) {
        $name = "$project/bin/Release/$targetFramework/$(Split-Path $project -Leaf).dll"
        $assembly = Read-QualificationAssemblyIdentity (Join-Path $repo $name)
        if ($assembly.Revision -cne $revision) {
            if ($project -eq $testProject) { throw 'Release test assembly was compiled for another source revision' }
            # A newly built harness can reuse production bytes when the source
            # and build inputs at the embedded revision still match HEAD.
            if ($verifiedProductionRevisions.Add($assembly.Revision)) {
                Assert-QualificationProductionInputs $repo $assembly.Revision $revision
            }
        }
        $compiled[$name] = $assembly
    }
    if ($compiled["src/Lokad.OoxPdf/bin/Release/$targetFramework/Lokad.OoxPdf.dll"].Sha256 -ne $libraryCopies[0]) {
        throw 'Release library changed during preflight'
    }
}
$inputRoot = 'tests/Lokad.OoxPdf.Tests/Cases'
$inputTree = git -C $repo rev-parse ("HEAD:" + $inputRoot)
if ($LASTEXITCODE -ne 0) { throw 'Cannot pin public input tree' }
# A clean status can still hide stale materialization or an index stat override.
# Compare raw bytes with committed blobs, without checkout filters or one Git
# process per fixture. Git's quoted paths round-trip through --stdin-paths.
$treeEntries = @(git -C $repo ls-tree -r '--format=%(objectname) %(path)' $revision -- $inputRoot)
if ($LASTEXITCODE -ne 0 -or !$treeEntries.Count) { throw 'Public input tree is unavailable' }
$materializedBlobs = @($treeEntries | ForEach-Object { $_.Substring($_.IndexOf(' ') + 1) } |
    git -C $repo hash-object --no-filters --stdin-paths)
if ($LASTEXITCODE -ne 0 -or $materializedBlobs.Count -ne $treeEntries.Count) { throw 'Cannot verify public input Git blobs' }
for ($i = 0; $i -lt $treeEntries.Count; $i++) {
    $separator = $treeEntries[$i].IndexOf(' ')
    if ($materializedBlobs[$i] -ne $treeEntries[$i].Substring(0, $separator)) {
        throw ('Public input bytes differ from committed Git blob: ' + $treeEntries[$i].Substring($separator + 1))
    }
}
$commonDirectory = git -C $repo rev-parse --path-format=absolute --git-common-dir
if ($LASTEXITCODE -ne 0) { throw 'Cannot locate shared Git workspace' }
# Earlier snapshots could contain noncanonical checkout bytes. Keep them intact
# for historical evidence; the canonical namespace cannot collide with them.
$snapshotName = 'git-' + $inputTree
$snapshot = Join-Path (Split-Path $commonDirectory) ('artifacts/qualification-inputs/' + $snapshotName)
$inputs = [ordered]@{}
foreach ($name in @(git -C $repo ls-files $inputRoot)) {
    $inputs[$name] = (Get-FileHash -LiteralPath (Join-Path $repo $name)).Hash
}
if ($LASTEXITCODE -ne 0 -or !$inputs.Count) { throw 'Public inputs are unavailable' }
# Publish one complete fixture snapshot per Git tree, shared across worktrees.
# Subsequent freezes still hash every input; reuse never trusts a directory name.
$reusedInputs = Test-Path -LiteralPath $snapshot
if (!$reusedInputs) {
    # One independently owned copy per input tree; frozen runtimes reuse its
    # directory link. Clearing a read-only cross-tree hard link during failure
    # cleanup would also change a published input's protection.
    $staging = [IO.Path]::GetFullPath($snapshot + '.staging-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    try {
        foreach ($name in $inputs.Keys) {
            $target = Join-Path $staging $name
            New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $repo $name) -Destination $target
            if ((Get-FileHash -LiteralPath $target).Hash -ne $inputs[$name]) { throw 'Public input changed during snapshot' }
            (Get-Item -LiteralPath $target).IsReadOnly = $true
        }
        [ordered]@{Tree=$inputTree;Inputs=$inputs;SharedInputFiles=0} | ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath (Join-Path $staging 'identity.json') -Encoding utf8
        # Directory.Move cannot nest staging inside a concurrent winner. Verify
        # the winner's complete identity below, just as for ordinary reuse.
        try { [IO.Directory]::Move($staging, $snapshot) }
        catch { if (!(Test-Path -LiteralPath $snapshot)) { throw } }
    }
    finally {
        if (Test-Path -LiteralPath $staging) {
            $parent = [IO.Path]::GetFullPath((Split-Path $snapshot)) + [IO.Path]::DirectorySeparatorChar
            if (!$staging.StartsWith($parent, [StringComparison]::OrdinalIgnoreCase) -or
                !(Split-Path $staging -Leaf).StartsWith($snapshotName + '.staging-', [StringComparison]::Ordinal)) {
                throw 'Refusing cleanup outside the owned input staging directory'
            }
            # These new copies have no aliases in a checkout or prior snapshot.
            Get-ChildItem -LiteralPath $staging -File -Recurse -Force |
                ForEach-Object { $_.IsReadOnly = $false }
            Remove-Item -LiteralPath $staging -Recurse -Force
        }
    }
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
$inputIdentityHash = (Get-FileHash -LiteralPath $inputIdentity).Hash
$linkType = if ($IsWindows) { 'Junction' } else { 'SymbolicLink' }
foreach ($targetFramework in $outputs.Keys) {
    $output = $outputs[$targetFramework]
    New-Item -ItemType Directory -Path $output | Out-Null
    $assets = [ordered]@{}
    $assemblyRevisions = [ordered]@{}
    $assetCopies = @{}
    $sharedAssetFiles = 0
    $sharedAssetBytes = 0L
    foreach ($project in $projects) {
        $relative = $project + "/bin/Release/$targetFramework"
        $source = Join-Path $repo $relative
        if (!(Test-Path -LiteralPath $source)) { throw ('Missing Release output: ' + $project) }
        $target = Join-Path $output $relative
        New-Item -ItemType Directory -Force -Path $target | Out-Null
        foreach ($file in Get-ChildItem -LiteralPath $source -File) {
            $name = $relative + '/' + $file.Name
            $hash = (Get-FileHash -LiteralPath $file.FullName).Hash
            if ($compiled.ContainsKey($name)) {
                if ($compiled[$name].Sha256 -ne $hash) { throw 'Release assembly changed during freeze' }
                $assemblyRevisions[$name] = $compiled[$name].Revision
            }
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
        if ($assets["$relative/Lokad.OoxPdf.dll"] -ne $buildLibraryHashes[$targetFramework]) {
            throw 'Release library changed during freeze'
        }
    }
    $testAssemblyHash = $assets["tests/Lokad.OoxPdf.Tests/bin/Release/$targetFramework/Lokad.OoxPdf.Tests.dll"]
    if ($assemblyRevisions.Count -ne $projects.Count) { throw 'Release assembly disappeared during freeze' }
    $inputLink = Join-Path $output $inputRoot
    New-Item -ItemType Directory -Path (Split-Path $inputLink) -Force | Out-Null
    New-Item -ItemType $linkType -Path $inputLink -Target (Join-Path $snapshot $inputRoot) | Out-Null
    [ordered]@{
        Revision=$revision;Framework=$targetFramework;Rebuilt=$false;Assets=$assets;AssemblyRevisions=$assemblyRevisions;Inputs=$inputs
        InputTree=$inputTree;InputSnapshot=$snapshot;ReusedInputSnapshot=$reusedInputs
        SharedInputFiles=$inputPin.SharedInputFiles
        SharedAssetFiles=$sharedAssetFiles;SharedAssetBytes=$sharedAssetBytes
        InputSnapshotIdentitySha256=$inputIdentityHash
        LibrarySha256=$buildLibraryHashes[$targetFramework];TestAssemblySha256=$testAssemblyHash
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
}
Write-Output "Froze $($outputs.Count) framework artifact(s) at $revision with one verification of $($inputs.Count) public inputs (reused snapshot: $reusedInputs)."
