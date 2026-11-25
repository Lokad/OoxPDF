# Exercise fixture sharing and rejection paths without building or running the library.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/qualification-freeze-test-' + [Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$repo = (Resolve-Path -LiteralPath $repo).Path
$script = Join-Path $repo 'tools/FreezeQualificationArtifact.ps1'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'FreezeQualificationArtifact.ps1') -Destination $script
$parentScript = Join-Path $repo 'tools/NewParentRegressionRuntime.ps1'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'NewParentRegressionRuntime.ps1') -Destination $parentScript
$transferScript = Join-Path $repo 'tools/TransferQualificationArtifact.ps1'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'TransferQualificationArtifact.ps1') -Destination $transferScript
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'BuildQualificationArtifact.ps1') -Destination (Join-Path $repo 'tools/BuildQualificationArtifact.ps1')
$assemblyIdentitySource = Join-Path $PSScriptRoot 'QualificationBuildIdentity.ps1'
$assemblyIdentityFixture = Join-Path $repo 'tools/QualificationBuildIdentity.ps1'
# Ownership controls use synthetic files. Exercise the real PE reader below.
Copy-Item -LiteralPath $assemblyIdentitySource -Destination $assemblyIdentityFixture
@'
function Read-QualificationAssemblyIdentity([string] $Path) {
    $revision = git -C (Split-Path $Path -Parent) rev-parse HEAD
    if ($LASTEXITCODE -ne 0) { throw 'Cannot bind synthetic fixture revision' }
    @{ Revision = $revision; Sha256 = (Get-FileHash -LiteralPath $Path).Hash }
}
'@ | Add-Content -LiteralPath $assemblyIdentityFixture
function Invoke-Git([string[]] $Arguments) {
    $result = git -C $repo @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw ($result -join "`n") }
}
Invoke-Git @('init', '--quiet')
Invoke-Git @('config', 'user.name', 'Qualification test')
Invoke-Git @('config', 'user.email', 'qualification@example.invalid')
Invoke-Git @('config', 'core.autocrlf', 'false')
if ($IsWindows) { Invoke-Git @('config', 'core.longpaths', 'true') }
@('bin/', 'artifacts/') | Set-Content -LiteralPath (Join-Path $repo '.gitignore')
'<Project />' | Set-Content -LiteralPath (Join-Path $repo 'Directory.Build.props')
Copy-Item -LiteralPath (Join-Path (Split-Path $PSScriptRoot) '.gitattributes') -Destination $repo
$inputPath = 'tests/Lokad.OoxPdf.Tests/Cases/sample.txt'
$inspectorSource = Join-Path $repo 'tools/Lokad.OoxPdf.DocxInspect/Program.cs'
New-Item -ItemType Directory -Path (Split-Path $inspectorSource) -Force | Out-Null
'// owned inspector source' | Set-Content -LiteralPath $inspectorSource
$inputFile = Join-Path $repo $inputPath
New-Item -ItemType Directory -Path (Split-Path $inputFile) -Force | Out-Null
[IO.File]::WriteAllText($inputFile, "first fixture`n")
$nested = Join-Path (Split-Path $inputFile) 'nested/second.txt'
New-Item -ItemType Directory -Path (Split-Path $nested) -Force | Out-Null
[IO.File]::WriteAllText($nested, "nested fixture`n")
$spaced = Join-Path (Split-Path $inputFile) 'nested/spaced fixture.txt'
[IO.File]::WriteAllText($spaced, "committed CRLF fixture`r`n")
foreach ($project in @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests', 'tools/Lokad.OoxPdf.DocxInspect')) {
    $projectName=Split-Path $project -Leaf
    New-Item -ItemType Directory -Path (Join-Path $repo $project) -Force | Out-Null
    '<Project />' | Set-Content -LiteralPath (Join-Path $repo ($project+'/'+$projectName+'.csproj'))
    foreach ($targetFramework in 'net10.0','net11.0') {
        $directory = Join-Path $repo ($project + "/bin/Release/$targetFramework")
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        "fake library for $targetFramework; no executable code" | Set-Content -LiteralPath (Join-Path $directory 'Lokad.OoxPdf.dll')
        "fake test assembly for $targetFramework" | Set-Content -LiteralPath (Join-Path $directory 'Lokad.OoxPdf.Tests.dll')
        if ($projectName -notin @('Lokad.OoxPdf', 'Lokad.OoxPdf.Tests')) {
            "fake $projectName assembly for $targetFramework" | Set-Content -LiteralPath (Join-Path $directory ($projectName+'.dll'))
        }
    }
}
foreach ($name in @('Lokad.OoxPdf.Tests.pdb','Lokad.OoxPdf.Tests.deps.json','Lokad.OoxPdf.Tests.runtimeconfig.json')) {
    'fake runner support; no executable code' | Set-Content -LiteralPath (Join-Path $repo ("tests/Lokad.OoxPdf.Tests/bin/Release/$Framework/" + $name))
}
Invoke-Git @('add', '.')
Invoke-Git @('commit', '--quiet', '-m', 'Initial fixtures')
function Freeze([string] $Name) {
    $target = Join-Path $repo ('artifacts/' + $Name)
    & $script -OutputDirectory $target -Framework $Framework | Out-Null
    return (Get-Content -LiteralPath (Join-Path $target 'identity.json') -Raw | ConvertFrom-Json)
}
function Reject([string] $Name, [string] $ExpectedMessage) {
    try { $null = Freeze $Name }
    catch {
        if (!$_.Exception.Message.Contains($ExpectedMessage)) { throw }
        if (Test-Path -LiteralPath (Join-Path $repo ('artifacts/' + $Name))) { throw 'Rejection published a partial artifact' }
        return
    }
    throw ('Expected rejection: ' + $Name)
}
function Reject-Noncanonical([string] $Name, [string] $Bytes) {
    # Force a clean status despite altered bytes; this models the observed gap
    # without claiming how the historical checkout acquired its CRLF files.
    Invoke-Git @('update-index', '--assume-unchanged', $inputPath)
    try {
        [IO.File]::WriteAllText($inputFile, $Bytes)
        Reject $Name 'Public input bytes differ from committed Git blob'
    }
    finally {
        [IO.File]::WriteAllText($inputFile, "first fixture`n")
        Invoke-Git @('update-index', '--no-assume-unchanged', $inputPath)
    }
}
Reject-Noncanonical 'noncanonical-before-snapshot' "first fixture`r`n"
if (Test-Path -LiteralPath (Join-Path $repo 'artifacts/qualification-inputs')) { throw 'Noncanonical rejection created a snapshot' }
# A retained preflight-era snapshot may own CRLF bytes under the same tree.
# New canonical snapshots must neither reuse nor overwrite that evidence.
$legacyTree = git -C $repo rev-parse 'HEAD:tests/Lokad.OoxPdf.Tests/Cases'
$legacySnapshot = Join-Path $repo ('artifacts/qualification-inputs/' + $legacyTree)
$legacyFile = Join-Path $legacySnapshot $inputPath
New-Item -ItemType Directory -Path (Split-Path $legacyFile) -Force | Out-Null
[IO.File]::WriteAllText($legacyFile, "first fixture`r`n")
(Get-Item -LiteralPath $legacyFile).IsReadOnly = $true
$first = Freeze 'first'
if ($first.InputSnapshot -eq $legacySnapshot -or [IO.File]::ReadAllText($legacyFile) -cne "first fixture`r`n" -or
    !(Get-Item -LiteralPath $legacyFile).IsReadOnly) { throw 'Canonical freeze changed retained legacy evidence' }
if ([IO.File]::ReadAllText((Join-Path $first.InputSnapshot 'tests/Lokad.OoxPdf.Tests/Cases/nested/spaced fixture.txt')) -cne "committed CRLF fixture`r`n") {
    throw 'Committed CRLF bytes or spaced fixture path changed'
}
Reject-Noncanonical 'noncanonical-reuse' "first fixture`r`n"
Reject-Noncanonical 'same-length-noncanonical' "other fixture`n"
$second = Freeze 'second'
$batch = @{
    'net11.0' = Join-Path $repo 'artifacts/batch-net11'
    'net10.0' = Join-Path $repo 'artifacts/batch-net10'
}
$hashReads = [Collections.Generic.List[string]]::new()
& {
    function Get-FileHash([string] $LiteralPath) {
        $hashReads.Add($LiteralPath)
        Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $LiteralPath
    }
    & $script -OutputDirectories $batch | Out-Null
}
$batchPins = @{}
foreach ($targetFramework in $batch.Keys) {
    $identity = Get-Content -LiteralPath (Join-Path $batch[$targetFramework] 'identity.json') -Raw | ConvertFrom-Json
    $runtime = "src/Lokad.OoxPdf/bin/Release/$targetFramework/Lokad.OoxPdf.dll"
    $sourceLibrary = Join-Path $repo $runtime
    # The synthetic revision read hashes the same file as the real PE reader.
    # Its digest serves preflight, with a separate hash at the copy boundary.
    if (@($hashReads | Where-Object { $_ -eq $sourceLibrary }).Count -ne 2) {
        throw 'Primary library must be hashed once in preflight and once at freeze time'
    }
    if ($identity.Framework -cne $targetFramework -or $identity.Revision -ne $first.Revision -or
        $identity.InputSnapshot -ne $first.InputSnapshot -or $identity.InputTree -ne $first.InputTree -or
        $identity.InputSnapshotIdentitySha256 -ne $first.InputSnapshotIdentitySha256 -or
        $identity.LibrarySha256 -ne (Get-FileHash -LiteralPath (Join-Path $batch[$targetFramework] $runtime)).Hash -or
        !$identity.ReusedInputSnapshot) {
        throw 'Dual-framework freeze lost runtime or common input identity'
    }
    $batchPins[$targetFramework] = $identity
}
if ($batchPins['net10.0'].LibrarySha256 -eq $batchPins['net11.0'].LibrarySha256) {
    throw 'Dual-framework freeze conflated two distinct builds'
}
foreach ($fixture in $first.Inputs.PSObject.Properties.Name) {
    $checkout = Join-Path $repo $fixture
    $frozen = Join-Path $first.InputSnapshot $fixture
    if (@($hashReads | Where-Object { $_ -eq $checkout }).Count -ne 1 -or
        @($hashReads | Where-Object { $_ -eq $frozen }).Count -ne 1) {
        throw 'A dual-framework freeze must verify each public input only once'
    }
}
$brokenBatch = @{
    'net11.0' = Join-Path $repo 'artifacts/preflight-net11'
    'net10.0' = Join-Path $repo 'artifacts/preflight-net10'
}
$staleLibrary = Join-Path $repo 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.dll'
$originalLibrary = [IO.File]::ReadAllBytes($staleLibrary)
try {
    [IO.File]::WriteAllText($staleLibrary, 'different fake build')
    try {
        & $script -OutputDirectories $brokenBatch | Out-Null
        throw 'Expected refusal of mismatched net10 runtime in a batch'
    }
    catch {
        if (!$_.Exception.Message.Contains('Release library copies differ')) { throw }
    }
}
finally { [IO.File]::WriteAllBytes($staleLibrary, $originalLibrary) }
if (@($brokenBatch.Values | Where-Object { Test-Path -LiteralPath $_ }).Count) {
    throw 'Batch preflight created a partial artifact'
}
$expectedSharedAssets = @($first.Assets.psobject.Properties).Count - @($first.Assets.psobject.Properties.Value | Sort-Object -Unique).Count
if ($expectedSharedAssets -lt 1 -or $first.SharedAssetFiles -ne $expectedSharedAssets -or $first.SharedAssetBytes -lt 1) {
    throw 'Frozen duplicate assets were not shared'
}
$libraryRelative = "src/Lokad.OoxPdf/bin/Release/$Framework/Lokad.OoxPdf.dll"
$cliLibraryRelative = "src/Lokad.OoxPdf.Cli/bin/Release/$Framework/Lokad.OoxPdf.dll"
$firstLibrary = Join-Path $repo ('artifacts/first/' + $libraryRelative)
$firstCliLibrary = Join-Path $repo ('artifacts/first/' + $cliLibraryRelative)
$secondLibrary = Join-Path $repo ('artifacts/second/' + $libraryRelative)
$libraryBytes = [IO.File]::ReadAllBytes($firstLibrary)
$libraryHash = (Get-FileHash -LiteralPath $firstLibrary).Hash
# Deliberately mutate only fake, newly created assets to prove sharing and
# ownership isolation. Historical artifacts and real library outputs are unused.
[IO.File]::WriteAllText($firstCliLibrary, 'changed shared fake asset')
if ((Get-Content -LiteralPath $firstLibrary -Raw) -ne 'changed shared fake asset' -or
    (Get-FileHash -LiteralPath $secondLibrary).Hash -ne $libraryHash -or
    (Get-FileHash -LiteralPath (Join-Path $repo $libraryRelative)).Hash -ne $libraryHash) {
    throw 'Frozen asset sharing escaped its own output'
}
[IO.File]::WriteAllBytes($firstLibrary, $libraryBytes)
[IO.File]::WriteAllText((Join-Path $repo $libraryRelative), 'changed writable fake build')
if ((Get-FileHash -LiteralPath $firstLibrary).Hash -ne $libraryHash) { throw 'A frozen asset aliases a writable build' }
[IO.File]::WriteAllBytes((Join-Path $repo $libraryRelative), $libraryBytes)
# Exercise the real transfer's copies with a fake conversion; no library executes.
$qualification = Join-Path $OutputDirectory 'qualification.json'
@{LibrarySha256=$first.LibrarySha256;TestAssemblySha256=$first.TestAssemblySha256} |
    ConvertTo-Json | Set-Content -LiteralPath $qualification
$smoke = Join-Path $OutputDirectory 'smoke.json'
$fakePdf = Join-Path $OutputDirectory 'expected-fake.pdf'
[IO.File]::WriteAllText([IO.Path]::GetFullPath($fakePdf), 'fake transfer smoke')
@(@{Input=$inputFile;InputSha256=(Get-FileHash -LiteralPath $inputFile).Hash;
    PdfSha256=(Get-FileHash -LiteralPath $fakePdf).Hash;Arguments=@()}) |
    ConvertTo-Json -AsArray -Depth 4 | Set-Content -LiteralPath $smoke
function dotnet {
    [IO.File]::WriteAllText([IO.Path]::GetFullPath([string]$args[3]), 'fake transfer smoke')
    $global:LASTEXITCODE = 0
}
try {
    & $transferScript -ArtifactDirectory (Join-Path $repo 'artifacts/first') -QualificationEvidence $qualification `
        -SmokeManifest $smoke -OutputDirectory (Join-Path $OutputDirectory 'transfer') | Out-Null
}
finally { Remove-Item Function:\dotnet }
[IO.File]::WriteAllText((Join-Path $repo $cliLibraryRelative), 'changed transferred fake asset')
if ((Get-FileHash -LiteralPath $firstLibrary).Hash -ne $libraryHash -or
    (Get-FileHash -LiteralPath (Join-Path $repo $libraryRelative)).Hash -ne $libraryHash) {
    throw 'Transferred outputs alias frozen assets or each other'
}
[IO.File]::WriteAllBytes((Join-Path $repo $cliLibraryRelative), $libraryBytes)
function Reject-Transfer([string] $Name) {
    $destination = Join-Path $OutputDirectory $Name
    $rejected = $false
    try {
        & $transferScript -ArtifactDirectory (Join-Path $repo 'artifacts/first') -QualificationEvidence $qualification `
            -SmokeManifest $smoke -OutputDirectory $destination | Out-Null
    }
    catch {
        if (!$_.Exception.Message.Contains('build-setting transfer identity differs')) { throw }
        $rejected = $true
    }
    if (!$rejected -or (Test-Path -LiteralPath $destination) -or
        (Get-FileHash -LiteralPath (Join-Path $repo $libraryRelative)).Hash -ne $libraryHash) {
        throw 'Build-setting refusal must precede transfer writes'
    }
}
'<Project><PropertyGroup><CheckForOverflowUnderflow>true</CheckForOverflowUnderflow></PropertyGroup></Project>' |
    Set-Content -LiteralPath (Join-Path $repo 'Directory.Build.props')
Reject-Transfer 'dirty-build-setting'
Invoke-Git @('add', 'Directory.Build.props')
Invoke-Git @('commit', '--quiet', '-m', 'Changed production build setting')
Reject-Transfer 'committed-build-setting'
'<Project />' | Set-Content -LiteralPath (Join-Path $repo 'Directory.Build.props')
Invoke-Git @('add', 'Directory.Build.props')
Invoke-Git @('commit', '--quiet', '-m', 'Restore production build setting')
'<Project />' | Set-Content -LiteralPath (Join-Path $repo 'Directory.Build.targets')
Reject-Transfer 'new-root-targets'
Remove-Item -LiteralPath (Join-Path $repo 'Directory.Build.targets')
'// changed owned inspector source' | Set-Content -LiteralPath $inspectorSource
Reject-Transfer 'dirty-inspector-source'
Invoke-Git @('add', 'tools/Lokad.OoxPdf.DocxInspect/Program.cs')
Invoke-Git @('commit', '--quiet', '-m', 'Changed frozen inspector source')
Reject-Transfer 'committed-inspector-source'
'// owned inspector source' | Set-Content -LiteralPath $inspectorSource
Invoke-Git @('add', 'tools/Lokad.OoxPdf.DocxInspect/Program.cs')
Invoke-Git @('commit', '--quiet', '-m', 'Restore frozen inspector source')
$toolsTargets = Join-Path $repo 'tools/Directory.Build.targets'
'<Project />' | Set-Content -LiteralPath $toolsTargets
Reject-Transfer 'new-tools-targets'
Invoke-Git @('add', 'tools/Directory.Build.targets')
Invoke-Git @('commit', '--quiet', '-m', 'Add inspector parent targets')
Reject-Transfer 'committed-tools-targets'
Invoke-Git @('rm', '--quiet', 'tools/Directory.Build.targets')
Invoke-Git @('commit', '--quiet', '-m', 'Restore inspector parent targets')
# Orchestration scripts do not compile into the frozen inspector. Permitted
# harness repairs must not force another library build or qualification suite.
$harness = Join-Path $repo 'tools/owned-harness.ps1'
'# owned harness change' | Set-Content -LiteralPath $harness
function dotnet {
    [IO.File]::WriteAllText([IO.Path]::GetFullPath([string]$args[3]), 'fake transfer smoke')
    $global:LASTEXITCODE = 0
}
try {
    & $transferScript -ArtifactDirectory (Join-Path $repo 'artifacts/first') -QualificationEvidence $qualification -SmokeManifest $smoke -OutputDirectory (Join-Path $OutputDirectory 'harness-repair-transfer') | Out-Null
}
finally { Remove-Item Function:\dotnet }
Remove-Item -LiteralPath $harness
if ($first.ReusedInputSnapshot -or !$second.ReusedInputSnapshot -or
    $first.InputSnapshot -ne $second.InputSnapshot -or $first.InputTree -ne $second.InputTree) { throw 'Unchanged input reuse failed' }
$link = Get-Item -LiteralPath (Join-Path $repo ('artifacts/second/' + (Split-Path $inputPath)))
if (!$link.LinkType -or (Get-Content -LiteralPath (Join-Path $link.FullName 'sample.txt') -Raw) -notmatch 'first fixture') { throw 'Frozen input link is unusable' }
# Mutating a parent fixture must never mutate the frozen source, including
# its read-only metadata. Reuse replaces owned entries rather than aliases.
$parentManifest = Join-Path $OutputDirectory 'parent-inputs.json'
@('sample.txt','nested\second.txt') | ConvertTo-Json | Set-Content -LiteralPath $parentManifest
$parentOutput = Join-Path $OutputDirectory 'parent-runtime'
$parentArguments = @{TestArtifactDirectory=(Join-Path $repo 'artifacts/first'); ParentArtifactDirectory=(Join-Path $repo 'artifacts/second'); InputManifest=$parentManifest; OutputDirectory=$parentOutput}
& $parentScript @parentArguments | Out-Null
$frozenInput = Join-Path $first.InputSnapshot $inputPath
$frozenHash = (Get-FileHash -LiteralPath $frozenInput).Hash
$ownedInput = Join-Path $parentOutput $inputPath
[IO.File]::WriteAllText([IO.Path]::GetFullPath($ownedInput), 'mutable parent fixture')
if ((Get-FileHash -LiteralPath $frozenInput).Hash -ne $frozenHash -or !(Get-Item -LiteralPath $frozenInput).IsReadOnly) { throw 'Parent fixture aliases a frozen input' }
& $parentScript @parentArguments | Out-Null
if ((Get-FileHash -LiteralPath $ownedInput).Hash -ne $frozenHash) { throw 'Owned parent reuse failed' }
$parentArguments.OutputDirectory = Join-Path $repo 'artifacts/second'
$parentRejected = $false
try { & $parentScript @parentArguments | Out-Null }
catch { if (!$_.Exception.Message.Contains('unowned existing')) { throw }; $parentRejected = $true }
if (!$parentRejected -or (Get-FileHash -LiteralPath $frozenInput).Hash -ne $frozenHash) { throw 'Frozen parent output was not rejected without writes' }
$worktree = Join-Path $repo 'artifacts/worktree'
Invoke-Git @('config', 'core.autocrlf', 'true')
Invoke-Git @('worktree', 'add', '--quiet', '--detach', $worktree, 'HEAD')
if ((Get-FileHash -LiteralPath (Join-Path $worktree $inputPath)).Hash -ne (Get-FileHash -LiteralPath $inputFile).Hash) {
    throw 'Verbatim public fixtures changed under autocrlf checkout'
}
foreach ($project in @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests', 'tools/Lokad.OoxPdf.DocxInspect')) {
    $relative = $project + "/bin/Release/$Framework"
    New-Item -ItemType Directory -Path (Join-Path $worktree $relative) -Force | Out-Null
    foreach ($asset in Get-ChildItem -LiteralPath (Join-Path $repo $relative) -File) {
        Copy-Item -LiteralPath $asset.FullName -Destination (Join-Path $worktree $relative)
    }
}
$worktreeOutput = Join-Path $repo 'artifacts/from-worktree'
& (Join-Path $worktree 'tools/FreezeQualificationArtifact.ps1') -OutputDirectory $worktreeOutput -Framework $Framework | Out-Null
$worktreePin = Get-Content -LiteralPath (Join-Path $worktreeOutput 'identity.json') -Raw | ConvertFrom-Json
if (!$worktreePin.ReusedInputSnapshot -or $worktreePin.InputSnapshot -ne $first.InputSnapshot) { throw 'Linked worktree reuse failed' }
$reuseBuildScript=Join-Path $worktree 'tools/BuildQualificationArtifact.ps1'
$reuseBuildRevision=git -C $worktree rev-parse HEAD
$reuseBuildState=[pscustomobject]@{Calls=0;Damage=$false}
function dotnet {
    $buildArguments=@($args)
    if ($buildArguments[0] -ne 'build' -or $buildArguments[1] -ne (Join-Path $worktree 'tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj') -or
        '-p:BuildProjectReferences=false' -notin $buildArguments) {
        throw 'Harness reuse attempted an unexpected build'
    }
    $reuseBuildState.Calls++
    'corrected harness; no executable code' | Set-Content -LiteralPath (Join-Path $worktree "tests/Lokad.OoxPdf.Tests/bin/Release/$Framework/Lokad.OoxPdf.Tests.dll")
    if ($reuseBuildState.Damage) {
        'unwanted runtime rewrite' | Set-Content -LiteralPath (Join-Path $worktree "src/Lokad.OoxPdf/bin/Release/$Framework/Lokad.OoxPdf.dll")
    }
    $global:LASTEXITCODE=0
}
function Reject-ReusedBuild([string]$Name,[string]$Expected,[string]$SelectedFramework=$Framework) {
    $target=Join-Path $repo ('artifacts/'+$Name)
    try {
        & $reuseBuildScript -WorktreeDirectory $worktree -Revision $reuseBuildRevision -OutputDirectory $target -Framework $SelectedFramework -ReuseRuntimeDirectory $worktreeOutput | Out-Null
    } catch {
        if (!$_.Exception.Message.Contains($Expected) -or (Test-Path -LiteralPath $target)) { throw }
        return
    }
    throw ('Expected reused-runtime build rejection: '+$Name)
}
try {
    $repairedOutput=Join-Path $repo 'artifacts/repaired-harness'
    & $reuseBuildScript -WorktreeDirectory $worktree -Revision $reuseBuildRevision -OutputDirectory $repairedOutput -Framework $Framework -ReuseRuntimeDirectory $worktreeOutput | Out-Null
    $repaired=Get-Content (Join-Path $repairedOutput 'identity.json') -Raw | ConvertFrom-Json
    if ($reuseBuildState.Calls -ne 1 -or $repaired.LibrarySha256 -ne $worktreePin.LibrarySha256 -or $repaired.TestAssemblySha256 -eq $worktreePin.TestAssemblySha256) { throw 'Harness-only build did not preserve its runtime' }
    $reuseBuildState.Damage=$true
    Reject-ReusedBuild 'harness-rewrites-runtime' 'Harness build changed a qualified runtime asset'
    $reuseBuildState.Damage=$false
    Reject-ReusedBuild 'harness-wrong-framework' 'Reused runtime framework differs' $(if ($Framework -eq 'net11.0') {'net10.0'} else {'net11.0'})
    $beforeRefusal=$reuseBuildState.Calls
    '// changed production' | Set-Content (Join-Path $worktree 'src/Lokad.OoxPdf/changed.cs')
    git -C $worktree add src/Lokad.OoxPdf/changed.cs
    git -C $worktree commit --quiet -m 'Changed synthetic production'
    $reuseBuildRevision=git -C $worktree rev-parse HEAD
    Reject-ReusedBuild 'harness-changed-production' 'Production source or build inputs changed'
    if ($reuseBuildState.Calls -ne $beforeRefusal) { throw 'Changed production reached the harness compiler' }
} finally { Remove-Item Function:\dotnet }
'second fixture' | Set-Content -LiteralPath $inputFile
Reject 'dirty' 'Commit source and public inputs'
Invoke-Git @('add', $inputPath)
Invoke-Git @('commit', '--quiet', '-m', 'Changed fixture')
$changed = Freeze 'changed'
if ($changed.ReusedInputSnapshot -or $changed.InputTree -eq $first.InputTree -or
    (Get-Content -LiteralPath (Join-Path $first.InputSnapshot $inputPath) -Raw) -notmatch 'first fixture') { throw 'Changed input isolation failed' }
$firstNested = Join-Path $first.InputSnapshot 'tests/Lokad.OoxPdf.Tests/Cases/nested/second.txt'
$changedNested = Join-Path $changed.InputSnapshot 'tests/Lokad.OoxPdf.Tests/Cases/nested/second.txt'
if ($changed.SharedInputFiles -ne 0) { throw 'Different input trees must own their copies' }
(Get-Item $changedNested).IsReadOnly = $false
[IO.File]::WriteAllText($changedNested, 'corrupt independent copy')
if ((Get-Content $firstNested -Raw) -notmatch 'nested fixture' -or
    !(Get-Item $firstNested).IsReadOnly) { throw 'Different input trees share mutable ownership' }
(Get-Item $changedNested).IsReadOnly = $true
Reject 'corrupt-copy' 'Shared public input differs'
# A newer tree uses the verified checkout, not a corrupted older snapshot.
(Get-Item $inputFile).IsReadOnly = $false
'new tree fixture' | Set-Content -LiteralPath $inputFile
Invoke-Git @('add', $inputPath)
Invoke-Git @('commit', '--quiet', '-m', 'Third fixture tree')
$repaired = Freeze 'repaired'
if ((Get-FileHash -LiteralPath (Join-Path $repaired.InputSnapshot 'tests/Lokad.OoxPdf.Tests/Cases/nested/second.txt')).Hash -ne
    (Get-FileHash -LiteralPath $nested).Hash) { throw 'New snapshot propagated old corruption' }

# Fail after one independent copy has become read-only. Only this call's staging
# may disappear; published input bytes and protection must survive.
'failure fixture' | Set-Content -LiteralPath $inputFile
Invoke-Git @('add', $inputPath)
Invoke-Git @('commit', '--quiet', '-m', 'Failure fixture tree')
$script:copiedInputs = 0
function Copy-Item {
    param([string] $LiteralPath, [string] $Destination)
    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination
    if ($Destination -match '[.]staging-' -and ++$script:copiedInputs -eq 2) { throw 'Injected input copy failure' }
}
try { Reject 'copy-failure' 'Injected input copy failure' }
finally { Remove-Item Function:\Copy-Item }
$inputSnapshots = Join-Path $repo 'artifacts/qualification-inputs'
if (@(Get-ChildItem -LiteralPath $inputSnapshots -Directory -Filter '*.staging-*').Count) { throw 'Failed copy left input staging' }
if ((Get-FileHash -LiteralPath $firstNested).Hash -ne (Get-FileHash -LiteralPath $nested).Hash -or
    !(Get-Item $firstNested).IsReadOnly) { throw 'Failure cleanup changed published input ownership' }

# Publish a complete competing copy just before Directory.Move. Cleanup must
# remove the losing staging directory and preserve the winner's read-only files.
function Set-Content {
    param([Parameter(ValueFromPipeline=$true)] $Value, [string] $LiteralPath, [string] $Encoding)
    process {
        Microsoft.PowerShell.Management\Set-Content -LiteralPath $LiteralPath -Value $Value -Encoding $Encoding
        if ($LiteralPath -match '[.]staging-[0-9a-f]+[\\/]identity[.]json$') {
            $staged = Split-Path $LiteralPath
            $winner = $staged.Substring(0, $staged.LastIndexOf('.staging-', [StringComparison]::Ordinal))
            # Write independent bytes before applying protection. Recursive
            # Copy-Item can create an unwritable destination on WSL/NTFS when
            # the synthetic winner's source is already read-only.
            foreach ($file in Get-ChildItem -LiteralPath $staged -File -Recurse) {
                $target = Join-Path $winner ([IO.Path]::GetRelativePath($staged, $file.FullName))
                New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
                [IO.File]::WriteAllBytes($target, [IO.File]::ReadAllBytes($file.FullName))
                (Get-Item -LiteralPath $target).IsReadOnly = $file.IsReadOnly
            }
        }
    }
}
try { $raced = Freeze 'race' }
finally { Remove-Item Function:\Set-Content }
if (@(Get-ChildItem -LiteralPath $inputSnapshots -Directory -Filter '*.staging-*').Count) { throw 'Losing publication left input staging' }
foreach ($entry in $raced.Inputs.psobject.Properties) {
    $published = Join-Path $raced.InputSnapshot $entry.Name
    if ((Get-FileHash -LiteralPath $published).Hash -ne $entry.Value -or !(Get-Item $published).IsReadOnly) {
        throw 'Publication race changed its winning input'
    }
}
$damaged = Join-Path $raced.InputSnapshot $inputPath
(Get-Item -LiteralPath $damaged).IsReadOnly = $false
'corrupted shared copy' | Set-Content -LiteralPath $damaged
Reject 'corrupt' 'Shared public input differs'
Copy-Item -LiteralPath $inputFile -Destination $damaged -Force
(Get-Item -LiteralPath $damaged).IsReadOnly = $true
'stale CLI library' | Set-Content -LiteralPath (Join-Path $repo "src/Lokad.OoxPdf.Cli/bin/Release/$Framework/Lokad.OoxPdf.dll")
Reject 'mixed' 'Release library copies differ'

# The normal freezer rejects coherent binaries from an older production
# revision, even when every library copy agrees and fixture bytes match.
Copy-Item -LiteralPath $assemblyIdentitySource -Destination $assemblyIdentityFixture -Force
Invoke-Git @('add', 'tools/QualificationBuildIdentity.ps1')
Invoke-Git @('commit', '--quiet', '-m', 'Use real qualification assembly metadata reader')
$compiledRevision = git -C $repo rev-parse HEAD
function New-StampedAssembly([string] $Revision) {
    $path = Join-Path $OutputDirectory ('revision-'+$Revision+'.dll')
    Add-Type -TypeDefinition @"
using System.Reflection;
[assembly: AssemblyInformationalVersion("1.0.0+$Revision")]
public sealed class Fixture_$Revision { }
"@ -OutputAssembly $path
    return [IO.File]::ReadAllBytes($path)
}
$compiledBytes = New-StampedAssembly $compiledRevision
$harness = "tests/Lokad.OoxPdf.Tests/bin/Release/$Framework/Lokad.OoxPdf.Tests.dll"
foreach ($project in @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests', 'tools/Lokad.OoxPdf.DocxInspect')) {
    $relative = "$project/bin/Release/$Framework"
    [IO.File]::WriteAllBytes((Join-Path $repo ($relative+'/Lokad.OoxPdf.dll')), $compiledBytes)
    [IO.File]::WriteAllBytes((Join-Path $repo ($relative+'/'+(Split-Path $project -Leaf)+'.dll')), $compiledBytes)
}
$fresh = Freeze 'compiled-source'
if ($fresh.Revision -ne $compiledRevision -or @($fresh.AssemblyRevisions.PSObject.Properties).Count -ne 4 -or
    @($fresh.AssemblyRevisions.PSObject.Properties.Value | Where-Object { $_ -ne $compiledRevision }).Count) {
    throw 'Fresh assembly identity lost the embedded revisions'
}
$production = 'src/Lokad.OoxPdf/ChangedSource.cs'
'// changed production input' | Set-Content -LiteralPath (Join-Path $repo $production)
Invoke-Git @('add', $production)
Invoke-Git @('commit', '--quiet', '-m', 'Change production without rebuilding')
Reject 'stale-coherent-assemblies' 'Production source or build inputs changed'
Invoke-Git @('rm', '--quiet', $production)
Invoke-Git @('commit', '--quiet', '-m', 'Restore the compiled production inputs')
Reject 'stale-test-assembly' 'Release test assembly was compiled for another source revision'
$testRevision = git -C $repo rev-parse HEAD
[IO.File]::WriteAllBytes((Join-Path $repo $harness), (New-StampedAssembly $testRevision))
$currentHarness = Freeze 'current-harness-with-reused-production'
if ($currentHarness.Revision -ne $testRevision -or $currentHarness.AssemblyRevisions.$harness -ne $testRevision -or
    $currentHarness.LibrarySha256 -ne $fresh.LibrarySha256 -or $currentHarness.TestAssemblySha256 -eq $fresh.TestAssemblySha256 -or
    @($currentHarness.AssemblyRevisions.PSObject.Properties | Where-Object { $_.Name -ne $harness -and $_.Value -eq $compiledRevision }).Count -ne 3) {
    throw 'Harness-only publication misreported reused production provenance'
}
[ordered]@{
    Checks=@('Shared unchanged fixture tree', 'Frozen runner input link', 'Linked worktree reuse', 'Verbatim autocrlf checkout', 'Dirty source rejection',
        'Canonical LF and committed CRLF with spaced path', 'Noncanonical refusal before snapshot',
        'Noncanonical reuse refusal', 'Same-length noncanonical refusal', 'Retained legacy snapshot isolation',
        'Changed fixture isolation', 'Cross-tree ownership isolation', 'Corrupt copy rejection',
        'Independent prior-tree corruption isolation', 'Failed staging cleanup', 'Concurrent publication cleanup',
        'Corrupt snapshot rejection', 'Mixed library rejection',
        'Owned parent input isolation', 'Owned parent reuse', 'Frozen parent destination rejection',
        'Duplicate asset sharing', 'Separate frozen output isolation', 'Writable build isolation',
        'Independent transfer copies', 'Dirty root build-setting refusal', 'Committed root build-setting refusal', 'New root targets refusal',
        'Dirty inspector source refusal', 'Committed inspector source refusal',
        'New tools targets refusal', 'Committed tools targets refusal', 'Harness-only repair transfer',
        'Harness-only runtime reuse', 'Runtime rewrite refusal', 'Reused framework refusal', 'Reused production-change refusal',
        'Dual-framework fixture verification once', 'All framework libraries checked before batch publication',
        'Real embedded revision', 'Stale coherent production refusal', 'Stale harness refusal', 'Reused production revision retained')
    LibraryBuilds=0;LibraryTestRuns=0;First=$first;Second=$second;Changed=$changed
    CompiledSource=$fresh.AssemblyRevisions;ReusedProduction=$currentHarness.AssemblyRevisions
    Batch=$batchPins;CheckoutAndSnapshotHashesPerFixture=2
    SharedInputPathLength=$changedNested.Length
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
Write-Output 'Fixture-freeze/asset/parent/harness-ownership checks passed; no library build or test-suite replay.'
