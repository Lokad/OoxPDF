# Exercise fixture sharing and rejection paths without building or running the library.
[CmdletBinding()]
param([string] $OutputDirectory)
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
$inputPath = 'tests/Lokad.OoxPdf.Tests/Cases/sample.txt'
$inputFile = Join-Path $repo $inputPath
New-Item -ItemType Directory -Path (Split-Path $inputFile) -Force | Out-Null
'first fixture' | Set-Content -LiteralPath $inputFile
$nested = Join-Path (Split-Path $inputFile) 'nested/second.txt'
New-Item -ItemType Directory -Path (Split-Path $nested) -Force | Out-Null
'nested fixture' | Set-Content -LiteralPath $nested
foreach ($project in @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests')) {
    $directory = Join-Path $repo ($project + '/bin/Release/net10.0')
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    'same fake library; no executable code' | Set-Content -LiteralPath (Join-Path $directory 'Lokad.OoxPdf.dll')
    'fake test assembly' | Set-Content -LiteralPath (Join-Path $directory 'Lokad.OoxPdf.Tests.dll')
}
foreach ($name in @('Lokad.OoxPdf.Tests.pdb','Lokad.OoxPdf.Tests.deps.json','Lokad.OoxPdf.Tests.runtimeconfig.json')) {
    'fake runner support; no executable code' | Set-Content -LiteralPath (Join-Path $repo ('tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/' + $name))
}
Invoke-Git @('add', '.')
Invoke-Git @('commit', '--quiet', '-m', 'Initial fixtures')
function Freeze([string] $Name) {
    $target = Join-Path $repo ('artifacts/' + $Name)
    & $script -OutputDirectory $target | Out-Null
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
$first = Freeze 'first'
$second = Freeze 'second'
$expectedSharedAssets = @($first.Assets.psobject.Properties).Count - @($first.Assets.psobject.Properties.Value | Sort-Object -Unique).Count
if ($expectedSharedAssets -lt 1 -or $first.SharedAssetFiles -ne $expectedSharedAssets -or $first.SharedAssetBytes -lt 1) {
    throw 'Frozen duplicate assets were not shared'
}
$libraryRelative = 'src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll'
$cliLibraryRelative = 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.dll'
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
Invoke-Git @('worktree', 'add', '--quiet', '--detach', $worktree, 'HEAD')
foreach ($project in @('src/Lokad.OoxPdf', 'src/Lokad.OoxPdf.Cli', 'tests/Lokad.OoxPdf.Tests')) {
    $relative = $project + '/bin/Release/net10.0'
    New-Item -ItemType Directory -Path (Join-Path $worktree $relative) -Force | Out-Null
    foreach ($asset in Get-ChildItem -LiteralPath (Join-Path $repo $relative) -File) {
        Copy-Item -LiteralPath $asset.FullName -Destination (Join-Path $worktree $relative)
    }
}
$worktreeOutput = Join-Path $repo 'artifacts/from-worktree'
& (Join-Path $worktree 'tools/FreezeQualificationArtifact.ps1') -OutputDirectory $worktreeOutput | Out-Null
$worktreePin = Get-Content -LiteralPath (Join-Path $worktreeOutput 'identity.json') -Raw | ConvertFrom-Json
if (!$worktreePin.ReusedInputSnapshot -or $worktreePin.InputSnapshot -ne $first.InputSnapshot) { throw 'Linked worktree reuse failed' }
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
            Microsoft.PowerShell.Management\Copy-Item -LiteralPath $staged -Destination $winner -Recurse
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
'stale CLI library' | Set-Content -LiteralPath (Join-Path $repo 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.dll')
Reject 'mixed' 'Release library copies differ'
[ordered]@{
    Checks=@('Shared unchanged fixture tree', 'Frozen runner input link', 'Linked worktree reuse', 'Dirty source rejection',
        'Changed fixture isolation', 'Cross-tree ownership isolation', 'Corrupt copy rejection',
        'Independent prior-tree corruption isolation', 'Failed staging cleanup', 'Concurrent publication cleanup',
        'Corrupt snapshot rejection', 'Mixed library rejection',
        'Owned parent input isolation', 'Owned parent reuse', 'Frozen parent destination rejection',
        'Duplicate asset sharing', 'Separate frozen output isolation', 'Writable build isolation',
        'Independent transfer copies')
    LibraryBuilds=0;LibraryTestRuns=0;First=$first;Second=$second;Changed=$changed
    SharedInputPathLength=$changedNested.Length
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
Write-Output 'Nineteen fixture-freeze/asset/parent-ownership checks passed; no library build or test-suite replay.'
