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
function Invoke-Git([string[]] $Arguments) {
    $result = git -C $repo @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw ($result -join "`n") }
}
Invoke-Git @('init', '--quiet')
Invoke-Git @('config', 'user.name', 'Qualification test')
Invoke-Git @('config', 'user.email', 'qualification@example.invalid')
Invoke-Git @('config', 'core.autocrlf', 'false')
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
if ($first.ReusedInputSnapshot -or !$second.ReusedInputSnapshot -or
    $first.InputSnapshot -ne $second.InputSnapshot -or $first.InputTree -ne $second.InputTree) { throw 'Unchanged input reuse failed' }
$link = Get-Item -LiteralPath (Join-Path $repo ('artifacts/second/' + (Split-Path $inputPath)))
if (!$link.LinkType -or (Get-Content -LiteralPath (Join-Path $link.FullName 'sample.txt') -Raw) -notmatch 'first fixture') { throw 'Frozen input link is unusable' }
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
if ($changed.SharedInputFiles -ne 1 -or (Get-Item $changedNested).LinkType -ne 'HardLink') { throw 'Unchanged file sharing failed' }
# Deliberate corruption proves actual shared ownership, not just equal bytes.
$nestedBytes = [IO.File]::ReadAllBytes($firstNested)
(Get-Item $changedNested).IsReadOnly = $false
[IO.File]::WriteAllText($changedNested, 'corrupt alias')
if ((Get-Content $firstNested -Raw) -ne 'corrupt alias') { throw 'Fixture aliases are not shared' }
(Get-Item $changedNested).IsReadOnly = $true
Reject 'corrupt-alias' 'Shared public input differs'
# A different input tree must not propagate a corrupted prior alias.
(Get-Item $inputFile).IsReadOnly = $false
'new tree fixture' | Set-Content -LiteralPath $inputFile
Invoke-Git @('add', $inputPath)
Invoke-Git @('commit', '--quiet', '-m', 'Third fixture tree')
Reject 'corrupt-prior' 'Prior shared input differs'
(Get-Item $changedNested).IsReadOnly = $false
[IO.File]::WriteAllBytes($changedNested, $nestedBytes)
(Get-Item $changedNested).IsReadOnly = $true
$repaired = Freeze 'repaired'
if ($repaired.SharedInputFiles -ne 1) { throw 'Restored shared input was not reused' }
$damaged = Join-Path $repaired.InputSnapshot $inputPath
(Get-Item -LiteralPath $damaged).IsReadOnly = $false
'corrupted shared copy' | Set-Content -LiteralPath $damaged
Reject 'corrupt' 'Shared public input differs'
Copy-Item -LiteralPath $inputFile -Destination $damaged -Force
(Get-Item -LiteralPath $damaged).IsReadOnly = $true
'stale CLI library' | Set-Content -LiteralPath (Join-Path $repo 'src/Lokad.OoxPdf.Cli/bin/Release/net10.0/Lokad.OoxPdf.dll')
Reject 'mixed' 'Release library copies differ'
[ordered]@{
    Checks=@('Shared unchanged fixture tree', 'Frozen runner input link', 'Linked worktree reuse', 'Dirty source rejection',
        'Changed fixture isolation', 'Cross-tree immutable file sharing', 'Corrupt alias rejection',
        'Corrupt prior-tree rejection', 'Corrupt snapshot rejection', 'Mixed library rejection')
    LibraryBuilds=0;LibraryTestRuns=0;First=$first;Second=$second;Changed=$changed
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
Write-Output 'Ten fixture-freeze checks passed; no library build or test-suite replay.'
