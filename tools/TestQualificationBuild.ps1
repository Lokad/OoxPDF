# Exercise workspace/build publication guards with a tiny Git repo and fake builds.
[CmdletBinding()]
param([string] $OutputDirectory,
    [string] $BuildScript=(Join-Path $PSScriptRoot 'BuildQualificationArtifact.ps1'))
$ErrorActionPreference='Stop'
if (!$OutputDirectory) {
    $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/build-guard-test-'+[Guid]::NewGuid().ToString('N'))
}
$output=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Use a fresh test directory' }
$repo=Join-Path $output 'repository'
$worktree=Join-Path $output 'isolated worktree'
New-Item -ItemType Directory -Path $repo -Force | Out-Null
git -C $repo init --quiet
git -C $repo config user.name 'Qualification test'
git -C $repo config user.email 'qualification@example.invalid'
@('artifacts/','**/bin/') | Set-Content -LiteralPath (Join-Path $repo '.gitignore')
'Committed source marker' | Set-Content -LiteralPath (Join-Path $repo 'marker.txt')
$control='tests/Lokad.OoxPdf.Tests/Cases/control.txt'
New-Item -ItemType Directory -Path (Split-Path (Join-Path $repo $control)) -Force | Out-Null
'Qualified input' | Set-Content -LiteralPath (Join-Path $repo $control)
foreach ($relative in @('tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj','src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj','tools/Lokad.OoxPdf.DocxInspect/Lokad.OoxPdf.DocxInspect.csproj')) {
    $project=Join-Path $repo $relative
    New-Item -ItemType Directory -Path (Split-Path $project) -Force | Out-Null
    '<Project />' | Set-Content -LiteralPath $project
}
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
@'
param($OutputDirectory,$Framework)
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$repo=Split-Path -Parent $PSScriptRoot
@{Revision=(git -C $repo rev-parse HEAD);LibrarySha256=('1'*64);Fake=$true} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'identity.json')
'@ | Set-Content -LiteralPath (Join-Path $repo 'tools/FreezeQualificationArtifact.ps1')
git -C $repo add .
git -C $repo commit --quiet -m 'Fixture for build workspace guards'
if ($LASTEXITCODE -ne 0) { throw 'Cannot commit test fixture' }
$revision=git -C $repo rev-parse HEAD
git -C $repo worktree add --quiet -b qualification-test $worktree
if ($LASTEXITCODE -ne 0) { throw 'Cannot create test worktree' }
$global:ooxPdfBuildGuardCalls=@()
$global:ooxPdfBuildGuardFault=''
$global:ooxPdfBuildGuardWorktree=$worktree
function dotnet {
    $global:ooxPdfBuildGuardCalls += ,@($args)
    if ($global:ooxPdfBuildGuardFault -eq 'source-change') {
        'Changed during build' | Set-Content -LiteralPath (Join-Path $global:ooxPdfBuildGuardWorktree 'marker.txt')
    }
    $global:LASTEXITCODE=if ($global:ooxPdfBuildGuardFault -eq 'build-failure') {17} else {0}
}
$passed=@()
try {
    foreach ($fault in @('main-checkout','subdirectory','revision','dirty-source','relative-path','outside-artifacts','existing-output')) {
        $destination=Join-Path $repo ('artifacts/'+$fault)
        $selected=$worktree; $pin=$revision
        switch ($fault) {
            'main-checkout' {$selected=$repo}
            'subdirectory' {$selected=Join-Path $worktree 'src'}
            'revision' {$pin='0'*40}
            'dirty-source' {'Uncommitted fixture edit' | Set-Content -LiteralPath (Join-Path $worktree 'marker.txt')}
            'relative-path' {$selected='.'}
            'outside-artifacts' {$destination=Join-Path $output 'outside'}
            'existing-output' {New-Item -ItemType Directory -Path $destination -Force | Out-Null}
        }
        $before=$global:ooxPdfBuildGuardCalls.Count
        $rejected=$false
        try { & $BuildScript -WorktreeDirectory $selected -Revision $pin -OutputDirectory $destination | Out-Null }
        catch {
            $expected=switch ($fault) {
                'main-checkout' {'Main checkout rejected'}
                'subdirectory' {'linked worktree root explicitly'}
                'revision' {'source revision differs'}
                'dirty-source' {'Commit source and public inputs'}
                'relative-path' {'Use absolute'}
                'outside-artifacts' {'shared artifacts directory'}
                'existing-output' {'Use fresh artifact'}
            }
            if (!$_.Exception.Message.Contains($expected)) { throw }
            $rejected=$true
        }
        if (!$rejected -or $global:ooxPdfBuildGuardCalls.Count -ne $before -or
            (Test-Path -LiteralPath ($destination+'.build'))) { throw ('Guard ran a build or created logs: '+$fault) }
        if ($fault -eq 'dirty-source') { git -C $worktree restore -- marker.txt }
        $passed+=$fault
    }
    foreach ($fault in @('build-failure','source-change')) {
        $destination=Join-Path $repo ('artifacts/'+$fault)
        $global:ooxPdfBuildGuardFault=$fault
        $before=$global:ooxPdfBuildGuardCalls.Count
        $rejected=$false
        try { & $BuildScript -WorktreeDirectory $worktree -Revision $revision -OutputDirectory $destination | Out-Null }
        catch {
            $expected=if ($fault -eq 'build-failure') {'Qualification build failed'} else {'Commit source and public inputs'}
            if (!$_.Exception.Message.Contains($expected)) { throw }
            $rejected=$true
        }
        if (!$rejected -or $global:ooxPdfBuildGuardCalls.Count -ne $before+1 -or
            (Test-Path -LiteralPath $destination)) { throw ('Failed build/source change published or continued: '+$fault) }
        if ($fault -eq 'source-change') { git -C $worktree restore -- marker.txt }
        $passed+=$fault
    }
    $global:ooxPdfBuildGuardFault=''
    $destination=Join-Path $repo 'artifacts/accepted'
    $before=$global:ooxPdfBuildGuardCalls.Count
    # Invoke from the main checkout; only the explicit linked worktree may build.
    Push-Location $repo
    try { & $BuildScript -WorktreeDirectory $worktree -Revision $revision -OutputDirectory $destination | Out-Null }
    finally { Pop-Location }
    $builds=@($global:ooxPdfBuildGuardCalls | Select-Object -Skip $before)
    if ($builds.Count -ne 3) { throw 'Expected test/CLI/DOCX inspector builds once each' }
    foreach ($arguments in $builds) {
        if ($arguments[0] -ne 'build' -or ![IO.Path]::IsPathFullyQualified($arguments[1]) -or
            !$arguments[1].StartsWith($worktree+[IO.Path]::DirectorySeparatorChar) -or
            $arguments[2] -ne '-c' -or $arguments[3] -ne 'Release' -or
            $arguments[4] -ne '-f' -or $arguments[5] -ne 'net11.0') { throw 'Build selected a relative/wrong project or configuration' }
    }
    $identity=Get-Content -LiteralPath (Join-Path $destination 'identity.json') -Raw | ConvertFrom-Json
    $evidence=Get-Content -LiteralPath ($destination+'.build/evidence.json') -Raw | ConvertFrom-Json
    if (!$identity.Fake -or $identity.Revision -ne $revision -or $evidence.TestsRun -ne 0) { throw 'Publication identity differs' }
    $passed+='absolute-projects-and-publication'

    # An extended harness uses the same qualified production bytes. Added
    # fixtures are a new complete snapshot; old fixtures cannot change/disappear.
    $seed=Join-Path $repo 'artifacts/reusable-runtime'
    $asset='src/Lokad.OoxPdf/bin/Release/net11.0/Lokad.OoxPdf.dll'
    New-Item -ItemType Directory -Path (Split-Path (Join-Path $seed $asset)) -Force | Out-Null
    'Qualified production bytes' | Set-Content -LiteralPath (Join-Path $seed $asset)
    $assetHash=(Get-FileHash -LiteralPath (Join-Path $seed $asset)).Hash
    @{Revision=$revision;Framework='net11.0';InputTree=(git -C $worktree rev-parse "${revision}:tests/Lokad.OoxPdf.Tests/Cases");
        Assets=@{$asset=$assetHash}} | ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath (Join-Path $seed 'identity.json')
    'Additional public coverage' | Set-Content -LiteralPath (Join-Path $worktree 'tests/Lokad.OoxPdf.Tests/Cases/added.txt')
    git -C $worktree add -- tests/Lokad.OoxPdf.Tests/Cases/added.txt
    git -C $worktree commit --quiet -m 'Add an independent public input'
    if ($LASTEXITCODE -ne 0) { throw 'Cannot commit added test input' }
    $extended=git -C $worktree rev-parse HEAD
    $destination=Join-Path $repo 'artifacts/extended-harness'
    $before=$global:ooxPdfBuildGuardCalls.Count
    & $BuildScript -WorktreeDirectory $worktree -Revision $extended -OutputDirectory $destination -ReuseRuntimeDirectory $seed | Out-Null
    $evidence=Get-Content -LiteralPath ($destination+'.build/evidence.json') -Raw | ConvertFrom-Json
    $arguments=$global:ooxPdfBuildGuardCalls[-1]
    if ($global:ooxPdfBuildGuardCalls.Count -ne $before+1 -or
        $arguments[1] -ne (Join-Path $worktree 'tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj') -or
        '-p:BuildProjectReferences=false' -notin $arguments -or
        $evidence.ReusedRuntimeAssets -ne 1 -or $evidence.AddedPublicInputs -ne 1 -or
        (Get-FileHash -LiteralPath (Join-Path $worktree $asset)).Hash -ne $assetHash) {
        throw 'Extended harness rebuilt or changed production, or lost its added-input evidence'
    }
    $passed+='added-input-reuses-runtime'
    foreach ($fault in @('changed-public-input','removed-public-input')) {
        $inputPath=Join-Path $worktree $control
        if ($fault -eq 'changed-public-input') { 'Changed qualified input' | Set-Content -LiteralPath $inputPath }
        else { Remove-Item -LiteralPath $inputPath }
        git -C $worktree add -- $control
        git -C $worktree commit --quiet -m $fault
        if ($LASTEXITCODE -ne 0) { throw 'Cannot commit input guard fixture' }
        $changedRevision=git -C $worktree rev-parse HEAD
        $destination=Join-Path $repo ('artifacts/'+$fault)
        $before=$global:ooxPdfBuildGuardCalls.Count
        $rejected=$false
        try { & $BuildScript -WorktreeDirectory $worktree -Revision $changedRevision -OutputDirectory $destination -ReuseRuntimeDirectory $seed | Out-Null }
        catch {
            if (!$_.Exception.Message.Contains('Reused public inputs changed or were removed')) { throw }
            $rejected=$true
        }
        if (!$rejected -or $global:ooxPdfBuildGuardCalls.Count -ne $before -or
            (Test-Path -LiteralPath ($destination+'.build'))) { throw ('Changed inputs reached a build: '+$fault) }
        $passed+=$fault
    }
    @{Passed=$passed;NativeBuilds=0;LibraryTests=0;Scope='Tiny Git repository; fake dotnet and freeze commands'} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $output 'evidence.json')
    Write-Output ($passed.Count.ToString()+' qualification build guards passed; no native builds or library tests.')
}
finally {
    Remove-Variable -Name ooxPdfBuildGuardCalls,ooxPdfBuildGuardFault,ooxPdfBuildGuardWorktree -Scope Global
}
