# Build one committed linked worktree, then delegate immutable publication to Freeze.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $WorktreeDirectory,
    [Parameter(Mandatory=$true)] [ValidatePattern('^(?:[0-9a-fA-F]{40}|[0-9a-fA-F]{64})$')] [string] $Revision,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [string] $ReuseRuntimeDirectory
)
$ErrorActionPreference='Stop'
# Handle native exit codes explicitly so failures always point to retained logs.
$PSNativeCommandUseErrorActionPreference=$false
if (![IO.Path]::IsPathFullyQualified($WorktreeDirectory) -or ![IO.Path]::IsPathFullyQualified($OutputDirectory)) {
    throw 'Use absolute worktree and artifact paths'
}
$source=[IO.Path]::TrimEndingDirectorySeparator((Resolve-Path -LiteralPath $WorktreeDirectory).Path)
$output=[IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($OutputDirectory))
$comparison=if ($IsWindows) {[StringComparison]::OrdinalIgnoreCase} else {[StringComparison]::Ordinal}
function Read-Git([string[]] $Arguments) {
    $value=@(& git -C $source @Arguments)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot verify qualification worktree' }
    return $value
}
$top=[IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath((Read-Git -Arguments @('rev-parse','--show-toplevel'))))
if (!$source.Equals($top,$comparison)) { throw 'Select the linked worktree root explicitly' }
$gitDirectory=[IO.Path]::GetFullPath((Read-Git -Arguments @('rev-parse','--path-format=absolute','--git-dir')))
$commonDirectory=[IO.Path]::GetFullPath((Read-Git -Arguments @('rev-parse','--path-format=absolute','--git-common-dir')))
if ($gitDirectory.Equals($commonDirectory,$comparison)) { throw 'Main checkout rejected; qualification requires a linked worktree' }
$artifacts=[IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $commonDirectory) 'artifacts'))
if (!$output.StartsWith($artifacts+[IO.Path]::DirectorySeparatorChar,$comparison)) { throw 'Publish qualification outputs under the shared artifacts directory' }
$logs=$output+'.build'
if ((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $logs)) { throw 'Use fresh artifact and build-log directories' }
$projects=@('tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj','src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj','tools/Lokad.OoxPdf.DocxInspect/Lokad.OoxPdf.DocxInspect.csproj')
$freeze=Join-Path $source 'tools/FreezeQualificationArtifact.ps1'
foreach ($path in @($projects | ForEach-Object { Join-Path $source $_ }) + @($freeze)) {
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) { throw ('Qualification input is missing: '+$path) }
}
function Assert-Source {
    if ((Read-Git -Arguments @('rev-parse','HEAD')) -ne $Revision) { throw 'Qualification source revision differs' }
    if (@(Read-Git -Arguments @('status','--porcelain','--untracked-files=normal')).Count) { throw 'Commit source and public inputs before qualification builds' }
}
Assert-Source
$reuse=$null; $runtime=@()
if ($ReuseRuntimeDirectory) {
    $original=(Resolve-Path -LiteralPath $ReuseRuntimeDirectory).Path
    if ($output.Equals($original,$comparison) -or $output.StartsWith($original+[IO.Path]::DirectorySeparatorChar,$comparison)) {
        throw 'Publish the repaired harness outside its immutable runtime artifact'
    }
    $reuse=Get-Content -LiteralPath (Join-Path $original 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
    $originalFramework=if ($reuse.Framework) {$reuse.Framework} else {'net10.0'}
    if ($originalFramework -ne $Framework) { throw 'Reused runtime framework differs' }
    $buildInputs=@('src','tools/Lokad.OoxPdf.DocxInspect',':(top,glob)**/*.props',':(top,glob)**/*.targets',
        ':(top,glob)**/*.csproj','Directory.Build.rsp','tools/Directory.Build.rsp',
        'global.json','tools/global.json','NuGet.Config','nuget.config','tools/NuGet.Config','tools/nuget.config','.gitattributes')
    if (@(Read-Git -Arguments (@('diff','--name-only',$reuse.Revision,$Revision,'--')+$buildInputs)).Count) {
        throw 'Production source or build inputs changed; build a new runtime'
    }
    if ((Read-Git -Arguments @('rev-parse','HEAD:tests/Lokad.OoxPdf.Tests/Cases')) -ne $reuse.InputTree) {
        throw 'Reused public input tree differs'
    }
    foreach ($name in $reuse.Assets.Keys) {
        if ((Get-FileHash -LiteralPath (Join-Path $original $name)).Hash -ne $reuse.Assets[$name]) { throw 'Reused runtime asset differs' }
    }
    $runtime=@($reuse.Assets.Keys | Where-Object {$_ -notmatch '/Lokad\.OoxPdf\.Tests\.(dll|exe|pdb)$'})
    if (!$runtime.Count) { throw 'Reused runtime assets are empty' }
    # Replace entries atomically; never write through a pre-existing file link.
    foreach ($name in $runtime) {
        $target=[IO.Path]::GetFullPath((Join-Path $source $name))
        if (!$target.StartsWith($source+[IO.Path]::DirectorySeparatorChar,$comparison) -or
            $name -cnotmatch ('^(?:src/Lokad\.OoxPdf(?:\.Cli)?|tests/Lokad\.OoxPdf\.Tests|tools/Lokad\.OoxPdf\.DocxInspect)/bin/Release/'+[regex]::Escape($Framework)+'/[^/\\]+$')) {
            throw 'Reused asset target escaped the build outputs'
        }
        $directory=Split-Path $target
        while ($directory.Length -gt $source.Length) {
            if ((Test-Path -LiteralPath $directory) -and
                ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Reused build output directory is linked' }
            $directory=Split-Path $directory
        }
        New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
        $temporary=$target+'.staging-'+[Guid]::NewGuid().ToString('N')
        try {
            [IO.File]::Copy((Join-Path $original $name),$temporary,$false)
            if ((Get-FileHash -LiteralPath $temporary).Hash -ne $reuse.Assets[$name]) { throw 'Reused asset copy differs' }
            [IO.File]::Move($temporary,$target,$true)
        } finally {
            if ([IO.File]::Exists($temporary)) { Remove-Item -LiteralPath $temporary }
        }
    }
    $projects=@($projects[0])
}
New-Item -ItemType Directory -Path $logs -Force | Out-Null
$builds=@()
foreach ($relative in $projects) {
    $project=Join-Path $source $relative
    $log=Join-Path $logs ($builds.Count.ToString()+'.log')
    # Absolute project arguments make the caller's working directory irrelevant.
    $options=if ($reuse) {@('--no-restore','-p:BuildProjectReferences=false')} else {@()}
    & dotnet build $project -c Release -f $Framework --nologo -v minimal @options *> $log
    if ($LASTEXITCODE -ne 0) { throw ('Qualification build failed; inspect '+$log) }
    $builds += [ordered]@{Project=$project;Log=$log;Configuration='Release';Framework=$Framework}
    Assert-Source
}
foreach ($name in $runtime) {
    if ((Get-FileHash -LiteralPath (Join-Path $source $name)).Hash -ne $reuse.Assets[$name]) {
        throw 'Harness build changed a qualified runtime asset'
    }
}
& $freeze -OutputDirectory $output -Framework $Framework
$pin=Get-Content -LiteralPath (Join-Path $output 'identity.json') -Raw | ConvertFrom-Json
if ($pin.Revision -ne $Revision) { throw 'Frozen source revision differs' }
[ordered]@{Worktree=$source;Revision=$Revision;Builds=$builds;Artifact=$output;
    Framework=$Framework;LibrarySha256=$pin.LibrarySha256;TestsRun=0;
    ReusedRuntimeRevision=$(if ($reuse) {$reuse.Revision} else {$null});ReusedRuntimeAssets=$runtime.Count} | ConvertTo-Json -Depth 5 |
    Set-Content -LiteralPath (Join-Path $logs 'evidence.json') -Encoding utf8
if ($reuse) {
    Write-Output ('Reused '+$runtime.Count+' runtime assets; built/froze the corrected harness at '+$Revision+'; no library build or tests replayed.')
} else {
    Write-Output ('Built and froze '+$Revision+' from explicit linked worktree; no tests replayed.')
}
