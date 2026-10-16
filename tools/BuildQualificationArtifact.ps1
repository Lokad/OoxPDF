# Build one committed linked worktree, then delegate immutable publication to Freeze.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $WorktreeDirectory,
    [Parameter(Mandatory=$true)] [ValidatePattern('^(?:[0-9a-fA-F]{40}|[0-9a-fA-F]{64})$')] [string] $Revision,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory
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
$projects=@('tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj','src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj')
$freeze=Join-Path $source 'tools/FreezeQualificationArtifact.ps1'
foreach ($path in @($projects | ForEach-Object { Join-Path $source $_ }) + @($freeze)) {
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) { throw ('Qualification input is missing: '+$path) }
}
function Assert-Source {
    if ((Read-Git -Arguments @('rev-parse','HEAD')) -ne $Revision) { throw 'Qualification source revision differs' }
    if (@(Read-Git -Arguments @('status','--porcelain','--untracked-files=normal')).Count) { throw 'Commit source and public inputs before qualification builds' }
}
Assert-Source
New-Item -ItemType Directory -Path $logs -Force | Out-Null
$builds=@()
foreach ($relative in $projects) {
    $project=Join-Path $source $relative
    $log=Join-Path $logs ($builds.Count.ToString()+'.log')
    # Absolute project arguments make the caller's working directory irrelevant.
    & dotnet build $project -c Release --nologo -v minimal *> $log
    if ($LASTEXITCODE -ne 0) { throw ('Qualification build failed; inspect '+$log) }
    $builds += [ordered]@{Project=$project;Log=$log;Configuration='Release'}
    Assert-Source
}
& $freeze -OutputDirectory $output
$pin=Get-Content -LiteralPath (Join-Path $output 'identity.json') -Raw | ConvertFrom-Json
if ($pin.Revision -ne $Revision) { throw 'Frozen source revision differs' }
[ordered]@{Worktree=$source;Revision=$Revision;Builds=$builds;Artifact=$output;
    LibrarySha256=$pin.LibrarySha256;TestsRun=0} | ConvertTo-Json -Depth 5 |
    Set-Content -LiteralPath (Join-Path $logs 'evidence.json') -Encoding utf8
Write-Output ('Built and froze '+$Revision+' from explicit linked worktree; no tests replayed.')
