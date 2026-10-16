# Prepare a mutable parent-library runner without writing through frozen input links.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $TestArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $ParentArtifactDirectory,
    [Parameter(Mandatory=$true)] [string] $InputManifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory
)
$ErrorActionPreference = 'Stop'
$tests = (Resolve-Path -LiteralPath $TestArtifactDirectory).Path
$parent = (Resolve-Path -LiteralPath $ParentArtifactDirectory).Path
$testPin = Get-Content (Join-Path $tests 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
$parentPin = Get-Content (Join-Path $parent 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
$output = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($OutputDirectory))
function Assert-PhysicalDirectory([string] $Path) {
    $current = $Path
    while ($current) {
        if ((Test-Path -LiteralPath $current) -and
            ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'A mutable parent runtime cannot use directory links'
        }
        $next = Split-Path -Parent $current
        if ($next -eq $current) { break }
        $current = $next
    }
}
Assert-PhysicalDirectory $output
$ownerPath = Join-Path $output 'parent-runtime-owner.json'
if (Test-Path -LiteralPath $output) {
    if (!(Test-Path -LiteralPath $ownerPath)) { throw 'Refuse an unowned existing parent runtime' }
    $owner = Get-Content -LiteralPath $ownerPath -Raw | ConvertFrom-Json
    if ($owner.Version -ne 1 -or $owner.Directory -ne $output) { throw 'Parent runtime ownership differs' }
}
$casePrefix = 'tests/Lokad.OoxPdf.Tests/Cases/'
$inputNames = @(Get-Content -LiteralPath $InputManifest -Raw | ConvertFrom-Json)
if (!$inputNames.Count) { throw 'Select explicit public parent-test inputs' }
$jobs = @()
foreach ($name in @('Lokad.OoxPdf.Tests.dll','Lokad.OoxPdf.Tests.pdb','Lokad.OoxPdf.Tests.deps.json','Lokad.OoxPdf.Tests.runtimeconfig.json')) {
    $relative = 'tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/' + $name
    $jobs += [pscustomobject]@{ Source=(Join-Path $tests $relative); Relative=$relative; Hash=$testPin.Assets[$relative] }
}
$library = 'src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll'
$jobs += [pscustomobject]@{ Source=(Join-Path $parent $library); Relative='tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/Lokad.OoxPdf.dll'; Hash=$parentPin.LibrarySha256 }
foreach ($name in $inputNames) {
    $relative = $casePrefix + $name.Replace('\','/')
    if (!$testPin.Inputs.ContainsKey($relative)) { throw ('Input is outside the pinned public set: ' + $name) }
    $jobs += [pscustomobject]@{ Source=(Join-Path $tests $relative); Relative=$relative; Hash=$testPin.Inputs[$relative] }
}
# Validate every source and destination before publishing any file.
foreach ($job in $jobs) {
    $target = [IO.Path]::GetFullPath((Join-Path $output $job.Relative))
    if (!$target.StartsWith($output + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Parent target escaped its owned directory' }
    Assert-PhysicalDirectory (Split-Path $target)
    if (!$job.Hash -or (Get-FileHash -LiteralPath $job.Source).Hash -ne $job.Hash) { throw 'Pinned parent source differs' }
}
New-Item -ItemType Directory -Force -Path $output | Out-Null
if (!(Test-Path -LiteralPath $ownerPath)) {
    [ordered]@{Version=1;Directory=$output} | ConvertTo-Json | Set-Content -LiteralPath $ownerPath -Encoding utf8
}
foreach ($job in $jobs) {
    $target = Join-Path $output $job.Relative
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    # Copy into an independent file, then replace the directory entry. Never
    # overwrite a possibly hard-linked destination or clear its attributes.
    $temporary = $target + '.staging-' + [Guid]::NewGuid().ToString('N')
    [IO.File]::Copy($job.Source, $temporary, $false)
    (Get-Item -LiteralPath $temporary).IsReadOnly = $false
    if ((Get-FileHash -LiteralPath $temporary).Hash -ne $job.Hash) { throw 'Parent copy differs' }
    [IO.File]::Move($temporary, $target, $true)
}
[ordered]@{TestRevision=$testPin.Revision;ParentRevision=$parentPin.Revision;TestAssemblySha256=$testPin.TestAssemblySha256;ParentLibrarySha256=$parentPin.LibrarySha256;InputManifestSha256=(Get-FileHash -LiteralPath $InputManifest).Hash;Inputs=$inputNames;OwnedPhysicalInputs=$true;Rebuilt=$false} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'identity.json') -Encoding utf8
Write-Output "Parent runtime prepared with $($inputNames.Count) independent public inputs; frozen inputs remain unchanged."
