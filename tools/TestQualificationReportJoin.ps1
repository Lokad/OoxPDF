# Exercise report reuse with small identity fixtures; no build, conversion or test-runner replay.
[CmdletBinding()]
param([string] $OutputDirectory, [string] $JoinScript = (Join-Path $PSScriptRoot 'JoinQualificationTestReports.ps1'))
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/qualification-report-join-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh self-test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root=(Resolve-Path -LiteralPath $OutputDirectory).Path
$repo=Join-Path $root 'repo'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$join=Join-Path $repo 'tools/JoinQualificationTestReports.ps1'
Copy-Item -LiteralPath $JoinScript -Destination $join
function Write-Text([string] $Path,[string] $Text) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [IO.File]::WriteAllText($Path,$Text,[Text.UTF8Encoding]::new($false))
}
function Write-Json([string] $Path,$Value) { Write-Text $Path ($Value | ConvertTo-Json -Depth 12) }
function Invoke-Git([string[]] $Arguments) {
    $value=@(& git -C $repo @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) { throw ($value -join "\n") }
    $value
}
$null=Invoke-Git @('init','--quiet')
$null=Invoke-Git @('config','user.name','Qualification report test')
$null=Invoke-Git @('config','user.email','qualification@example.invalid')
$null=Invoke-Git @('config','commit.gpgsign','false')
Write-Text (Join-Path $repo '.gitignore') 'artifacts/'
Write-Text (Join-Path $repo 'src/source.txt') 'unchanged production'
$null=Invoke-Git @('add','.')
$null=Invoke-Git @('commit','--quiet','-m','initial source')
$firstRevision=[string](Invoke-Git @('rev-parse','HEAD'))
Write-Text (Join-Path $repo 'tests/observer.txt') 'reviewed test repair'
$null=Invoke-Git @('add','.')
$null=Invoke-Git @('commit','--quiet','-m','test repair')
$finalRevision=[string](Invoke-Git @('rev-parse','HEAD'))
$library='src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll'
$tests='tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/Lokad.OoxPdf.Tests.dll'
$deps='tests/Lokad.OoxPdf.Tests/bin/Release/net10.0/Lokad.OoxPdf.Tests.deps.json'
$inputName='tests/Lokad.OoxPdf.Tests/Cases/sample.txt'
foreach ($id in @('before','after')) {
    $artifact=Join-Path $repo ('artifacts/'+$id)
    Write-Text (Join-Path $artifact $library) 'same runtime'
    Write-Text (Join-Path $artifact $tests) ('test version '+$id)
    Write-Text (Join-Path $artifact $deps) 'same test dependencies'
    Write-Text (Join-Path $artifact $inputName) 'same public input'
    $assets=@{};foreach ($name in @($library,$tests,$deps)) { $assets[$name]=(Get-FileHash -LiteralPath (Join-Path $artifact $name)).Hash }
    Write-Json (Join-Path $artifact 'identity.json') @{
        Revision=$(if($id -eq 'before'){$firstRevision}else{$finalRevision})
        Assets=$assets;Inputs=@{$inputName=(Get-FileHash -LiteralPath (Join-Path $artifact $inputName)).Hash}
        LibrarySha256=$assets[$library];TestAssemblySha256=$assets[$tests]
    }
}
Write-Json (Join-Path $root 'initial.json') @{passed=1;failed=1;skipped=0;tests=@(
    @{name='Unchanged';outcome='passed';group='fixture';milliseconds=1;detail=''},
    @{name='Corrected';outcome='failed';group='fixture';milliseconds=1;detail='retained failure'})}
Write-Json (Join-Path $root 'repair.json') @{passed=2;failed=0;skipped=0;tests=@(
    @{name='Corrected';outcome='passed';group='fixture';milliseconds=1;detail=''},
    @{name='Neighbor';outcome='passed';group='fixture';milliseconds=1;detail=''})}
Write-Json (Join-Path $root 'selection.json') @('Unchanged','Corrected','Neighbor')
$config=[ordered]@{Platform='Windows';Initial=@{Artifact='repo/artifacts/before';Report='initial.json'}
    Repairs=@(@{Artifact='repo/artifacts/after';Report='repair.json'});FinalArtifact='repo/artifacts/after'
    ExpectedSelection='selection.json';CorrectedMethods=@('Corrected','Neighbor')}
$manifest=Join-Path $root 'manifest.json'
Write-Json $manifest $config
$positive=Join-Path $root 'joined.json'
$null=& $join -Manifest $manifest -OutputPath $positive
$result=Get-Content -LiteralPath $positive -Raw | ConvertFrom-Json
if ($result.passed -ne 3 -or $result.MethodAttempts -ne 4 -or $result.RuntimeAssetsUnchanged -ne 2 -or
    $result.InputsUnchanged -ne 1 -or @($result.InitialFailuresRetained)[0] -ne 'Corrected' -or
    @($result.ObservedFailuresRetained)[0] -ne 'Corrected' -or
    ($result.tests | Where-Object name -eq 'Unchanged').SourceTestAssemblySha256 -eq $result.TestAssemblySha256) {
    throw 'Joined result lost counts, failure provenance or unchanged-test origin'
}
$rejections=@()
$positiveInputs=@{}
foreach ($path in @($manifest,(Join-Path $root 'initial.json'),(Join-Path $root 'repair.json'),(Join-Path $root 'selection.json'),
    (Join-Path $repo 'artifacts/after/identity.json'))) { $positiveInputs[$path]=[IO.File]::ReadAllBytes($path) }
function Reject([string] $Name,[string] $Pattern) {
    $path=Join-Path $root ($Name+'.json')
    $caught=$null
    try { $null=& $join -Manifest $manifest -OutputPath $path } catch { $caught=$_.Exception.Message }
    if (!$caught -or $caught -notmatch $Pattern -or (Test-Path -LiteralPath $path)) { throw ('Missing pre-publication rejection: '+$Name+' / '+$caught) }
    $script:rejections += $Name
}
# A final artifact may be stored separately, but its harness must have been run.
Write-Text (Join-Path $repo 'tests/observer.txt') 'unrun final harness'
$null=Invoke-Git @('add','tests/observer.txt');$null=Invoke-Git @('commit','--quiet','-m','unrun test repair')
$unrunRevision=[string](Invoke-Git @('rev-parse','HEAD'))
$unrun=Join-Path $repo 'artifacts/unrun-final'
Copy-Item -LiteralPath (Join-Path $repo 'artifacts/after') -Destination $unrun -Recurse
Write-Text (Join-Path $unrun $tests) 'different final test version, never run'
$unrunPin=Get-Content -LiteralPath (Join-Path $unrun 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
$unrunPin.Revision=$unrunRevision
$unrunPin.TestAssemblySha256=$unrunPin.Assets[$tests]=(Get-FileHash -LiteralPath (Join-Path $unrun $tests)).Hash
Write-Json (Join-Path $unrun 'identity.json') $unrunPin
$config.FinalArtifact='repo/artifacts/unrun-final';Write-Json $manifest $config
Reject 'untested-final-harness' 'Final artifact test assembly differs'
$config.FinalArtifact='repo/artifacts/before';Write-Json $manifest $config
Reject 'earlier-final-harness' 'Final artifact test assembly differs'
# An empty final report must not stand in for execution of the new harness.
Write-Json (Join-Path $root 'empty-final.json') @{passed=0;failed=0;skipped=0;tests=@()}
$config.FinalArtifact='repo/artifacts/unrun-final'
$config.Repairs+=@{Artifact='repo/artifacts/unrun-final';Report='empty-final.json'}
Write-Json $manifest $config
Reject 'empty-final-run' 'at least one test method'
[IO.File]::WriteAllBytes($manifest,$positiveInputs[$manifest])
$config=Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json -AsHashtable
$copy=Join-Path $repo 'artifacts/tested-final-copy'
Copy-Item -LiteralPath (Join-Path $repo 'artifacts/after') -Destination $copy -Recurse
$copyConfig=$config.Clone();$copyConfig.FinalArtifact='repo/artifacts/tested-final-copy'
$copyManifest=Join-Path $root 'tested-final-copy-manifest.json';Write-Json $copyManifest $copyConfig
$copyOutput=Join-Path $root 'tested-final-copy.json'
$null=& $join -Manifest $copyManifest -OutputPath $copyOutput
$copyResult=Get-Content -LiteralPath $copyOutput -Raw | ConvertFrom-Json
if ($copyResult.TestAssemblySha256 -ne $result.TestAssemblySha256 -or $copyResult.passed -ne 3 -or
    $copyResult.MethodAttempts -ne 4 -or
    ($copyResult.tests | Where-Object name -eq 'Unchanged').SourceTestAssemblySha256 -eq $copyResult.TestAssemblySha256) {
    throw 'Equivalent final copies must preserve tested identity and earlier-pass provenance'
}
# A method first failing during a repair must also finish with a real pass.
Write-Json (Join-Path $root 'repair-failed.json') @{passed=1;failed=1;skipped=0;tests=@(
    @{name='Corrected';outcome='passed'},@{name='Neighbor';outcome='failed';detail='failure during repair'})}
Write-Json (Join-Path $root 'repair-skipped.json') @{passed=0;failed=0;skipped=1;tests=@(
    @{name='Neighbor';outcome='skipped';detail='cannot hide a repair failure'})}
Write-Json (Join-Path $root 'repair-passed.json') @{passed=1;failed=0;skipped=0;tests=@(
    @{name='Neighbor';outcome='passed'})}
$config.Repairs=@(@{Artifact='repo/artifacts/after';Report='repair-failed.json'},
    @{Artifact='repo/artifacts/after';Report='repair-skipped.json'})
$config.Initial.Report='initial-repair-sequence.json'
Write-Json $manifest $config
foreach ($initialOutcome in @('skipped','passed')) {
    $initial=Get-Content -LiteralPath (Join-Path $root 'initial.json') -Raw | ConvertFrom-Json
    $initial.tests=@($initial.tests | Where-Object name -ne 'Neighbor') + @(@{name='Neighbor';outcome=$initialOutcome})
    $initial.passed=@($initial.tests | Where-Object outcome -eq 'passed').Count
    $initial.skipped=@($initial.tests | Where-Object outcome -eq 'skipped').Count
    Write-Json (Join-Path $root 'initial-repair-sequence.json') $initial
    Reject ('repair-failure-after-'+$initialOutcome+'-replaced-by-skip') 'failed check remains'
}
$config.Repairs=@(@{Artifact='repo/artifacts/after';Report='repair-failed.json'})
Write-Json $manifest $config
Reject 'terminal-repair-failure' 'failed check remains'
$config.Repairs+=@{Artifact='repo/artifacts/after';Report='repair-passed.json'}
$resolvedManifest=Join-Path $root 'repair-failure-resolved-manifest.json'
Write-Json $resolvedManifest $config
$resolvedPath=Join-Path $root 'repair-failure-resolved.json'
$null=& $join -Manifest $resolvedManifest -OutputPath $resolvedPath
$resolved=Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
if ($resolved.passed -ne 3 -or $resolved.failed -ne 0 -or $resolved.skipped -ne 0 -or
    @($resolved.InitialFailuresRetained).Count -ne 1 -or $resolved.InitialFailuresRetained[0] -ne 'Corrected' -or
    @($resolved.ObservedFailuresRetained).Count -ne 2 -or $resolved.ObservedFailuresRetained[1] -ne 'Neighbor' -or
    $resolved.Sources.Count -ne 3) { throw 'Resolved repair lost failures or attempt provenance' }
foreach ($path in @($manifest,(Join-Path $root 'initial.json'))) { [IO.File]::WriteAllBytes($path,$positiveInputs[$path]) }
$config=Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json -AsHashtable
$config.CorrectedMethods=@('Corrected');Write-Json $manifest $config
Reject 'undeclared-repair' 'undeclared corrected'
$config.CorrectedMethods=@('Corrected','Neighbor');Write-Json $manifest $config
Write-Json (Join-Path $root 'selection.json') @('Corrected','Neighbor')
Reject 'dropped-method' 'exact selection'
Write-Json (Join-Path $root 'selection.json') @('Unchanged','Corrected','Neighbor')
Write-Json (Join-Path $root 'repair.json') @{passed=1;failed=0;skipped=1;tests=@(
    @{name='Corrected';outcome='skipped';group='fixture';milliseconds=0;detail='cannot hide failure'},
    @{name='Neighbor';outcome='passed';group='fixture';milliseconds=1;detail=''})}
Reject 'failure-replaced-by-skip' 'failed check remains'
$finalPinPath=Join-Path $repo 'artifacts/after/identity.json'
$finalPin=Get-Content -LiteralPath $finalPinPath -Raw | ConvertFrom-Json -AsHashtable
$finalPin.Inputs[$inputName]='F'*64;Write-Json $finalPinPath $finalPin
Reject 'changed-input' 'Frozen inputs differ'
$finalPin.Inputs[$inputName]=(Get-FileHash -LiteralPath (Join-Path $repo ('artifacts/after/'+$inputName))).Hash
$finalPin.Assets[$deps]='E'*64;Write-Json $finalPinPath $finalPin
Reject 'changed-runtime' 'Frozen runtime assets differ'
$finalPin.Assets[$deps]=(Get-FileHash -LiteralPath (Join-Path $repo ('artifacts/after/'+$deps))).Hash
Write-Text (Join-Path $repo 'Directory.Build.props') '<Project><PropertyGroup><CheckForOverflowUnderflow>true</CheckForOverflowUnderflow></PropertyGroup></Project>'
$null=Invoke-Git @('add','Directory.Build.props');$null=Invoke-Git @('commit','--quiet','-m','production build-setting change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-build-setting' 'Production source changed'
Write-Text (Join-Path $repo 'src/source.txt') 'changed production'
$null=Invoke-Git @('add','src/source.txt');$null=Invoke-Git @('commit','--quiet','-m','production change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-production' 'Production source changed'
foreach ($path in $positiveInputs.Keys) { [IO.File]::WriteAllBytes($path,$positiveInputs[$path]) }
$positivePins=@($result.Manifest,$result.Selection,$result.Tool,$resolved.Manifest,$resolved.Selection,$resolved.Tool,
    $copyResult.Manifest,$copyResult.Selection,$copyResult.Tool)
foreach ($source in @($result.Sources)+@($resolved.Sources)+@($copyResult.Sources)) { $positivePins+=@($source.Report,$source.ArtifactIdentity) }
foreach ($pin in $positivePins) {
    if ((Get-FileHash -LiteralPath $pin.Path).Hash -ne $pin.Sha256) { throw 'Negative fixtures invalidated retained positive provenance' }
}
Write-Json (Join-Path $root 'evidence.json') @{PositiveReportSha256=(Get-FileHash -LiteralPath $positive).Hash
    ResolvedRepairReportSha256=(Get-FileHash -LiteralPath $resolvedPath).Hash
    TestedFinalCopyReportSha256=(Get-FileHash -LiteralPath $copyOutput).Hash;PositiveJoins=3
    Rejections=$rejections;FilesRetained=$true;Builds=0;Conversions=0;RunnerInvocations=0}
"Qualification report join checks pass: three provenance-preserving joins, $($rejections.Count) pre-publication rejections; no builds/tests/conversions."
