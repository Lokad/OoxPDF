# Exercise report reuse with small identity fixtures; no build, conversion or test-runner replay.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/qualification-report-join-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh self-test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root=(Resolve-Path -LiteralPath $OutputDirectory).Path
$repo=Join-Path $root 'repo'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$join=Join-Path $repo 'tools/JoinQualificationTestReports.ps1'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'JoinQualificationTestReports.ps1') -Destination $join
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
    ($result.tests | Where-Object name -eq 'Unchanged').SourceTestAssemblySha256 -eq $result.TestAssemblySha256) {
    throw 'Joined result lost counts, failure provenance or unchanged-test origin'
}
$rejections=@()
$positiveInputs=@{}
foreach ($path in @($manifest,(Join-Path $root 'repair.json'),(Join-Path $root 'selection.json'),
    (Join-Path $repo 'artifacts/after/identity.json'))) { $positiveInputs[$path]=[IO.File]::ReadAllBytes($path) }
function Reject([string] $Name,[string] $Pattern) {
    $path=Join-Path $root ($Name+'.json')
    $caught=$null
    try { $null=& $join -Manifest $manifest -OutputPath $path } catch { $caught=$_.Exception.Message }
    if (!$caught -or $caught -notmatch $Pattern -or (Test-Path -LiteralPath $path)) { throw ('Missing pre-publication rejection: '+$Name+' / '+$caught) }
    $script:rejections += $Name
}
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
Write-Text (Join-Path $repo 'src/source.txt') 'changed production'
$null=Invoke-Git @('add','src/source.txt');$null=Invoke-Git @('commit','--quiet','-m','production change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-production' 'Production source changed'
foreach ($path in $positiveInputs.Keys) { [IO.File]::WriteAllBytes($path,$positiveInputs[$path]) }
$positivePins=@($result.Manifest,$result.Selection,$result.Tool)
foreach ($source in $result.Sources) { $positivePins+=@($source.Report,$source.ArtifactIdentity) }
foreach ($pin in $positivePins) {
    if ((Get-FileHash -LiteralPath $pin.Path).Hash -ne $pin.Sha256) { throw 'Negative fixtures invalidated retained positive provenance' }
}
Write-Json (Join-Path $root 'evidence.json') @{PositiveReportSha256=(Get-FileHash -LiteralPath $positive).Hash
    Rejections=$rejections;FilesRetained=$true;Builds=0;Conversions=0;RunnerInvocations=0}
"Qualification report join checks pass: one provenance-preserving join, $($rejections.Count) pre-publication rejections; no builds/tests/conversions."
