# Exercise report reuse with small identity fixtures; no build, conversion or test-runner replay.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [string] $OutputDirectory, [string] $JoinScript = (Join-Path $PSScriptRoot 'JoinQualificationTestReports.ps1'))
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/qualification-report-join-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh self-test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root=(Resolve-Path -LiteralPath $OutputDirectory).Path
$repo=Join-Path $root 'repo'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$join=Join-Path $repo 'tools/JoinQualificationTestReports.ps1'
Copy-Item -LiteralPath $JoinScript -Destination $join
Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $JoinScript) 'QualificationBuildIdentity.ps1') -Destination (Split-Path -Parent $join)
function Write-Text([string] $Path,[string] $Text) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [IO.File]::WriteAllText($Path,$Text,[Text.UTF8Encoding]::new($false))
}
function Write-Json([string] $Path,$Value) { Write-Text $Path ($Value | ConvertTo-Json -Depth 12) }
function Write-BoundReport([string] $Name,$Report,[string] $Artifact='after') {
    $pin=Get-Content -LiteralPath (Join-Path $repo "artifacts/$Artifact/identity.json") -Raw | ConvertFrom-Json
    if (!$Report.ContainsKey('provenance')) { $Report.provenance=@{} }
    $Report.provenance.LibrarySha256=$pin.LibrarySha256
    $Report.provenance.TestAssemblySha256=$pin.TestAssemblySha256
    Write-Json (Join-Path $root $Name) $Report
}
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
$library="src/Lokad.OoxPdf/bin/Release/$Framework/Lokad.OoxPdf.dll"
$tests="tests/Lokad.OoxPdf.Tests/bin/Release/$Framework/Lokad.OoxPdf.Tests.dll"
$deps="tests/Lokad.OoxPdf.Tests/bin/Release/$Framework/Lokad.OoxPdf.Tests.deps.json"
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
        Framework=$(if ($Framework -eq 'net11.0') {$Framework} else {$null});Assets=$assets;Inputs=@{$inputName=(Get-FileHash -LiteralPath (Join-Path $artifact $inputName)).Hash}
        LibrarySha256=$assets[$library];TestAssemblySha256=$assets[$tests]
    }
}
Write-BoundReport 'initial.json' @{passed=1;failed=1;skipped=0;tests=@(
    @{name='Unchanged';outcome='passed';group='fixture';milliseconds=1;detail=''},
    @{name='Corrected';outcome='failed';group='fixture';milliseconds=1;detail='retained failure'})} 'before'
Write-BoundReport 'repair.json' @{passed=2;failed=0;skipped=0;tests=@(
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
function Reject([string] $Name,[string] $Pattern,[string] $ManifestPath=$manifest) {
    $path=Join-Path $root ($Name+'.json')
    $caught=$null
    try { $null=& $join -Manifest $ManifestPath -OutputPath $path } catch { $caught=$_.Exception.Message }
    if (!$caught -or $caught -notmatch $Pattern -or (Test-Path -LiteralPath $path)) { throw ('Missing pre-publication rejection: '+$Name+' / '+$caught) }
    $script:rejections += $Name
}
$toolConfig=Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json -AsHashtable
$toolConfig.Initial.Report='initial-with-tool.json'
$initialTool=Get-Content -LiteralPath (Join-Path $root 'initial.json') -Raw | ConvertFrom-Json -AsHashtable
$repairTool=Get-Content -LiteralPath (Join-Path $root 'repair.json') -Raw | ConvertFrom-Json -AsHashtable
$binding='OOXPDF_TEST_PDF_INSPECTOR'
$firstPath=Join-Path $root 'initial-tool/PdfInspect.dll'
$secondPath=Join-Path $root 'repair-tool/PdfInspect.dll'
Write-Text $firstPath 'one frozen validation tool'
Write-Text $secondPath 'one frozen validation tool'
$initialTool.provenance.ValidationTools=@{$binding=@{Path=$firstPath;Sha256=(Get-FileHash -LiteralPath $firstPath).Hash}}
Write-Json (Join-Path $root $toolConfig.Initial.Report) $initialTool
$repairTool.provenance.ValidationTools=@{$binding=@{Path=$secondPath;Sha256=(Get-FileHash -LiteralPath $secondPath).Hash}}
$toolChecks=@()
$toolPins=@()
foreach ($case in 'relocated','disjoint','changed','malformed') {
    $toolManifest=Join-Path $root ($case+'-tool-manifest.json')
    $toolConfig.Repairs[0].Report=$case+'-repair-tool.json'
    Write-Json $toolManifest $toolConfig
    switch ($case) {
        'disjoint' {
            $secondPath=Join-Path $root 'repair-tool/VisualDiff.dll'
            Write-Text $secondPath 'independent frozen validation tool'
            $repairTool.provenance.ValidationTools=@{OOXPDF_TEST_VISUAL_DIFF=@{Path=$secondPath;Sha256=(Get-FileHash -LiteralPath $secondPath).Hash}}
        }
        'changed' {
            $secondPath=Join-Path $root 'changed-tool/PdfInspect.dll'
            Write-Text $secondPath 'different frozen validation tool'
            $repairTool.provenance.ValidationTools=@{$binding=@{Path=$secondPath;Sha256=(Get-FileHash -LiteralPath $secondPath).Hash}}
        }
        'malformed' { $repairTool.provenance.ValidationTools[$binding].Sha256='missing' }
    }
    Write-Json (Join-Path $root $toolConfig.Repairs[0].Report) $repairTool
    if ($case -eq 'changed') {
        Reject 'changed-validation-tool' "Source reports bind different validation tool bytes: $binding" $toolManifest
    }
    elseif ($case -eq 'malformed') {
        Reject 'malformed-validation-tool' "Invalid validation tool evidence: $binding" $toolManifest
    }
    else {
        $toolOutput=Join-Path $root ($case+'-validation-tool.json')
        $null=& $join -Manifest $toolManifest -OutputPath $toolOutput
        $toolResult=Get-Content -LiteralPath $toolOutput -Raw | ConvertFrom-Json
        $secondBinding=if ($case -eq 'disjoint') {'OOXPDF_TEST_VISUAL_DIFF'} else {$binding}
        if ($toolResult.passed -ne 3 -or $toolResult.Sources[0].ValidationTools.$binding.Path -cne $firstPath -or
            $toolResult.Sources[1].ValidationTools.$secondBinding.Path -cne $secondPath -or
            $toolResult.Sources[1].ValidationTools.$secondBinding.Sha256 -cne (Get-FileHash -LiteralPath $secondPath).Hash) {
            throw 'Joined report omitted the bound validation tool provenance'
        }
        $toolChecks+=$case
        $toolPins+=@($toolResult.Manifest,$toolResult.Sources[0].Report,$toolResult.Sources[1].Report,
            $toolResult.Sources[0].ValidationTools.$binding,$toolResult.Sources[1].ValidationTools.$secondBinding)
    }
}
$repairPath=Join-Path $root 'repair.json'
foreach ($fault in @('missing-provenance','wrong-harness','wrong-library')) {
    $report=Get-Content -LiteralPath $repairPath -Raw | ConvertFrom-Json -AsHashtable
    switch ($fault) {
        'missing-provenance' { $report.Remove('provenance') }
        'wrong-harness' {
            $original=Get-Content -LiteralPath (Join-Path $root 'initial.json') -Raw | ConvertFrom-Json
            $report.provenance.TestAssemblySha256=$original.provenance.TestAssemblySha256
        }
        'wrong-library' { $report.provenance.LibrarySha256='F'*64 }
    }
    $name=$fault+'-source.json'
    Write-Json (Join-Path $root $name) $report
    $config.Repairs[0].Report=$name
    Write-Json $manifest $config
    $pattern=if ($fault -eq 'missing-provenance') {'requires producing assembly evidence'} else {'producing assemblies differ'}
    Reject $fault $pattern
}
$config.Repairs[0].Report='repair.json'
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
Write-BoundReport 'empty-final.json' @{passed=0;failed=0;skipped=0;tests=@()} 'unrun-final'
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
$snapshot=Join-Path $repo 'artifacts/shared-inputs'
$snapshotInput=Join-Path $snapshot $inputName
$otherInput=Join-Path (Join-Path $repo 'artifacts/other-inputs') $inputName
Write-Text $snapshotInput 'same public input'
Write-Text $otherInput 'changed public input'
$linkType=if ($IsWindows) {'Junction'} else {'SymbolicLink'}
function Write-SnapshotArtifact([string] $Name, [string] $From, [string] $InputTarget) {
    $source=Join-Path $repo ('artifacts/'+$From)
    $directory=Join-Path $repo ('artifacts/'+$Name)
    $pin=Get-Content -LiteralPath (Join-Path $source 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
    $pin.InputSnapshot=$snapshot
    foreach ($asset in $pin.Assets.Keys) {
        $file=Join-Path $directory $asset
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $file) | Out-Null
        Copy-Item -LiteralPath (Join-Path $source $asset) -Destination $file
    }
    $inputFile=Join-Path $directory $inputName
    if ($InputTarget) {
        $leaf=Split-Path -Parent $inputFile
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $leaf) | Out-Null
        New-Item -ItemType $linkType -Path $leaf -Target (Split-Path -Parent $InputTarget) | Out-Null
    } else {
        Write-Text $inputFile 'same public input'
    }
    Write-Json (Join-Path $directory 'identity.json') $pin
}
Write-SnapshotArtifact 'linked-first' 'before' $snapshotInput
Write-SnapshotArtifact 'linked-repair' 'after' $snapshotInput
Write-SnapshotArtifact 'standalone' 'after' ''
Write-SnapshotArtifact 'retargeted' 'after' $otherInput
$snapshotChecks=@()
foreach ($repair in 'linked-repair','standalone','retargeted') {
    $linkConfig=$config.Clone()
    $linkConfig.Initial=@{Artifact='repo/artifacts/linked-first';Report='initial.json'}
    $linkConfig.FinalArtifact='repo/artifacts/'+$repair
    $linkConfig.Repairs=@(@{Artifact=$linkConfig.FinalArtifact;Report='repair.json'})
    $linkManifest=Join-Path $root ($repair+'-manifest.json')
    Write-Json $linkManifest $linkConfig
    if ($repair -eq 'retargeted') {
        Reject 'retargeted-fixture-link' 'A pinned public input differs' $linkManifest
        continue
    }
    $reads=[Collections.Generic.List[string]]::new()
    $linkOutput=Join-Path $root ($repair+'-join.json')
    & {
        function Get-FileHash([string] $LiteralPath) {
            if ($LiteralPath.Replace('\','/').EndsWith('/'+$inputName,[StringComparison]::Ordinal)) { $reads.Add($LiteralPath) }
            Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $LiteralPath
        }
        $null=& $join -Manifest $linkManifest -OutputPath $linkOutput
    }
    $expectedReads=@($snapshotInput)
    if ($repair -eq 'standalone') { $expectedReads+=Join-Path $repo ('artifacts/standalone/'+$inputName) }
    $linked=Get-Content -LiteralPath $linkOutput -Raw | ConvertFrom-Json
    if ($reads.Count -ne $expectedReads.Count -or @(Compare-Object $expectedReads $reads -CaseSensitive).Count -or
        $linked.passed -ne $result.passed -or $linked.MethodAttempts -ne $result.MethodAttempts -or
        $linked.InputsUnchanged -ne 1) { throw 'Fixture ownership changed hash coverage or the joined result' }
    $snapshotChecks+=@{Artifact=$repair;FixtureReads=$reads.Count;Report=(Get-FileHash -LiteralPath $linkOutput).Hash}
}
# A retained skip follows the policy of the report that produced it.
$skipChecks=@()
foreach ($policy in @('unexpected','undeclared','expected','fast-group-only')) {
    $provenance=@{}
    switch ($policy) {
        'unexpected' {
            $provenance.GatePassed=$false
            $provenance.UnexpectedSkips=@('Neighbor')
            $provenance.ExpectedSkips=@()
            $provenance.SkipValidationMethods=@('Corrected','Neighbor')
        }
        'expected' { $provenance.ExpectedSkips=@('Neighbor') }
        'fast-group-only' { $provenance.SkipValidationMethods=@('Corrected') }
    }
    $name='skip-'+$policy+'-source.json'
    Write-BoundReport $name @{passed=1;failed=0;skipped=1;provenance=$provenance;tests=@(
        @{name='Corrected';outcome='passed'},@{name='Neighbor';outcome='skipped'})}
    $config.Repairs=@(@{Artifact='repo/artifacts/after';Report=$name})
    Write-Json $manifest $config
    if ($policy -in @('unexpected','undeclared')) {
        Reject ('final-skip-'+$policy) 'Unexpected skips remain.*Neighbor'
    }
    else {
        $joined=Join-Path $root ('joined-skip-'+$policy+'.json')
        $null=& $join -Manifest $manifest -OutputPath $joined
        $allowed=Get-Content -LiteralPath $joined -Raw | ConvertFrom-Json
        if ($allowed.passed -ne 2 -or $allowed.skipped -ne 1 -or
            ($allowed.tests | Where-Object name -ceq 'Neighbor').outcome -cne 'skipped') {
            throw 'Joined report lost a declared skip'
        }
        $skipChecks+=$policy
    }
}
$config.Repairs=@(@{Artifact='repo/artifacts/after';Report='skip-unexpected-source.json'},
    @{Artifact='repo/artifacts/after';Report='repair.json'})
$retriedManifest=Join-Path $root 'retried-skip-manifest.json';Write-Json $retriedManifest $config
$retriedPath=Join-Path $root 'retried-skip.json'
$null=& $join -Manifest $retriedManifest -OutputPath $retriedPath
$retried=Get-Content -LiteralPath $retriedPath -Raw | ConvertFrom-Json
if ($retried.passed -ne 3 -or $retried.skipped -ne 0 -or $retried.Sources.Count -ne 3) {
    throw 'A later pass must resolve an unexpected skip and retain both source reports'
}
$skipChecks+='unexpected-skip-retried'
# A method first failing during a repair must also finish with a real pass.
Write-BoundReport 'repair-failed.json' @{passed=1;failed=1;skipped=0;tests=@(
    @{name='Corrected';outcome='passed'},@{name='Neighbor';outcome='failed';detail='failure during repair'})}
Write-BoundReport 'repair-skipped.json' @{passed=0;failed=0;skipped=1;tests=@(
    @{name='Neighbor';outcome='skipped';detail='cannot hide a repair failure'})}
Write-BoundReport 'repair-passed.json' @{passed=1;failed=0;skipped=0;tests=@(
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
Write-BoundReport 'repair.json' @{passed=1;failed=0;skipped=1;tests=@(
    @{name='Corrected';outcome='skipped';group='fixture';milliseconds=0;detail='cannot hide failure'},
    @{name='Neighbor';outcome='passed';group='fixture';milliseconds=1;detail=''})}
Reject 'failure-replaced-by-skip' 'failed check remains'
$finalPinPath=Join-Path $repo 'artifacts/after/identity.json'
$finalPin=Get-Content -LiteralPath $finalPinPath -Raw | ConvertFrom-Json -AsHashtable
$originalFramework=$finalPin.Framework
$finalPin.Framework=if ($Framework -eq 'net11.0') {'net10.0'} else {'net11.0'}
Write-Json $finalPinPath $finalPin
Reject 'mixed-framework' 'Frozen target frameworks differ'
$finalPin.Framework=$originalFramework
$finalPin.Inputs[$inputName]='F'*64;Write-Json $finalPinPath $finalPin
Reject 'changed-input' 'Frozen inputs differ'
$finalPin.Inputs[$inputName]=(Get-FileHash -LiteralPath (Join-Path $repo ('artifacts/after/'+$inputName))).Hash
$finalPin.Assets[$deps]='E'*64;Write-Json $finalPinPath $finalPin
Reject 'changed-runtime' 'Frozen runtime assets differ'
$finalPin.Assets[$deps]=(Get-FileHash -LiteralPath (Join-Path $repo ('artifacts/after/'+$deps))).Hash
Write-Text (Join-Path $repo 'Directory.Build.props') '<Project><PropertyGroup><CheckForOverflowUnderflow>true</CheckForOverflowUnderflow></PropertyGroup></Project>'
$null=Invoke-Git @('add','Directory.Build.props');$null=Invoke-Git @('commit','--quiet','-m','production build-setting change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-build-setting' 'Production source or build inputs changed'
# Branch from the original final revision to isolate the nested import.
$null=Invoke-Git @('checkout','--quiet','--detach',$finalRevision)
Write-Text (Join-Path $repo 'tools/Directory.Build.props') '<Project><PropertyGroup><CheckForOverflowUnderflow>true</CheckForOverflowUnderflow></PropertyGroup></Project>'
$null=Invoke-Git @('add','tools/Directory.Build.props');$null=Invoke-Git @('commit','--quiet','-m','inspector build-setting change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-tools-build-setting' 'Production source or build inputs changed'
$null=Invoke-Git @('checkout','--quiet','--detach',$finalRevision)
Write-Text (Join-Path $repo 'src/source.txt') 'changed production'
$null=Invoke-Git @('add','src/source.txt');$null=Invoke-Git @('commit','--quiet','-m','production change')
$finalPin.Revision=[string](Invoke-Git @('rev-parse','HEAD'));Write-Json $finalPinPath $finalPin
Reject 'changed-production' 'Production source or build inputs changed'
foreach ($path in $positiveInputs.Keys) { [IO.File]::WriteAllBytes($path,$positiveInputs[$path]) }
$positivePins=@($result.Manifest,$result.Selection,$result.Tool,$resolved.Manifest,$resolved.Selection,$resolved.Tool,
    $copyResult.Manifest,$copyResult.Selection,$copyResult.Tool)+$toolPins
foreach ($source in @($result.Sources)+@($resolved.Sources)+@($copyResult.Sources)) { $positivePins+=@($source.Report,$source.ArtifactIdentity) }
foreach ($pin in $positivePins) {
    if ((Get-FileHash -LiteralPath $pin.Path).Hash -ne $pin.Sha256) { throw 'Negative fixtures invalidated retained positive provenance' }
}
Write-Json (Join-Path $root 'evidence.json') @{PositiveReportSha256=(Get-FileHash -LiteralPath $positive).Hash
    ResolvedRepairReportSha256=(Get-FileHash -LiteralPath $resolvedPath).Hash
    TestedFinalCopyReportSha256=(Get-FileHash -LiteralPath $copyOutput).Hash;PositiveJoins=(3+$skipChecks.Count+$toolChecks.Count+$snapshotChecks.Count)
    ValidationToolPasses=$toolChecks
    SkipPolicyPasses=$skipChecks
    SnapshotChecks=$snapshotChecks;Rejections=$rejections;FilesRetained=$true;Builds=0;Conversions=0;RunnerInvocations=0}
"Qualification report join checks pass: $(3+$skipChecks.Count+$toolChecks.Count+$snapshotChecks.Count) provenance-preserving joins, $($rejections.Count) pre-publication rejections; no builds/tests/conversions."
