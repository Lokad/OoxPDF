# Aggregate reviewed test-only repairs without replaying unchanged checks.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputPath
)
$ErrorActionPreference='Stop'
$manifestPath=(Resolve-Path -LiteralPath $Manifest).Path
$base=Split-Path -Parent $manifestPath
$repo=Split-Path -Parent $PSScriptRoot
$config=Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$output=[IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $output) { throw 'Use a fresh output path; preserve prior reports' }
if ($config.Platform -notin @('Windows','Ubuntu')) { throw 'Declare one platform per report join' }
function Resolve-Input([string] $Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'A manifest input path is missing' }
    $path=if ([IO.Path]::IsPathFullyQualified($Name)) {$Name} else {Join-Path $base $Name}
    (Resolve-Path -LiteralPath $path).Path
}
function Pin([string] $Path) { [ordered]@{Path=$Path;Sha256=(Get-FileHash -LiteralPath $Path).Hash} }
$selectionPath=Resolve-Input $config.ExpectedSelection
$expected=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($name in @(Get-Content -LiteralPath $selectionPath -Raw | ConvertFrom-Json)) {
    if ([string]::IsNullOrWhiteSpace($name) -or !$expected.Add($name)) { throw 'Expected selection must contain unique exact method names' }
}
if (!$expected.Count) { throw 'An exact nonempty selection is required' }
$corrected=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($name in @($config.CorrectedMethods)) {
    if (!$expected.Contains($name) -or !$corrected.Add($name)) { throw 'Corrected methods must be unique members of the final selection' }
}
$initial=Resolve-Input $config.Initial.Artifact
$final=Resolve-Input $config.FinalArtifact
$attempts=@([pscustomobject]@{Artifact=$initial;Report=(Resolve-Input $config.Initial.Report);Repair=$false})
foreach ($repair in @($config.Repairs)) {
    $attempts += [pscustomobject]@{Artifact=(Resolve-Input $repair.Artifact);Report=(Resolve-Input $repair.Report);Repair=$true}
}
$pins=@{}
foreach ($artifact in @(@($attempts.Artifact)+@($final) | Select-Object -Unique)) {
    $pins[$artifact]=Get-Content -LiteralPath (Join-Path $artifact 'identity.json') -Raw | ConvertFrom-Json -AsHashtable
}
$first=$pins[$initial];$last=$pins[$final]
# Pre-0.3 artifacts were net10-only and did not carry a Framework field.
$framework=if ($first.Framework) {$first.Framework} else {'net10.0'}
if ($framework -notin @('net10.0','net11.0')) { throw 'Unknown frozen target framework' }
$runtime=@($first.Assets.Keys | Where-Object {$_ -notmatch '/Lokad\.OoxPdf\.Tests\.(dll|exe|pdb)$'})
if (!$runtime.Count -or !$first.Inputs.Count) { throw 'Frozen runtime and public input identities are required' }
foreach ($artifact in $pins.Keys) {
    $pin=$pins[$artifact]
    $selected=if ($pin.Framework) {$pin.Framework} else {'net10.0'}
    if ($selected -ne $framework) { throw 'Frozen target frameworks differ' }
    if ($pin.LibrarySha256 -ne $pin.Assets["src/Lokad.OoxPdf/bin/Release/$framework/Lokad.OoxPdf.dll"] -or
        $pin.TestAssemblySha256 -ne $pin.Assets["tests/Lokad.OoxPdf.Tests/bin/Release/$framework/Lokad.OoxPdf.Tests.dll"]) {
        throw 'Frozen assembly metadata differs from asset identities'
    }
    $sourceChanges=@(& git -C $repo diff --name-only $first.Revision $pin.Revision -- src `
        ':(top,glob)*.props' ':(top,glob)*.targets' Directory.Build.rsp global.json NuGet.Config nuget.config .gitattributes)
    if ($LASTEXITCODE -ne 0) { throw 'Frozen source revisions are unavailable' }
    if ($sourceChanges.Count) { throw 'Production source changed; qualify the changed runtime separately' }
    if ($pin.Assets.Count -ne $first.Assets.Count -or
        @($first.Assets.Keys | Where-Object {!$pin.Assets.ContainsKey($_)}).Count -or
        @($runtime | Where-Object {$pin.Assets[$_] -ne $first.Assets[$_]}).Count -or
        $pin.LibrarySha256 -ne $first.LibrarySha256) { throw 'Frozen runtime assets differ' }
    if ($pin.Inputs.Count -ne $first.Inputs.Count -or
        @($first.Inputs.Keys | Where-Object {$pin.Inputs[$_] -ne $first.Inputs[$_]}).Count) { throw 'Frozen inputs differ' }
    foreach ($name in $pin.Assets.Keys) {
        if ((Get-FileHash -LiteralPath (Join-Path $artifact $name)).Hash -ne $pin.Assets[$name]) { throw 'A pinned runtime file differs' }
    }
    foreach ($name in $pin.Inputs.Keys) {
        if ((Get-FileHash -LiteralPath (Join-Path $artifact $name)).Hash -ne $pin.Inputs[$name]) { throw 'A pinned public input differs' }
    }
}
# Earlier passes may retain their original harness, but the declared final
# harness must be the one used by the final supplied run, wherever it is stored.
if (!$last.TestAssemblySha256 -or $last.TestAssemblySha256 -ne $pins[$attempts[-1].Artifact].TestAssemblySha256) {
    throw 'Final artifact test assembly differs from the final report-producing harness'
}
$rows=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
$initialFailures=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$observedFailures=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$sources=@();$methodAttempts=0
foreach ($attempt in $attempts) {
    $reportHash=(Get-FileHash -LiteralPath $attempt.Report).Hash
    $report=Get-Content -LiteralPath $attempt.Report -Raw | ConvertFrom-Json
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $counts=@{passed=0;failed=0;skipped=0}
    foreach ($test in @($report.tests)) {
        if (!$expected.Contains($test.name) -or !$seen.Add($test.name)) { throw 'Report methods differ from the exact selection or contain duplicates' }
        if ($test.outcome -notin @('passed','failed','skipped')) { throw 'Unknown test outcome' }
        if ($attempt.Repair -and !$corrected.Contains($test.name)) { throw 'Repair contains an undeclared corrected method' }
        if ($test.outcome -eq 'failed') {
            $null=$observedFailures.Add($test.name)
            if (!$attempt.Repair) { $null=$initialFailures.Add($test.name) }
        }
        $counts[$test.outcome]++;$methodAttempts++
        $row=[ordered]@{}
        foreach ($property in $test.PSObject.Properties) { $row[$property.Name]=$property.Value }
        $row.SourceReport=$attempt.Report
        $row.SourceReportSha256=$reportHash
        $row.SourceRevision=$pins[$attempt.Artifact].Revision
        $row.SourceTestAssemblySha256=$pins[$attempt.Artifact].TestAssemblySha256
        $rows[$test.name]=$row
    }
    if (!$seen.Count) { throw 'A source report must contain at least one test method' }
    foreach ($outcome in @('passed','failed','skipped')) {
        if ($counts[$outcome] -ne $report.$outcome) { throw 'Report totals differ from method outcomes' }
    }
    if ((Get-FileHash -LiteralPath $attempt.Report).Hash -ne $reportHash) { throw 'A source report changed during the join' }
    $sources += [ordered]@{Report=(Pin $attempt.Report);ArtifactIdentity=(Pin (Join-Path $attempt.Artifact 'identity.json'));
        Repair=$attempt.Repair;Counts=$counts}
}
if ($rows.Count -ne $expected.Count -or @($expected | Where-Object {!$rows.ContainsKey($_)}).Count) {
    throw 'The final exact selection is incomplete'
}
if (@($rows.Values | Where-Object {$_.outcome -eq 'failed'}).Count -or
    @($observedFailures | Where-Object {$rows[$_].outcome -ne 'passed'}).Count) {
    throw 'A failed check remains unresolved; a skip cannot replace a failure'
}
$result=[ordered]@{
    Scope='Join reviewed test-only repairs; declared corrected methods still require source review.'
    Platform=$config.Platform;Framework=$framework;FrozenRevision=$last.Revision
    LibrarySha256=$last.LibrarySha256;TestAssemblySha256=$last.TestAssemblySha256
    passed=@($rows.Values | Where-Object {$_.outcome -eq 'passed'}).Count
    failed=0;skipped=@($rows.Values | Where-Object {$_.outcome -eq 'skipped'}).Count
    UniqueMethods=$rows.Count;MethodAttempts=$methodAttempts
    RuntimeAssetsUnchanged=$runtime.Count;InputsUnchanged=$first.Inputs.Count
    InitialFailuresRetained=@($initialFailures | Sort-Object)
    ObservedFailuresRetained=@($observedFailures | Sort-Object)
    CorrectedMethods=@($corrected | Sort-Object)
    Manifest=(Pin $manifestPath);Selection=(Pin $selectionPath);Tool=(Pin $PSCommandPath)
    Sources=$sources;tests=@($expected | Sort-Object | ForEach-Object {$rows[$_]})
}
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $output) | Out-Null
$staging=$output+'.staging-'+[Guid]::NewGuid().ToString('N')
[IO.File]::WriteAllText($staging,($result | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
[IO.File]::Move($staging,$output,$false)
"Joined $($rows.Count) exact methods; $($runtime.Count) runtime assets/$($first.Inputs.Count) inputs unchanged. No builds or tests."
