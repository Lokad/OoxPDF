# Exercise real package preflight and the restore boundary without running dotnet.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $PackagePath,
    [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$package=(Resolve-Path -LiteralPath $PackagePath).Path
if (!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/package-inspection-tests-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test output directory' }
$output=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$originalHash=(Get-FileHash -LiteralPath $package).Hash
$archive=[IO.Compression.ZipFile]::OpenRead($package)
try {
    $payloads=@{}
    foreach ($framework in @('net10.0','net11.0')) {
        $stream=$archive.GetEntry("lib/$framework/Lokad.OoxPdf.dll").Open()
        try { $bytes=[IO.MemoryStream]::new(); try { $stream.CopyTo($bytes);$payloads[$framework]=$bytes.ToArray() } finally {$bytes.Dispose()} }
        finally {$stream.Dispose()}
    }
} finally {$archive.Dispose()}
$hashes=@{}
foreach ($framework in $payloads.Keys) { $hashes[$framework]=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([byte[]]$payloads[$framework])) }
$tool=Join-Path $PSScriptRoot 'TestPackageSmoke.ps1'
$expectedFeed = Split-Path -Parent $package
$expectedVersion = [IO.Path]::GetFileNameWithoutExtension($package).Substring('Lokad.OoxPdf.'.Length)
# The restore mock supplies a different archive with the same selected DLL.
# The tool must refuse those bytes before it calls dotnet run.
$failedConsumers = [Collections.Generic.List[string]]::new()
function dotnet {
    if ($args[0] -eq 'restore') {
        $directory = Split-Path $args[1]
        [xml]$config = Get-Content -LiteralPath (Join-Path $directory 'NuGet.config') -Raw
        [xml]$project = Get-Content -LiteralPath $args[1] -Raw
        if ($config.configuration.packageSources.add.value -cne $expectedFeed -or
            $project.Project.ItemGroup.PackageReference.Version -cne "[$expectedVersion]") {
            throw 'Fixture: consumer restore must select the original local feed and exact version.'
        }
        $failedConsumers.Add($directory)
        foreach ($name in @('bin', 'obj', 'packages')) {
            $cache = New-Item -ItemType Directory -Path (Join-Path $directory $name)
            Set-Content -LiteralPath (Join-Path $cache.FullName 'failure.txt') -Value 'retain failed consumer'
        }
        $restored = New-Item -ItemType Directory -Path (Join-Path $directory "packages/lokad.ooxpdf/$expectedVersion") -Force
        Copy-Item -LiteralPath $differentRestoreArchive -Destination (Join-Path $restored.FullName "lokad.ooxpdf.$expectedVersion.nupkg")
        $global:LASTEXITCODE = 0
        return
    }
    throw 'Fixture: unexpected dotnet invocation.'
}
$checks=@()
function Check([string]$Name,[string]$InputPackage,[string[]]$Hash,[string[]]$Framework,
    [string]$Failure='', [switch]$ConsumerBoundary) {
    $arguments=@{Framework=$Framework}
    if ($InputPackage) { $arguments.PackagePath=$InputPackage }
    if ($Hash.Count -and $Hash[0]) { $arguments.ExpectedLibrarySha256=$Hash }
    if (!$ConsumerBoundary) { $arguments.InspectOnly=$true }
    $log=Join-Path $output ($Name+'.log')
    $caught=$null
    try { & $tool @arguments *> $log } catch { $caught=$_.Exception.Message }
    if ($Failure) {
        if (!$caught -or !$caught.Contains($Failure,[StringComparison]::Ordinal)) { throw ($Name+': wrong refusal: '+$caught) }
        $reports=@()
    } else {
        if ($caught) { throw ($Name+': '+$caught) }
        $lines=@(Get-Content -LiteralPath $log | Where-Object { $_ -like 'Evidence: *' })
        if ($lines.Count -ne $Framework.Count) { throw ($Name+': missing or duplicate evidence') }
        $reports=@(for ($index=0; $index -lt $lines.Count; $index++) {
            $reportPath=$lines[$index].Substring('Evidence: '.Length)
            $report=Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
            if (!$report.InspectionOnly -or $report.LibraryRebuilt -or $report.ConsumerRan -or
                $report.Framework -cne $Framework[$index] -or $report.LibrarySha256 -ne $Hash[$index]) {
                throw ($Name+': incorrect inspection result')
            }
            $report
        })
    }
    $script:checks += [ordered]@{Name=$Name;InputPackage=$InputPackage;
        InputPackageSha256=$(if($InputPackage){(Get-FileHash -LiteralPath $InputPackage).Hash}else{$null});
        Framework=$Framework;ExpectedLibrarySha256=$Hash;ConsumerBoundary=[bool]$ConsumerBoundary;
        ExpectedRefusal=$Failure;ObservedRefusal=$caught;Passed=$true;Log=$log;Reports=$reports}
    return $reports
}
function Write-Variant([string]$Name,[byte[]]$Net10Bytes,[string]$Omit='') {
    $folder=(New-Item -ItemType Directory -Path (Join-Path $output $Name)).FullName
    $target=Join-Path $folder ([IO.Path]::GetFileName($package))
    $sourceArchive=[IO.Compression.ZipFile]::OpenRead($package)
    try {
        $file=[IO.File]::Open($target,[IO.FileMode]::CreateNew)
        try {
            $zip=[IO.Compression.ZipArchive]::new($file,[IO.Compression.ZipArchiveMode]::Create,$true)
            try {
                foreach ($entry in $sourceArchive.Entries) {
                    if ($entry.FullName -ceq $Omit) { continue }
                    $copy=$zip.CreateEntry($entry.FullName)
                    $writer=$copy.Open()
                    try {
                        if ($Net10Bytes -and $entry.FullName -ceq 'lib/net10.0/Lokad.OoxPdf.dll') { $writer.Write($Net10Bytes) }
                        else { $reader=$entry.Open();try {$reader.CopyTo($writer)} finally {$reader.Dispose()} }
                    } finally {$writer.Dispose()}
                }
            } finally {$zip.Dispose()}
        } finally {$file.Dispose()}
    } finally {$sourceArchive.Dispose()}
    return $target
}
function Replace-UniqueUtf8([byte[]]$Bytes,[string]$Before,[string]$After) {
    $needle=[Text.Encoding]::UTF8.GetBytes($Before)
    $replacement=[Text.Encoding]::UTF8.GetBytes($After)
    if ($needle.Length -ne $replacement.Length) { throw 'Fixture replacement must retain metadata length' }
    # Latin1 is a byte-preserving search view, not an assembly parser. Production
    # preflight uses the BCL metadata reader to read the resulting attribute.
    $view=[Text.Encoding]::Latin1.GetString($Bytes)
    $pattern=[Text.Encoding]::Latin1.GetString($needle)
    $offset=$view.IndexOf($pattern,[StringComparison]::Ordinal)
    if ($offset -lt 0 -or $view.IndexOf($pattern,$offset+1,[StringComparison]::Ordinal) -ge 0) { throw 'Fixture metadata string is not unique' }
    $changed=[byte[]]$Bytes.Clone()
    [Buffer]::BlockCopy($replacement,0,$changed,$offset,$replacement.Length)
    return ,$changed
}
$net11=Check 'coherent-net11' $package $hashes['net11.0'] 'net11.0'
$frameworks=@('net11.0','net10.0')
$pins=@($hashes['net11.0'],$hashes['net10.0'])
$dual=@(Check 'coherent-dual' $package $pins $frameworks)
if ($net11.FrameworkIdentities.Count -ne 2 -or
    @($dual | Where-Object { $_.FrameworkIdentities.Count -ne 2 -or $_.PackageSha256 -ne $originalHash -or
        $_.LocalDocumentationLinksChecked -ne $net11.LocalDocumentationLinksChecked }).Count) {
    throw 'Expected shared archive, framework identities and documentation checks'
}
$version=$net11.FrameworkIdentities[0].InformationalVersion
$last=$version.Substring($version.Length-1)
$replacement=$version.Substring(0,$version.Length-1)+$(if($last -ceq '0') {'1'} else {'0'})
$mixed=Write-Variant 'mixed-build' (Replace-UniqueUtf8 $payloads['net10.0'] $version $replacement)
$null=Check 'mixed-build-refused' $mixed $hashes['net11.0'] 'net11.0' 'Packed framework library build identities differ.'
$missing=Write-Variant 'missing-version' (Replace-UniqueUtf8 $payloads['net10.0'] 'AssemblyInformationalVersionAttribute' 'XssemblyInformationalVersionAttribute')
$null=Check 'missing-version-refused' $missing $hashes['net11.0'] 'net11.0' 'Packed library version metadata is missing:'
$single=Write-Variant 'single-framework' $null 'lib/net10.0/Lokad.OoxPdf.dll'
$singleReport=Check 'single-framework-accepted' $single $hashes['net11.0'] 'net11.0'
if ($singleReport.FrameworkIdentities.Count -ne 1) { throw 'Single framework result differs' }
$null=Check 'missing-framework-refused' $single $hashes['net10.0'] 'net10.0' 'Packed library DLL is missing or duplicated.'
$null=Check 'wrong-qualified-hash-refused' $package ('0'*64) 'net11.0' 'Packed library differs from the qualified artifact.'
$null=Check 'missing-package-refused' '' $hashes['net11.0'] 'net11.0' 'InspectOnly requires an existing PackagePath.'
$null=Check 'missing-qualified-hash-refused' $package '' 'net11.0' 'Reusing a package requires the qualified library hash.'
$null=Check 'missing-batch-hash-refused' $package $pins[0] $frameworks 'Supply one expected library hash per framework, in the same order.'
$null=Check 'duplicate-consumer-refused' $package $pins @('net11.0','net11.0') 'Each consumer framework must be selected once.'
$null=Check 'second-hash-refused-before-restore' $package @($pins[0],('0'*64)) $frameworks 'Packed library differs from the qualified artifact. Framework: net10.0.' -ConsumerBoundary
$differentRestoreArchive=$single
$null=Check 'restore-from-pinned-feed-refuses-different-archive' $package $pins $frameworks 'Restored package bytes differ from the inspected package.' -ConsumerBoundary
if ($failedConsumers.Count -ne 1) { throw 'Failed consumer boundary was not observed exactly once.' }
foreach ($name in @('bin', 'obj', 'packages')) {
    if ((Get-Content -LiteralPath (Join-Path $failedConsumers[0] "$name/failure.txt") -Raw).Trim() -cne 'retain failed consumer') {
        throw 'Failed consumer cache was removed or changed.'
    }
}
if ((Get-FileHash -LiteralPath $package).Hash -ne $originalHash) { throw 'Original package changed' }
[ordered]@{Checks=$checks;Package=$package;PackageSha256=$originalHash;ToolSha256=(Get-FileHash -LiteralPath $tool).Hash;
    TestScriptSha256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;
    RealDotnetInvocations=0;LibraryBuilds=0;Conversions=0;OriginalPackageUnchanged=$true;
    OriginalFeedAndExactVersionSelected=$true;DifferentRestoredArchiveRefused=$true;FailedConsumerCachesRetained=$true} |
    ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $output 'evidence.json') -Encoding utf8
"$($checks.Count) package inspection checks passed; no builds, restores, conversions or assembly execution."
