# Exercise the real wrapper and existing native tool; never build or export Office.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $InputPdf,
    [Parameter(Mandatory=$true)] [string] $RasterizerPath,
    [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$inputFull=(Resolve-Path -LiteralPath $InputPdf).Path
$rasterizer=(Resolve-Path -LiteralPath $RasterizerPath).Path
if (!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/raster-path-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
# The copied wrapper, tool assembly and native loader path all contain spaces.
$ownedRepo=Join-Path $root 'repository with spaces'
$ownedTools=Join-Path $ownedRepo 'tools'
$ownedRuntime=Join-Path $ownedTools 'Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/net10.0'
New-Item -ItemType Directory -Path $ownedRuntime -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'RasterizePdf.ps1') -Destination $ownedTools
foreach ($file in Get-ChildItem -LiteralPath (Split-Path $rasterizer) -File) {
    Copy-Item -LiteralPath $file.FullName -Destination $ownedRuntime
}
$nativeRelative='vendor/pdfium/win-x64'
$nativeDirectory=Join-Path $ownedTools $nativeRelative
New-Item -ItemType Directory -Path (Join-Path $nativeDirectory 'bin') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot ($nativeRelative+'/pdfium.sha256')) -Destination $nativeDirectory
Copy-Item -LiteralPath (Join-Path $PSScriptRoot ($nativeRelative+'/bin/pdfium.dll')) -Destination (Join-Path $nativeDirectory 'bin')
$plain=Join-Path $root 'input.pdf'
$spaced=Join-Path $root 'input file.pdf'
Copy-Item -LiteralPath $inputFull -Destination $plain
Copy-Item -LiteralPath $inputFull -Destination $spaced
$invalid=Join-Path $root 'invalid.pdf'
[IO.File]::WriteAllText($invalid,'Owned invalid PDF sentinel')
$rows=@()
$expectedHashes=$null
foreach ($job in @(
    @{Id='plain';Input=$plain;Output='plain-raster'},
    @{Id='spaced-input';Input=$spaced;Output='input-raster'},
    @{Id='spaced-output';Input=$plain;Output='output raster'},
    @{Id='invalid-pdf';Input=$invalid;Output='invalid-raster';Failure='failed with exit code'},
    @{Id='zero-deadline';Input=$plain;Output='timeout-raster';Failure='timed out after 0 seconds';Timeout=0})) {
    $destination=Join-Path $root $job.Output
    $timeout=if ($job.ContainsKey('Timeout')) { $job.Timeout } else { 30 }
    $failure=$null
    try {
        & (Join-Path $ownedTools 'RasterizePdf.ps1') -InputPdf $job.Input -OutputDirectory $destination -Dpi 36 -SkipBuild -TimeoutSeconds $timeout *> (Join-Path $root ($job.Id+'.log'))
    }
    catch { $failure=$_.Exception.Message }
    $pages=@(Get-ChildItem -LiteralPath $destination -Filter 'page-*.png' -File)
    if ($job.Failure) {
        if (!$failure -or !$failure.Contains($job.Failure)) { throw ('Expected native failure: '+$job.Id+'; actual: '+$failure) }
    }
    else {
        if ($failure -or !$pages.Count) { throw ('Raster path failed: '+$job.Id+'; '+$failure) }
        $hashes=@($pages | Sort-Object Name | ForEach-Object { (Get-FileHash -LiteralPath $_.FullName).Hash })
        if ($null -eq $expectedHashes) { $expectedHashes=$hashes }
        elseif (@(Compare-Object $expectedHashes $hashes).Count) { throw ('Spaced path changed raster output: '+$job.Id) }
    }
    $rows+=@{Id=$job.Id;Outcome='passed';Pages=$pages.Count;ExpectedFailure=$failure}
}
$inputHash=(Get-FileHash -LiteralPath $inputFull).Hash
if ((Get-FileHash -LiteralPath $plain).Hash -ne $inputHash -or (Get-FileHash -LiteralPath $spaced).Hash -ne $inputHash) { throw 'Probe changed its PDF input' }
@{InputSha256=$inputHash;RasterizerSha256=(Get-FileHash -LiteralPath $rasterizer).Hash;
    LibraryBuilds=0;OfficeExports=0;Checks=$rows} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
Write-Output ('Raster wrapper: '+$rows.Count+' checks passed; no build.')
