# Exercise the real wrapper and existing native tool; never build or export Office.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory=$true)] [string] $RasterizerPath,
    [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$rasterizer=(Resolve-Path -LiteralPath $RasterizerPath).Path
if (!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/raster-path-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName

function Write-TestPdf([string] $Path, [int] $Pages) {
    $kids=@(for ($index=0; $index -lt $Pages; $index++) { (3+2*$index).ToString()+' 0 R' }) -join ' '
    $bodies=@('<< /Type /Catalog /Pages 2 0 R >>',('<< /Type /Pages /Kids ['+$kids+'] /Count '+$Pages+' >>'))
    $bodies+=@(for ($index=0; $index -lt $Pages; $index++) {
        $contentId=4+2*$index
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 72 72] /Resources << >> /Contents '+$contentId+' 0 R >>'
        $ink=if ($index -eq 0) { '8 8 16 24 re f' } else { '40 32 24 16 re f' }
        ('<< /Length '+$ink.Length+" >>`nstream`n"+$ink+"`nendstream")
    })
    $pdf=[Text.StringBuilder]::new("%PDF-1.7`n")
    $offsets=@(for ($index=0; $index -lt $bodies.Count; $index++) {
        $pdf.Length
        [void]$pdf.Append(($index+1).ToString()+" 0 obj`n"+$bodies[$index]+"`nendobj`n")
    })
    $xref=$pdf.Length
    [void]$pdf.Append("xref`n0 "+($bodies.Count+1)+"`n0000000000 65535 f `n")
    foreach ($offset in $offsets) { [void]$pdf.Append($offset.ToString('0000000000')+" 00000 n `n") }
    [void]$pdf.Append("trailer`n<< /Size "+($bodies.Count+1)+" /Root 1 0 R >>`nstartxref`n"+$xref+"`n%%EOF`n")
    [IO.File]::WriteAllText($Path,$pdf.ToString(),[Text.Encoding]::ASCII)
}

function Get-NamedHashes([string] $Directory, [string] $Filter='page-*.png') {
    $files=Get-ChildItem -LiteralPath $Directory -File -Filter $Filter | Sort-Object Name -CaseSensitive
    $rows=foreach ($file in $files) {
        $file.Name + ' ' + (Get-FileHash -LiteralPath $file.FullName).Hash
    }
    return $rows -join "`n"
}

$twoPages=Join-Path $root 'two-pages.pdf'
$onePage=Join-Path $root 'one-page.pdf'
$spaced=Join-Path $root 'input file.pdf'
Write-TestPdf $twoPages 2
Write-TestPdf $onePage 1
$inputHash=(Get-FileHash -LiteralPath $twoPages).Hash
Copy-Item -LiteralPath $twoPages -Destination $spaced

# The copied wrapper, tool assembly and native loader path all contain spaces.
$ownedRepo=Join-Path $root 'repository with spaces'
$ownedTools=Join-Path $ownedRepo 'tools'
$ownedRuntime=Join-Path $ownedTools "Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$Framework"
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
$invalid=Join-Path $root 'invalid.pdf'
[IO.File]::WriteAllText($invalid,'Owned invalid PDF sentinel')
$rows=@()
$expectedHashes=$null
foreach ($job in @(
    @{Id='plain';Input=$twoPages;Output='plain-raster'},
    @{Id='spaced-input';Input=$spaced;Output='input-raster'},
    @{Id='spaced-output';Input=$twoPages;Output='output raster'},
    @{Id='invalid-pdf';Input=$invalid;Output='invalid-raster';Failure='failed with exit code'},
    @{Id='zero-deadline';Input=$twoPages;Output='timeout-raster';Failure='timed out after 0 seconds';Timeout=0})) {
    $destination=Join-Path $root $job.Output
    $timeout=if ($job.ContainsKey('Timeout')) { $job.Timeout } else { 30 }
    $failure=$null
    try {
        & (Join-Path $ownedTools 'RasterizePdf.ps1') -Framework $Framework -InputPdf $job.Input -OutputDirectory $destination -Dpi 36 -SkipBuild -TimeoutSeconds $timeout *> (Join-Path $root ($job.Id+'.log'))
    }
    catch { $failure=$_.Exception.Message }
    $pages=@(Get-ChildItem -LiteralPath $destination -Filter 'page-*.png' -File)
    if ($job.Failure) {
        if (!$failure -or !$failure.Contains($job.Failure)) { throw ('Expected native failure: '+$job.Id+'; actual: '+$failure) }
        if ($pages.Count) { throw ('Failed fresh raster left consumable pages: '+$job.Id) }
    }
    else {
        if ($failure -or (($pages.Name | Sort-Object) -join ',') -cne 'page-001.png,page-002.png') {
            throw ('Raster path failed: '+$job.Id+'; '+$failure)
        }
        $hashes=Get-NamedHashes $destination
        if ($null -eq $expectedHashes) { $expectedHashes=$hashes }
        elseif ($expectedHashes -cne $hashes) { throw ('Spaced path changed raster output: '+$job.Id) }
    }
    $rows+=@{Id=$job.Id;Outcome='passed';Pages=$pages.Count;ExpectedFailure=$failure}
}
# Swapping contents under the original filenames must change the comparison.
$plainRaster=Join-Path $root 'plain-raster'
$swapped=Join-Path $root 'swapped-raster'
New-Item -ItemType Directory -Path $swapped | Out-Null
Copy-Item -LiteralPath (Join-Path $plainRaster 'page-001.png') -Destination (Join-Path $swapped 'page-002.png')
Copy-Item -LiteralPath (Join-Path $plainRaster 'page-002.png') -Destination (Join-Path $swapped 'page-001.png')
if ($expectedHashes -ceq (Get-NamedHashes $swapped)) { throw 'Expected a distinct image on each fixture page' }
$rows+=@{Id='swapped-pages';Outcome='passed';Pages=2}
$rasterizerDll=Join-Path $ownedRuntime 'Lokad.OoxPdf.PdfiumRasterizer.dll'
# Each 72pt square page produces 1,296 pixels at 36 DPI. Both pages fit
# individually at either budget, so the low bound detects a per-page reset.
foreach ($quota in @(
    @{Id='total-quota-exact';Limit=2592;Exit=0;Names=@('page-001.png','page-002.png')},
    @{Id='total-quota-over';Limit=2591;Exit=1;Names=@('page-001.png')})) {
    $destination=Join-Path $root $quota.Id
    $log=Join-Path $root ($quota.Id+'.log')
    & dotnet $rasterizerDll $twoPages $destination 36 --max-total-pixels $quota.Limit *> $log
    $exitCode=$LASTEXITCODE
    $pageNames=@(Get-ChildItem -LiteralPath $destination -File -Filter '*.png' | Sort-Object Name | ForEach-Object Name)
    if ($exitCode -ne $quota.Exit -or ($pageNames -join ',') -cne ($quota.Names -join ',')) {
        throw ('Unexpected raster quota result: '+$quota.Id)
    }
    if ($quota.Exit -ne 0 -and !(Select-String -LiteralPath $log -SimpleMatch 'PDF total raster pixel count 2592 exceeds the maximum of 2591' -Quiet)) {
        throw 'An aggregate quota refusal must identify the cumulative pixel count'
    }
    $rows+=@{Id=$quota.Id;Outcome='passed';Pages=$pageNames.Count;TotalPixelLimit=$quota.Limit;ExitCode=$exitCode}
}
$reused=Join-Path $root 'shrinking-raster'
& (Join-Path $ownedTools 'RasterizePdf.ps1') -Framework $Framework -InputPdf $twoPages -OutputDirectory $reused -Dpi 36 -SkipBuild *> (Join-Path $root 'two-pages.log')
if (@(Get-ChildItem -LiteralPath $reused -Filter '*.png' -File).Count -ne 2) { throw 'Expected two initial raster pages' }
$before=Get-NamedHashes $reused '*'
foreach ($job in @(@{Id='shrinking-reuse';Input=$onePage;Timeout=30},
                  @{Id='failed-reuse';Input=$invalid;Timeout=30},
                  @{Id='timeout-reuse';Input=$twoPages;Timeout=0})) {
    $failure=$null
    try { & (Join-Path $ownedTools 'RasterizePdf.ps1') -Framework $Framework -InputPdf $job.Input -OutputDirectory $reused -Dpi 36 -SkipBuild -TimeoutSeconds $job.Timeout *> (Join-Path $root ($job.Id+'.log')) }
    catch { $failure=$_.Exception.Message }
    if (!$failure -or !$failure.Contains('choose a fresh directory')) { throw ('Reused raster did not refuse: '+$job.Id) }
    $rows+=@{Id=$job.Id;Outcome='passed';Pages=2;ExpectedFailure=$failure}
}
& dotnet $rasterizerDll $onePage $reused 36 *> (Join-Path $root 'direct-reuse.log')
if ($LASTEXITCODE -ne 2) { throw 'Direct rasterizer must also refuse stale output' }
$after=Get-NamedHashes $reused '*'
if ($before -cne $after) { throw 'Refused raster reuse modified caller files' }
$rows+=@{Id='direct-reuse';Outcome='passed';Pages=2;ExpectedFailure='Exit 2; existing pages preserved'}
$fresh=Join-Path $root 'fresh-shorter-raster'
& (Join-Path $ownedTools 'RasterizePdf.ps1') -Framework $Framework -InputPdf $onePage -OutputDirectory $fresh -Dpi 36 -SkipBuild *> (Join-Path $root 'fresh-shorter.log')
if (@(Get-ChildItem -LiteralPath $fresh -Filter '*.png' -File).Count -ne 1) { throw 'Fresh shorter PDF must produce exactly one page' }
$rows+=@{Id='fresh-shorter';Outcome='passed';Pages=1}
$bound=Join-Path $root 'explicit-frozen-tool'
$otherFramework=if ($Framework -eq 'net11.0') {'net10.0'} else {'net11.0'}
& (Join-Path $ownedTools 'RasterizePdf.ps1') -Framework $otherFramework -RasterizerPath $rasterizerDll -InputPdf $onePage -OutputDirectory $bound -Dpi 36 *> (Join-Path $root 'explicit-frozen-tool.log')
if (@(Get-ChildItem -LiteralPath $bound -Filter '*.png' -File).Count -ne 1) { throw 'Explicit frozen rasterizer did not override framework discovery without building' }
$rows+=@{Id='explicit-frozen-tool';Outcome='passed';Pages=1}
if ((Get-FileHash -LiteralPath $twoPages).Hash -ne $inputHash -or (Get-FileHash -LiteralPath $spaced).Hash -ne $inputHash) { throw 'Probe changed its PDF input' }
@{InputSha256=$inputHash;RasterizerSha256=(Get-FileHash -LiteralPath $rasterizer).Hash;
    LibraryBuilds=0;OfficeExports=0;Checks=$rows} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
Write-Output ('Raster wrapper: '+$rows.Count+' checks passed; no build.')
# Expected child-process refusals must not become the exit code of this suite.
$global:LASTEXITCODE=0
