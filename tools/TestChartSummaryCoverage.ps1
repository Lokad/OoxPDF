# End-to-end optional summary coverage with tiny independent PDFs; never builds.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $InspectorPath,
    [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$inspector=(Resolve-Path -LiteralPath $InspectorPath).Path
$inspectorHash=(Get-FileHash -LiteralPath $inspector).Hash
$summary=Join-Path $PSScriptRoot 'SummarizeChartStructureDeltas.ps1'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path -Parent $PSScriptRoot) ('artifacts/chart-summary-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh chart-summary test directory' }
$scratch=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
# Check semantic summary records, independently of terminal formatting/color.
. (Join-Path $PSScriptRoot 'CompareMath.ps1')
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($summary,[ref]$tokens,[ref]$parseErrors)
if($parseErrors.Count){throw 'Chart summary script must parse'}
foreach($functionName in @('Get-RegionValue','BoundsDelta','SignedDelta','Format-Bounds','Summarize-Kind')) {
    $definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName},$true)
    . ([scriptblock]::Create($definition.Extent.Text))
}
$ByRegion=$true

function Write-ControlPdf([string] $Path, [int] $Pages, [bool] $Moved) {
    $kids=@(for($i=0;$i -lt $Pages;$i++){ ('{0} 0 R' -f (3+2*$i)) }) -join ' '
    $objects=[Collections.Generic.List[string]]::new()
    $objects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objects.Add(('<< /Type /Pages /Count {0} /Kids [{1}] >>' -f $Pages,$kids))
    for($i=0;$i -lt $Pages;$i++) {
        $origin=if($i -eq 0){10}elseif($Moved){400}else{200}
        $end=$origin+100
        $content="$origin $origin m $end $origin l S`n$origin $origin m $origin $end l S`n"
        $objects.Add(('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 800 800] /Resources << >> /Contents {0} 0 R >>' -f (4+2*$i)))
        $objects.Add(('<< /Length {0} >>' -f $content.Length)+"`nstream`n"+$content+'endstream')
    }
    $pdf=[Text.StringBuilder]::new("%PDF-1.4`n")
    $offsets=[Collections.Generic.List[int]]::new()
    for($i=0;$i -lt $objects.Count;$i++) {
        $offsets.Add($pdf.Length)
        [void]$pdf.Append(('{0} 0 obj' -f ($i+1))+"`n"+$objects[$i]+"`nendobj`n")
    }
    $xref=$pdf.Length
    [void]$pdf.Append("xref`n0 "+($objects.Count+1)+"`n0000000000 65535 f `n")
    foreach($offset in $offsets){[void]$pdf.Append($offset.ToString('0000000000')+" 00000 n `n")}
    [void]$pdf.Append(('trailer << /Size {0} /Root 1 0 R >>' -f ($objects.Count+1))+"`nstartxref`n$xref`n%%EOF`n")
    [IO.File]::WriteAllText($Path,$pdf.ToString(),[Text.Encoding]::ASCII)
}

$results=@()
Push-Location $scratch
try {
    foreach($pages in @(1,2)) {
        $id='pptx-ladder-11-summary-'+$pages
        $case=New-Item -ItemType Directory -Path ('visual-cases/cases/'+$id) -Force
        @{id=$id;expected=@{}} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $case.FullName 'case.json')
        $run=New-Item -ItemType Directory -Path ('artifacts/visual/'+$id+'/review') -Force
        foreach($side in @('reference','candidate')) {
            $folder=New-Item -ItemType Directory -Path (Join-Path $run.FullName $side)
            $file=if($side -eq 'reference'){'reference.pdf'}else{'output.pdf'}
            Write-ControlPdf (Join-Path $folder.FullName $file) $pages ($side -eq 'candidate')
        }
        $log=Join-Path $scratch ($pages.ToString()+'-fresh.log')
        & pwsh -NoProfile -File $summary -Case $id -ByRegion -ShowBounds -InspectorPath $inspector -SkipBuild *> $log
        if($LASTEXITCODE -ne 0){throw ('Fresh summary failed: '+$log)}
        $text=Get-Content -LiteralPath $log -Raw
        if($text -notmatch 'freshly inspected all PDF pages'){throw 'Fresh coverage must be explicit'}
        $probeGraphics=@{}
        $hashes=@(foreach($side in @('reference','candidate')) {
            $probe=Join-Path $run.FullName ('probe-'+$side)
            $build=Get-Content -LiteralPath (Join-Path $probe 'pdf-inspect-build-info.json') -Raw | ConvertFrom-Json
            if($build.BuildRequested -ne $false -or $build.OutputSha256 -ne $inspectorHash){throw 'Inspector identity or no-build binding changed'}
            $graphics=Join-Path $probe 'chart-graphics-structures.json'
            $probeGraphics[$side]=@(Get-Content -LiteralPath $graphics -Raw | ConvertFrom-Json)
            $axes=@($probeGraphics[$side] | Where-Object Kind -eq 'AxisPairPlotBoxCandidate')
            if(($axes.PageNumber -join ',') -ne ((1..$pages) -join ',')){throw 'Expected an independent axis pair on every page'}
            @{Path=$graphics;Sha256=(Get-FileHash -LiteralPath $graphics).Hash}
        })
        $row=Summarize-Kind 'AxisPairPlotBoxCandidate' ($pages.ToString()+'/0') $probeGraphics.reference $probeGraphics.candidate 'GRAPHICS'
        $expected=if($pages -eq 1){0}else{200}
        if($row.ReferenceCount -ne 1 -or $row.CandidateCount -ne 1 -or $row.MaxBoundsDelta -ne $expected){throw 'Summary must retain the selected page geometry'}
        $reuseLog=Join-Path $scratch ($pages.ToString()+'-reuse.log')
        & pwsh -NoProfile -File $summary -Case $id -ByRegion -SkipProbe -InspectorPath $inspector -SkipBuild *> $reuseLog
        if($LASTEXITCODE -ne 0 -or (Get-Content -LiteralPath $reuseLog -Raw) -notmatch 'reused reports; PDF page coverage is unverified'){throw 'Reused coverage must not imply a fresh full-document check'}
        foreach($row in $hashes){if((Get-FileHash -LiteralPath $row.Path).Hash -ne $row.Sha256){throw 'Reuse rewrote retained evidence'}}
        $results+=@{Pages=$pages;FreshCoverage='All PDF pages';ReusedCoverage='Unverified';RetainedReportsUnchanged=$true;BuildRequested=$false}
    }
}
finally {Pop-Location}
if((Get-FileHash -LiteralPath $inspector).Hash -ne $inspectorHash){throw 'Inspector changed'}
@{InspectorSha256=$inspectorHash;Controls=$results} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $scratch 'checks.json')
Write-Output 'Chart-summary single/two-page coverage and retained-report checks passed without builds.'
