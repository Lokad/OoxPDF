# Office-authored SVG stroke/radial probe, generated under the reference supervisor.
param(
    [ValidateSet('stroke-radial', 'focal-controls', 'radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls')]
    [string] $ProbeSet = 'stroke-radial',
    [string] $OutputPath,
    [string] $OutputDirectory,
    [int] $Dpi = 144,
    [int] $TimeoutSeconds = 120,
    [string] $InputPath,
    [string] $WorkDirectory,
    [string] $ProgressLog,
    [string] $StatusPath
)
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = if ($ProbeSet -eq 'stroke-radial') { 'tests/Lokad.OoxPdf.Tests/Cases/pptx-svg-stroke-radial.pptx' } else { "artifacts/svg-$ProbeSet/pptx-svg-$ProbeSet.pptx" }
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = if ($ProbeSet -eq 'stroke-radial') { 'artifacts/svg-stroke-radial-reference' } else { "artifacts/svg-$ProbeSet-reference" }
}
if ([string]::IsNullOrWhiteSpace($WorkDirectory)) {
    $outputFull = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $repoRoot $OutputDirectory }))
    # The supervisor starts a fresh worker; pass the probe choice through a small
    # wrapper rather than relying on this process's parameter state.
    $workerScript = $PSCommandPath
    if ($ProbeSet -ne 'stroke-radial') {
        New-Item -ItemType Directory -Force -Path $outputFull | Out-Null
        $workerScript = Join-Path $outputFull 'probe-worker.ps1'
        $quotedGenerator = $PSCommandPath.Replace("'", "''")
        $workerTemplate = @'
param([string] $InputPath, [string] $WorkDirectory, [int] $Dpi, [string] $ProgressLog, [string] $StatusPath)
& 'GENERATOR_PATH' -ProbeSet PROBE_SET -InputPath $InputPath -WorkDirectory $WorkDirectory -Dpi $Dpi -ProgressLog $ProgressLog -StatusPath $StatusPath
'@
        [IO.File]::WriteAllText($workerScript, $workerTemplate.Replace('GENERATOR_PATH', $quotedGenerator).Replace('PROBE_SET', $ProbeSet))
    }
    & (Join-Path $PSScriptRoot "RenderReference.ps1") -InputPath (Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/pptx-blank.pptx") -OutputDirectory $outputFull -Dpi $Dpi -TimeoutSeconds $TimeoutSeconds -WorkerScript $workerScript
    $fixtureFull = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path $repoRoot $OutputPath }))
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $fixtureFull) | Out-Null
    Copy-Item -LiteralPath (Join-Path $outputFull "fixture.pptx") -Destination $fixtureFull -Force
    Write-Host "SVG stroke/radial fixture: $fixtureFull; Office reference: $outputFull"
    return
}
function Stage([string] $Name) {
    Add-Content -LiteralPath $ProgressLog -Value ("stage:$Name " + [DateTime]::UtcNow.ToString("O"))
}
function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
}
$svgInputs = if ($ProbeSet -eq 'focal-controls') {
    $controls = [ordered]@{}
    $focalVariants = [ordered]@{
        'numeric' = 'fx="25" fy="30"'
        'percent' = 'fx="25%" fy="30%"'
        'px' = 'fx="25px" fy="30px"'
        'frzero' = 'fx="25" fy="30" fr="0"'
        'outside' = 'fx="99" fy="30"'
        'centered' = 'fx="50" fy="50"'
    }
    foreach ($variant in $focalVariants.GetEnumerator()) {
        $controls[$variant.Key] = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="50" r="40" ' + $variant.Value + '><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs><path d="M5 5H95V95H5Z" fill="url(#g)"/></svg>'
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-transform-controls') {
    $controls = [ordered]@{}
    foreach ($name in @('uniform','picture-stretch','group-stretch','group-uniform','group-rotate','group-shear','non-scaling-uniform','non-scaling-stretch','dashed-stretch','square-caps-stretch')) {
        $width = if ($name -in @('picture-stretch','dashed-stretch','square-caps-stretch')) { 100 } else { 200 }
        $transform = switch ($name) {
            'group-stretch' { 'transform="scale(2,1)"' }
            'group-uniform' { 'transform="scale(2)"' }
            'group-rotate' { 'transform="translate(35,15) rotate(20)"' }
            'group-shear' { 'transform="matrix(1,0,0.5,1,0,0)"' }
            'non-scaling-uniform' { 'transform="scale(2)"' }
            'non-scaling-stretch' { 'transform="scale(2,1)"' }
            default { '' }
        }
        $effect = if ($name.StartsWith('non-scaling')) { 'vector-effect="non-scaling-stroke"' } else { '' }
        $dash = if ($name -eq 'dashed-stretch') { 'stroke-dasharray="8 4" stroke-dashoffset="3"' } else { '' }
        $cap = if ($name -eq 'square-caps-stretch') { 'square' } else { 'round' }
        $path = if ($name -in @('group-uniform','non-scaling-uniform')) { 'M10 10H70 M15 10V40 M25 40L60 10' } else { 'M15 25H85 M25 15V85 M35 80L75 25' }
        $controls[$name] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $width 100`"><g $transform><path d=`"$path`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"$cap`" stroke-linejoin=`"round`" $effect $dash/></g></svg>"
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-element-controls') {
    $controls = [ordered]@{}
    $transforms = [ordered]@{
        'horizontal-stretch' = 'scale(2,1)'
        'vertical-stretch' = 'scale(1,2)'
        'horizontal-contract' = 'scale(0.5,1)'
        'unequal-stretch' = 'scale(3,2)'
        'rotated-stretch' = 'translate(40,5) rotate(20) scale(2,1)'
        'small-shear' = 'matrix(1,0,0.5,1,0,0)'
        'unit-shear' = 'matrix(1,0,1,1,0,0)'
        'large-shear' = 'matrix(1,0,2,1,0,0)'
        'reflected-stretch' = 'translate(180,0) scale(-2,1)'
    }
    foreach ($item in $transforms.GetEnumerator()) {
        $controls[$item.Key] = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><g transform="' + $item.Value + '"><path d="M10 20H50 M20 10V35 M25 35L45 10" fill="none" stroke="#0000FF" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/></g></svg>'
    }
    $controls['filled-stretch'] = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><g transform="scale(2,1)"><path d="M10 20H50V35H10Z" fill="#00FF00" stroke="#0000FF" stroke-width="6" stroke-linejoin="round"/></g></svg>'
    $controls['filled-shear'] = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><g transform="matrix(1,0,0.5,1,0,0)"><path d="M10 20H50V35H10Z" fill="#00FF00" stroke="#0000FF" stroke-width="6" stroke-linejoin="round"/></g></svg>'
    $controls
} elseif ($ProbeSet -eq 'radial-spread-controls') {
    $controls = [ordered]@{}
    foreach ($spread in @('repeat', 'reflect')) {
        foreach ($variant in @('two-stops', 'three-stops', 'shifted', 'small-radius', 'scaled-units')) {
            $width = if ($variant -eq 'scaled-units') { 1000 } else { 100 }
            $height = $width / 2
            $factor = $width / 100
            $radius = if ($variant -eq 'small-radius') { 5 * $factor } else { 20 * $factor }
            $cx = if ($variant -eq 'shifted') { 30 * $factor } else { $width / 2 }
            $cy = if ($variant -eq 'shifted') { 15 * $factor } else { $height / 2 }
            $stops = if ($variant -eq 'three-stops') { '<stop offset="0" stop-color="#FF0000"/><stop offset="0.4" stop-color="#00FF00"/><stop offset="1" stop-color="#0000FF"/>' } else { '<stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/>' }
            $controls[$spread + '-' + $variant] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $width $height`"><defs><radialGradient id=`"g`" gradientUnits=`"userSpaceOnUse`" cx=`"$cx`" cy=`"$cy`" r=`"$radius`" spreadMethod=`"$spread`">$stops</radialGradient></defs><path d=`"M0 0H${width}V${height}H0Z`" fill=`"url(#g)`"/></svg>"
        }
    }
    $controls
} elseif ($ProbeSet -eq 'radial-paint-controls') {
    $controls = [ordered]@{}
    foreach ($name in @('two-stops', 'scaled-units', 'three-stops', 'extended-stops', 'opacity', 'small-radius', 'repeat', 'reflect', 'hard-stop')) {
        $width = if ($name -eq 'scaled-units') { 1000 } else { 100 }
        $height = $width / 2
        $radius = if ($name -eq 'small-radius') { $width / 10 } else { $width / 4 }
        $spread = if ($name -in @('repeat', 'reflect')) { "spreadMethod=`"$name`"" } else { '' }
        $opacity = if ($name -eq 'opacity') { 'opacity="0.5"' } else { '' }
        $stops = switch ($name) {
            'three-stops' { '<stop offset="0" stop-color="#FF0000"/><stop offset="0.4" stop-color="#00FF00"/><stop offset="1" stop-color="#0000FF"/>' }
            'extended-stops' { '<stop offset="0.2" stop-color="#FF0000"/><stop offset="0.8" stop-color="#0000FF"/>' }
            'hard-stop' { '<stop offset="0" stop-color="#FF0000"/><stop offset="0.5" stop-color="#FF0000"/><stop offset="0.5" stop-color="#0000FF"/><stop offset="1" stop-color="#0000FF"/>' }
            default { '<stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/>' }
        }
        $cx = $width / 2
        $cy = $height / 2
        $controls[$name] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $width $height`"><defs><radialGradient id=`"g`" gradientUnits=`"userSpaceOnUse`" cx=`"$cx`" cy=`"$cy`" r=`"$radius`" $spread>$stops</radialGradient></defs><path d=`"M0 0H${width}V${height}H0Z`" fill=`"url(#g)`" $opacity/></svg>"
    }
    $controls
} else {
    [ordered]@{
    'radial-pad' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="50" r="30"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs><path d="M5 5H95V95H5Z" fill="url(#g)"/></svg>'
    'radial-focus' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="50" r="50" fx="25" fy="30"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></radialGradient></defs><path d="M5 5H95V95H5Z" fill="url(#g)"/></svg>'
    'radial-user' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="50" cy="50" r="30"><stop offset="0" stop-color="#00FF00"/><stop offset="1" stop-color="#000000"/></radialGradient></defs><path d="M5 5H95V95H5Z" fill="url(#g)"/></svg>'
    'stroke-caps' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><path d="M15 25H85" fill="none" stroke="#008000" stroke-width="6" stroke-linecap="round"/><path d="M15 50H85" fill="none" stroke="#FF0000" stroke-width="6" stroke-linecap="square"/><path d="M15 75H85" fill="none" stroke="#0000FF" stroke-width="6" stroke-linecap="butt"/></svg>'
    'stroke-vector' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><g transform="scale(2)"><path d="M10 15H40" fill="none" stroke="#0000FF" stroke-width="4"/><path d="M10 35H40" fill="none" stroke="#FF0000" stroke-width="4" vector-effect="non-scaling-stroke"/></g></svg>'
    'stroke-dash' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><path d="M10 25H90" fill="none" stroke="#000000" stroke-width="4" stroke-dasharray="8 4" stroke-dashoffset="3"/><path d="M15 80L50 40L85 80" fill="none" stroke="#FF8000" stroke-width="5" stroke-linejoin="miter" stroke-miterlimit="2"/></svg>'
}
}
$svgRoot = Join-Path $WorkDirectory '_svg'
New-Item -ItemType Directory -Force -Path $svgRoot | Out-Null
$app = $null
$presentation = $null
$stage = 'activation'
try {
    Stage $stage
    $app = New-Object -ComObject PowerPoint.Application
    $version = [string]$app.Version
    $app.DisplayAlerts = 1
    $stage = 'create'; Stage $stage
    $presentation = $app.Presentations.Add($false)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $slide = $null
    if ($ProbeSet -notin @('radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls')) {
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = 16777215
    }
    $index = 0
    foreach ($item in $svgInputs.GetEnumerator()) {
        $svgPath = Join-Path $svgRoot ($item.Key + '.svg')
        Set-Content -LiteralPath $svgPath -Value $item.Value -Encoding utf8
        if ($ProbeSet -in @('radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls')) {
            $slide = $presentation.Slides.Add($index + 1, 12)
            $slide.Background.Fill.ForeColor.RGB = 16777215
            $slide.Shapes.AddPicture($svgPath, $false, $true, 72, 72, 432, 216) | Out-Null
        } else {
            $left = 72 + 288 * ($index % 3)
            $top = 72 + 234 * [Math]::Floor($index / 3)
            $slide.Shapes.AddPicture($svgPath, $false, $true, $left, $top, 240, 180) | Out-Null
        }
        $index++
    }
    $stage = 'export'; Stage $stage
    $presentation.SaveAs((Join-Path $WorkDirectory 'fixture.pptx'), 24)
    $presentation.SaveAs((Join-Path $WorkDirectory 'reference.pdf'), 32)
} catch {
    [ordered]@{Status='export-failed';Stage=$stage;OfficeApp='PowerPoint';OfficeVersion='';ExportSettings='';Error=$_.Exception.Message} | ConvertTo-Json | Set-Content -LiteralPath $StatusPath -Encoding utf8
    throw
} finally {
    $stage = 'cleanup'; Stage $stage
    try { if ($presentation) { $presentation.Close() } }
    finally {
        try { if ($app) { $app.Quit() } }
        finally { Release-ComObject $presentation; Release-ComObject $app; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    }
}
$stage = 'rasterize'; Stage $stage
& (Join-Path $PSScriptRoot 'RasterizePdf.ps1') -InputPdf (Join-Path $WorkDirectory 'reference.pdf') -OutputDirectory $WorkDirectory -Dpi $Dpi
$stage = 'done'; Stage $stage
[ordered]@{Status='ok';Stage=$stage;OfficeApp='PowerPoint';OfficeVersion=$version;ExportSettings="Generated $index SVG pictures ($ProbeSet) with AddPicture; 960x540 slides; SaveAs PPTX(24) then PDF(32)";Error=''} | ConvertTo-Json | Set-Content -LiteralPath $StatusPath -Encoding utf8
