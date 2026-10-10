# Office-authored SVG stroke/radial probe, generated under the reference supervisor.
param(
    [ValidateSet('stroke-radial', 'focal-controls', 'radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls', 'stroke-filled-controls', 'stroke-dash-controls', 'stroke-dash-phase-controls', 'stroke-square-dash-controls', 'stroke-round-dash-controls', 'stroke-short-dash-controls', 'radial-opacity-controls', 'linear-stroke-controls')]
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
} elseif ($ProbeSet -eq 'stroke-filled-controls') {
    $controls = [ordered]@{}
    foreach ($name in @('opaque','evenodd','fill-alpha','stroke-alpha','combined-alpha','style-alpha','square-caps','non-scaling')) {
        $path = switch ($name) {
            'evenodd' { 'M15 15H85V85H15Z M35 35H65V65H35Z' }
            'square-caps' { 'M15 25H85V75L15 75' }
            default { 'M15 25H85V75H15Z' }
        }
        $rule = if ($name -eq 'evenodd') { 'fill-rule="evenodd"' } else { '' }
        $alpha = switch ($name) {
            'fill-alpha' { 'fill-opacity="0.5"' }
            'stroke-alpha' { 'stroke-opacity="0.5"' }
            'combined-alpha' { 'opacity="0.5"' }
            'style-alpha' { 'style="fill-opacity:0.5;stroke-opacity:0.25"' }
            default { '' }
        }
        $cap = if ($name -eq 'square-caps') { 'square' } else { 'round' }
        $effect = if ($name -eq 'non-scaling') { 'vector-effect="non-scaling-stroke"' } else { '' }
        $controls[$name] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 100 100`"><path d=`"$path`" fill=`"#00FF00`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"$cap`" stroke-linejoin=`"round`" $rule $alpha $effect/></svg>"
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
} elseif ($ProbeSet -eq 'stroke-dash-controls') {
    $controls = [ordered]@{}
    foreach ($mapping in @('uniform','stretch')) {
        $width = if ($mapping -eq 'uniform') { 200 } else { 100 }
        foreach ($cap in @('butt','round','square')) {
            foreach ($phase in @(0,1,3)) {
                $key = "$mapping-$cap-$phase"
                $controls[$key] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $width 100`"><path d=`"M15 25H85 M25 15V85 M35 80L75 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"$cap`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$phase`"/></svg>"
            }
        }
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-dash-phase-controls') {
    $controls = [ordered]@{}
    foreach ($mapping in @('uniform','stretch')) {
        $width = if ($mapping -eq 'uniform') { 200 } else { 100 }
        foreach ($strokeWidth in @(3,6,9)) {
            foreach ($phase in @(-1,1,3)) {
                $key = "$mapping-$strokeWidth-$phase"
                $controls[$key] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $width 100`"><path d=`"M15 25H85 M25 15V85 M35 80L75 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"$strokeWidth`" stroke-linecap=`"butt`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$phase`"/></svg>"
            }
        }
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-square-dash-controls') {
    $controls = [ordered]@{}
    foreach ($mapping in @('uniform','stretch')) {
        $viewWidth = if ($mapping -eq 'uniform') { 200 } else { 100 }
        foreach ($width in @(3,6)) {
            foreach ($phase in @(-1,0,1,3)) {
                $key = "$mapping-$width-$phase"
                $controls[$key] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`"><path d=`"M15 25H85 M25 15V85 M35 80L75 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"$width`" stroke-linecap=`"square`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$phase`"/></svg>"
            }
        }
        foreach ($probe in @(
            @{Name='equal';Width=6;Dash='6 4'},
            @{Name='odd';Width=3;Dash='12 8 10'},
            @{Name='multiple';Width=6;Dash='10 4 8 6'},
            @{Name='zero';Width=6;Dash='0 4'},
            @{Name='short';Width=6;Dash='4 12'}
        )) {
            $key = "$mapping-$($probe.Name)"
            $controls[$key] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`"><path d=`"M15 25H85 M25 15V85 M35 80L75 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"$($probe.Width)`" stroke-linecap=`"square`" stroke-dasharray=`"$($probe.Dash)`"/></svg>"
        }
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-round-dash-controls') {
    $controls = [ordered]@{}
    foreach ($mapping in @('uniform','stretch')) {
        $viewWidth = if ($mapping -eq 'uniform') { 200 } else { 100 }
        foreach ($probe in @(
            @{Name='regular';Length=70;Phase=0},
            @{Name='partial-first';Length=70;Phase=1},
            @{Name='initial-gap';Length=70;Phase=1.5},
            @{Name='negative-phase';Length=70;Phase=-1},
            @{Name='single-dash';Length=3;Phase=0},
            @{Name='empty';Length=2;Phase=1.5},
            @{Name='partial-last';Length=64;Phase=0},
            @{Name='last-gap';Length=70;Phase=0}
        )) {
            $end=15+$probe.Length
            $key="$mapping-$($probe.Name)"
            $controls[$key]="<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`"><path d=`"M15 25H$end M25 15V$end M35 80L$end 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"round`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$($probe.Phase)`"/></svg>"
        }
    }
    foreach ($mapping in @('uniform','stretch')) {
        $viewWidth=if($mapping -eq 'uniform'){200}else{100}
        foreach($probe in @(
            @{Name='cycle-end-draw';Length=68;Phase=0},
            @{Name='cycle-end-gap';Length=72;Phase=0},
            @{Name='cycle-start-gap';Length=63;Phase=1.5},
            @{Name='cycle-partial';Length=65;Phase=1.5}
        )){
            $end=15+$probe.Length
            $key="$mapping-$($probe.Name)"
            $controls[$key]="<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`"><path d=`"M15 25H$end M25 15V$end M35 80L$end 25`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"round`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$($probe.Phase)`"/></svg>"
        }
    }
    foreach ($mapping in @('uniform','stretch')) {
        $viewWidth=if($mapping -eq 'uniform'){200}else{100}
        foreach($probe in @(
            @{Name='endpoint-neither';Width=3;Length=27;Phase=3},
            @{Name='endpoint-first';Width=3;Length=70;Phase=0},
            @{Name='endpoint-last';Width=3;Length=70;Phase=3},
            @{Name='endpoint-cycle';Width=6;Length=72;Phase=0}
        )){
            $end=15+$probe.Length
            $key="$mapping-$($probe.Name)"
            $controls[$key]="<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`"><path d=`"M15 25H$end`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"$($probe.Width)`" stroke-linecap=`"round`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$($probe.Phase)`"/></svg>"
        }
    }
    $controls
} elseif ($ProbeSet -eq 'stroke-short-dash-controls') {
    $controls = [ordered]@{}
    foreach ($width in @(3,6,9)) {
        foreach ($length in @(2,3,4,6,8,12,15)) {
            foreach ($phaseSource in @(0,9)) {
                $phase=($phaseSource / $width).ToString('R',[Globalization.CultureInfo]::InvariantCulture)
                $end=15+$length
                $key="$width-$length-$phaseSource"
                $controls[$key]="<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 200 100`"><path d=`"M15 25H$end`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"$width`" stroke-linecap=`"round`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"$phase`"/></svg>"
            }
        }
    }
    foreach ($mapping in @('uniform','stretch')) {
        $viewWidth=if($mapping -eq 'uniform'){200}else{100}
        foreach($probe in @(
            @{Name='whole-empty';Data='M15 25H17 M25 15V17';Separate=$false},
            @{Name='whole-boundary';Data='M15 25H18 M25 15V18';Separate=$false},
            @{Name='mixed-short-first';Data='M15 25H17 M25 15V85';Separate=$false},
            @{Name='mixed-long-first';Data='M25 15V85 M15 25H17';Separate=$false},
            @{Name='separate-elements';Data='';Separate=$true}
        )){
            $key="$mapping-$($probe.Name)"
            $paths=if($probe.Separate){
                '<path d="M15 25H17" fill="none" stroke="#0000FF" stroke-width="6" stroke-linecap="round" stroke-dasharray="8 4" stroke-dashoffset="1.5"/><path d="M25 15V85" fill="none" stroke="#0000FF" stroke-width="6" stroke-linecap="round" stroke-dasharray="8 4" stroke-dashoffset="1.5"/>'
            }else{
                "<path d=`"$($probe.Data)`" fill=`"none`" stroke=`"#0000FF`" stroke-width=`"6`" stroke-linecap=`"round`" stroke-dasharray=`"8 4`" stroke-dashoffset=`"1.5`"/>"
            }
            $controls[$key]="<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 $viewWidth 100`">$paths</svg>"
        }
    }
    $controls
} elseif ($ProbeSet -eq 'linear-stroke-controls') {
    # RV07-E10: reproducible public opaque gradient-stroke case.
    [ordered]@{ 'linear-gradient-stroke' = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50" ><defs><linearGradient id="stroke" ><stop offset="0" stop-color="#FF0000" stop-opacity="1"/><stop offset="0.45" stop-color="#00AA88" stop-opacity="1"/><stop offset="1" stop-color="#0000FF" stop-opacity="1"/></linearGradient></defs><path d="M15 10H85V40H15Z"  fill="none" stroke="url(#stroke)" stroke-width="5" /></svg>' }
} elseif ($ProbeSet -eq 'radial-opacity-controls') {
    $controls = [ordered]@{}
    foreach ($spread in @('pad', 'repeat', 'reflect')) {
        foreach ($variant in @('quarter', 'half', 'three-quarters', 'fill-opacity', 'combined', 'style')) {
            $opacity = switch ($variant) { 'quarter' { '0.25' } 'three-quarters' { '0.75' } default { '0.5' } }
            $attributes = switch ($variant) {
                'fill-opacity' { 'fill-opacity="0.5"' }
                'combined' { 'opacity="0.5" fill-opacity="0.5"' }
                'style' { 'style="opacity:0.5;fill-opacity:0.5"' }
                default { "opacity=`"$opacity`"" }
            }
            $controls[$spread + '-' + $variant] = "<svg xmlns=`"http://www.w3.org/2000/svg`" viewBox=`"0 0 100 50`"><defs><radialGradient id=`"g`" gradientUnits=`"userSpaceOnUse`" cx=`"50`" cy=`"25`" r=`"20`" spreadMethod=`"$spread`"><stop offset=`"0`" stop-color=`"#FF0000`"/><stop offset=`"0.4`" stop-color=`"#00FF00`"/><stop offset=`"1`" stop-color=`"#0000FF`"/></radialGradient></defs><path d=`"M0 0H100V50H0Z`" fill=`"url(#g)`" $attributes/></svg>"
        }
    }
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
    if ($ProbeSet -notin @('radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls', 'stroke-filled-controls', 'stroke-dash-controls', 'stroke-dash-phase-controls', 'stroke-square-dash-controls', 'stroke-round-dash-controls', 'stroke-short-dash-controls', 'radial-opacity-controls', 'linear-stroke-controls')) {
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = 16777215
    }
    $index = 0
    foreach ($item in $svgInputs.GetEnumerator()) {
        $svgPath = Join-Path $svgRoot ($item.Key + '.svg')
        Set-Content -LiteralPath $svgPath -Value $item.Value -Encoding utf8
        if ($ProbeSet -in @('radial-paint-controls', 'radial-spread-controls', 'stroke-transform-controls', 'stroke-element-controls', 'stroke-filled-controls', 'stroke-dash-controls', 'stroke-dash-phase-controls', 'stroke-square-dash-controls', 'stroke-round-dash-controls', 'stroke-short-dash-controls', 'radial-opacity-controls', 'linear-stroke-controls')) {
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
    if ($ProbeSet -in @('radial-opacity-controls', 'linear-stroke-controls')) {
        # Keep PNG preview exports separate from the PDF raster pages. The
        # two Office outputs disagree for some SVG opacity constructs.
        $previewRoot = Join-Path $WorkDirectory 'office-preview'
        New-Item -ItemType Directory -Force -Path $previewRoot | Out-Null
        $pixelWidth = [int][Math]::Round($presentation.PageSetup.SlideWidth * $Dpi / 72)
        $pixelHeight = [int][Math]::Round($presentation.PageSetup.SlideHeight * $Dpi / 72)
        foreach ($entry in $presentation.Slides) {
            $entry.Export((Join-Path $previewRoot ("page-{0:D3}.png" -f [int]$entry.SlideIndex)), 'PNG', $pixelWidth, $pixelHeight)
            Release-ComObject $entry
        }
    }
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
