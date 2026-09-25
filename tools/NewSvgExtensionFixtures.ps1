# Regenerates the SVG-extension visual fixture (Office-authored via PowerPoint COM).
# Run on an Office setup, then commit tests/Lokad.OoxPdf.Tests/Cases/pptx-svg-extensions.pptx.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null

function Release-ComObject($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

$svgDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ooxpdf-svg-" + [guid]::NewGuid())
New-Item -ItemType Directory -Force -Path $svgDir | Out-Null

$hgrad = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/></linearGradient></defs><path d="M 5 5 L 95 5 L 95 95 L 5 95 Z" fill="url(#g)"/></svg>';
$vgrad = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#00FF00"/><stop offset="1" stop-color="#FFFF00"/></linearGradient></defs><path d="M 5 5 L 95 5 L 95 95 L 5 95 Z" fill="url(#g)"/></svg>';
$dgrad = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#00FFFF"/><stop offset="1" stop-color="#FF00FF"/></linearGradient></defs><path d="M 5 5 L 95 5 L 95 95 L 5 95 Z" fill="url(#g)"/></svg>';
$rot = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><g transform="rotate(30 50 50)"><path d="M 20 20 L 80 20 L 80 80 L 20 80 Z" fill="#00AA00"/></g></svg>';
$style = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><path d="M 5 5 L 95 5 L 95 95 L 5 95 Z" fill="#0000FF" style="fill:#FF8800;fill-opacity:0.6"/></svg>';
$reflect = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><linearGradient id="g" spreadMethod="reflect" x1="0" y1="0" x2="0.25" y2="0"><stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#000000"/></linearGradient></defs><path d="M 5 5 L 95 5 L 95 95 L 5 95 Z" fill="url(#g)"/></svg>';

$files = [ordered]@{ "svg-hgrad.svg" = $hgrad; "svg-vgrad.svg" = $vgrad; "svg-dgrad.svg" = $dgrad; "svg-rot.svg" = $rot; "svg-style.svg" = $style; "svg-reflect.svg" = $reflect }
foreach ($entry in $files.GetEnumerator()) {
    Set-Content -LiteralPath (Join-Path $svgDir $entry.Key) -Value $entry.Value -Encoding UTF8
}

$powerPoint = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add($false)
    try {
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = 16777215
        $slots = @(
            @("svg-hgrad.svg", 72, 72), @("svg-vgrad.svg", 360, 72), @("svg-dgrad.svg", 648, 72),
            @("svg-rot.svg", 72, 306), @("svg-style.svg", 360, 306), @("svg-reflect.svg", 648, 306)
        )
        foreach ($slot in $slots) {
            $path = Join-Path $svgDir $slot[0]
            $slide.Shapes.AddPicture($path, $false, $true, $slot[1], $slot[2], 240, 180) | Out-Null
        }

        $presentation.SaveAs((Join-Path $cases "pptx-svg-extensions.pptx"), 24)
    }
    finally {
        $presentation.Close()
        Release-ComObject $presentation
    }
}
finally {
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
        Release-ComObject $powerPoint
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-ChildItem -LiteralPath (Join-Path $cases "pptx-svg-extensions.pptx") | Select-Object FullName, Length
