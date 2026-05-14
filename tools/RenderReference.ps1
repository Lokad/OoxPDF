param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144
)

$ErrorActionPreference = "Stop"

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path
$extension = [System.IO.Path]::GetExtension($inputFull).ToLowerInvariant()

if ($extension -eq ".pptx") {
    $powerPoint = $null
    $presentation = $null
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $presentation = $powerPoint.Presentations.Open($inputFull, $true, $true, $false)
        $width = [int][Math]::Round($presentation.PageSetup.SlideWidth * $Dpi / 72.0)
        $height = [int][Math]::Round($presentation.PageSetup.SlideHeight * $Dpi / 72.0)
        $presentation.Export($outputFull, "PNG", $width, $height)
    }
    finally {
        if ($presentation -ne $null) { $presentation.Close() }
        if ($powerPoint -ne $null) { $powerPoint.Quit() }
    }

    $slides = Get-ChildItem -LiteralPath $outputFull -Filter "*.PNG" | Sort-Object {
        if ($_.BaseName -match '\d+$') { [int]$Matches[0] } else { [int]::MaxValue }
    }, Name
    $index = 1
    foreach ($slide in $slides) {
        Move-Item -LiteralPath $slide.FullName -Destination (Join-Path $outputFull ("page-{0:000}.png" -f $index)) -Force
        $index++
    }
    return
}

if ($extension -eq ".docx") {
    $word = $null
    $document = $null
    $referencePdf = Join-Path $outputFull "reference.pdf"
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open($inputFull, $false, $true)
        $document.ExportAsFixedFormat($referencePdf, 17)
    }
    finally {
        if ($document -ne $null) { $document.Close($false) }
        if ($word -ne $null) { $word.Quit() }
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $outputFull -Dpi $Dpi
    return
}

throw "Unsupported reference input extension '$extension'. Expected .pptx or .docx."
