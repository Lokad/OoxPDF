$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null

function Rgb($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

$powerPoint = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application

    $blank = $powerPoint.Presentations.Add($false)
    try {
        $blank.Slides.Add(1, 12) | Out-Null
        $blank.SaveAs((Join-Path $cases "pptx-blank.pptx"), 24)
    }
    finally {
        $blank.Close()
    }

    $shapes = $powerPoint.Presentations.Add($false)
    try {
        $slide = $shapes.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 248 248 248

        $rect = $slide.Shapes.AddShape(1, 72, 72, 216, 108)
        $rect.Fill.ForeColor.RGB = Rgb 47 128 237
        $rect.Line.ForeColor.RGB = Rgb 27 79 156
        $rect.Line.Weight = 2

        $oval = $slide.Shapes.AddShape(9, 360, 108, 144, 108)
        $oval.Fill.ForeColor.RGB = Rgb 39 174 96
        $oval.Line.ForeColor.RGB = Rgb 20 90 50
        $oval.Line.Weight = 2
        $oval.Rotation = 15

        $line = $slide.Shapes.AddLine(72, 324, 504, 396)
        $line.Line.ForeColor.RGB = Rgb 235 87 87
        $line.Line.Weight = 3

        $shapes.SaveAs((Join-Path $cases "pptx-shapes.pptx"), 24)
    }
    finally {
        $shapes.Close()
    }
}
finally {
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
    }
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    try {
        $doc.PageSetup.PageWidth = 612
        $doc.PageSetup.PageHeight = 792
        $doc.SaveAs2((Join-Path $cases "docx-blank.docx"), 16)
    }
    finally {
        $doc.Close($false)
    }
}
finally {
    if ($word -ne $null) {
        $word.Quit()
    }
}

Get-ChildItem -LiteralPath $cases -Include "*.pptx", "*.docx" -Recurse | Select-Object FullName, Length
