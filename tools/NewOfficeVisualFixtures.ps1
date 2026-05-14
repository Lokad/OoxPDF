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

    $text = $powerPoint.Presentations.Add($false)
    try {
        $slide = $text.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255
        $title = $slide.Shapes.AddTextbox(1, 72, 72, 576, 72)
        $title.TextFrame.TextRange.Text = "Hello OOXML"
        $title.TextFrame.TextRange.Font.Name = "Arial"
        $title.TextFrame.TextRange.Font.Size = 36
        $title.TextFrame.TextRange.Font.Color.RGB = Rgb 47 128 237

        $body = $slide.Shapes.AddTextbox(1, 72, 180, 576, 144)
        $body.TextFrame.TextRange.Text = "Basic Latin text rendering"
        $body.TextFrame.TextRange.Font.Name = "Arial"
        $body.TextFrame.TextRange.Font.Size = 24
        $body.TextFrame.TextRange.Font.Color.RGB = Rgb 30 30 30

        $text.SaveAs((Join-Path $cases "pptx-text.pptx"), 24)
    }
    finally {
        $text.Close()
    }

    Add-Type -AssemblyName System.Drawing
    $imagePath = Join-Path $env:TEMP ("ooxpdf-image-" + [guid]::NewGuid() + ".png")
    $bitmap = [System.Drawing.Bitmap]::new(160, 90)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([System.Drawing.Color]::FromArgb(47, 128, 237))
            $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(235, 87, 87))
            try {
                $graphics.FillRectangle($brush, 40, 20, 80, 50)
            }
            finally {
                $brush.Dispose()
            }
        }
        finally {
            $graphics.Dispose()
        }
        $bitmap.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }

    $images = $powerPoint.Presentations.Add($false)
    try {
        $slide = $images.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255
        $slide.Shapes.AddPicture($imagePath, $false, $true, 144, 144, 432, 243) | Out-Null
        $images.SaveAs((Join-Path $cases "pptx-images.pptx"), 24)
    }
    finally {
        $images.Close()
        Remove-Item -LiteralPath $imagePath -Force -ErrorAction SilentlyContinue
    }

    $tables = $powerPoint.Presentations.Add($false)
    try {
        $slide = $tables.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255
        $title = $slide.Shapes.AddTextbox(1, 72, 36, 576, 48)
        $title.TextFrame.TextRange.Text = "Quarterly summary"
        $title.TextFrame.TextRange.Font.Name = "Arial"
        $title.TextFrame.TextRange.Font.Size = 28
        $title.TextFrame.TextRange.Font.Color.RGB = Rgb 35 35 35

        $tableShape = $slide.Shapes.AddTable(3, 3, 72, 108, 576, 216)
        $table = $tableShape.Table
        $values = @(
            @("Region", "Q1", "Q2"),
            @("North", "42", "48"),
            @("South", "37", "41")
        )

        for ($row = 1; $row -le 3; $row++) {
            for ($column = 1; $column -le 3; $column++) {
                $cell = $table.Cell($row, $column)
                $cell.Shape.TextFrame.TextRange.Text = $values[$row - 1][$column - 1]
                $cell.Shape.TextFrame.TextRange.Font.Name = "Arial"
                $cell.Shape.TextFrame.TextRange.Font.Size = 18
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = Rgb 30 30 30
                if ($row -eq 1) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 47 128 237
                    $cell.Shape.TextFrame.TextRange.Font.Color.RGB = Rgb 255 255 255
                    $cell.Shape.TextFrame.TextRange.Font.Bold = $true
                }
                elseif (($row + $column) % 2 -eq 0) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 232 244 253
                }
                else {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 250 250 250
                }
            }
        }

        $tables.SaveAs((Join-Path $cases "pptx-table.pptx"), 24)
    }
    finally {
        $tables.Close()
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
