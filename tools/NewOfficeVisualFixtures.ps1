$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Rgb($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Release-ComObject($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function New-MultilineTableFixtureFromWrapped($sourcePath, $targetPath) {
    $sourceFullPath = (Resolve-Path -LiteralPath $sourcePath).Path
    $targetFullPath = [System.IO.Path]::GetFullPath($targetPath)
    $temporaryPath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($targetFullPath),
        ([System.IO.Path]::GetFileNameWithoutExtension($targetFullPath) + ".tmp.pptx"))

    if (Test-Path -LiteralPath $temporaryPath) {
        Remove-Item -LiteralPath $temporaryPath -Force
    }

    $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($sourceFullPath)
    $targetZip = [System.IO.Compression.ZipFile]::Open($temporaryPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $sourceZip.Entries) {
            $targetEntry = $targetZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            if ($entry.FullName -eq "ppt/slides/slide1.xml") {
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try {
                    $xml = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }

                $xml = [System.Text.RegularExpressions.Regex]::Replace(
                    $xml,
                    "<a:t>R(?<row>\d{2})C(?<column>\d{2})[^<]*</a:t>",
                    [System.Text.RegularExpressions.MatchEvaluator]{
                        param($match)

                        $row = $match.Groups["row"].Value
                        $column = [int]$match.Groups["column"].Value
                        $tag = "R{0}C{1:00}" -f $row, $column
                        if ($column -eq 4) {
                            $text = "$tag replenishment policy keeps enough neutral planning words to wrap inside the cell"
                        }
                        elseif ($column -eq 5) {
                            $text = "$tag service target and lead time note wraps across the cell"
                        }
                        else {
                            $text = "$tag compact planning text wraps"
                        }

                        return "<a:t>$([System.Security.SecurityElement]::Escape($text))</a:t>"
                    })

                $writer = [System.IO.StreamWriter]::new($targetEntry.Open())
                try {
                    $writer.Write($xml)
                }
                finally {
                    $writer.Dispose()
                }
            }
            else {
                $sourceStream = $entry.Open()
                $targetStream = $targetEntry.Open()
                try {
                    $sourceStream.CopyTo($targetStream)
                }
                finally {
                    $targetStream.Dispose()
                    $sourceStream.Dispose()
                }
            }
        }
    }
    finally {
        $targetZip.Dispose()
        $sourceZip.Dispose()
    }

    Move-Item -LiteralPath $temporaryPath -Destination $targetFullPath -Force
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

    $trailingEmphasis = $powerPoint.Presentations.Add($false)
    try {
        $slide = $trailingEmphasis.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $box = $slide.Shapes.AddTextbox(1, 72, 72, 345, 110)
        $box.TextFrame.MarginLeft = 0
        $box.TextFrame.MarginRight = 0
        $box.TextFrame.MarginTop = 0
        $box.TextFrame.MarginBottom = 0
        $box.TextFrame.WordWrap = -1
        $box.TextFrame.AutoSize = 0

        $textRange = $box.TextFrame.TextRange
        $textRange.Text = "Quality decisions depend on careful operational planning and reliable daily execution."
        $textRange.Font.Name = "Cambria Math"
        $textRange.Font.Size = 14
        $textRange.Font.Color.RGB = Rgb 30 30 30

        $emphasisStart = $textRange.Text.IndexOf("planning") + 1
        $emphasis = $textRange.Characters($emphasisStart, "planning".Length)
        $emphasis.Font.Bold = $true
        $emphasis.Font.Color.RGB = Rgb 30 30 30

        $trailingEmphasis.SaveAs((Join-Path $cases "pptx-ladder-04-typography-trailing-emphasis-probe.pptx"), 24)
    }
    finally {
        $trailingEmphasis.Close()
    }

    $corporate = $powerPoint.Presentations.Add($false)
    try {
        $master = $corporate.SlideMaster
        $master.Background.Fill.ForeColor.RGB = Rgb 245 247 250

        $band = $master.Shapes.AddShape(1, 0, 0, 720, 84)
        $band.Fill.ForeColor.RGB = Rgb 31 78 121
        $band.Line.Visible = 0

        $rule = $master.Shapes.AddShape(1, 0, 84, 720, 6)
        $rule.Fill.ForeColor.RGB = Rgb 112 173 71
        $rule.Line.Visible = 0

        $footer = $master.Shapes.AddTextbox(1, 54, 488, 612, 24)
        $footer.TextFrame.TextRange.Text = "Quarterly operating review"
        $footer.TextFrame.TextRange.Font.Name = "Aptos"
        $footer.TextFrame.TextRange.Font.Size = 10
        $footer.TextFrame.TextRange.Font.Color.RGB = Rgb 95 95 95

        $layout = $master.CustomLayouts.Item(1)
        $slide = $corporate.Slides.AddSlide(1, $layout)
        $title = $slide.Shapes.AddTextbox(1, 54, 28, 612, 42)
        $title.TextFrame.TextRange.Text = "Regional growth dashboard"
        $title.TextFrame.TextRange.Font.Name = "Aptos Display"
        $title.TextFrame.TextRange.Font.Size = 26
        $title.TextFrame.TextRange.Font.Bold = $true
        $title.TextFrame.TextRange.Font.Color.RGB = Rgb 255 255 255

        $subtitle = $slide.Shapes.AddTextbox(1, 54, 112, 420, 64)
        $subtitle.TextFrame.TextRange.Text = "The themed master supplies the banner, rule, background, and footer while the slide carries local KPI content."
        $subtitle.TextFrame.TextRange.Font.Name = "Aptos"
        $subtitle.TextFrame.TextRange.Font.Size = 18
        $subtitle.TextFrame.TextRange.Font.Color.RGB = Rgb 45 45 45

        $kpi = $slide.Shapes.AddShape(1, 54, 216, 180, 108)
        $kpi.Fill.ForeColor.RGB = Rgb 255 255 255
        $kpi.Line.ForeColor.RGB = Rgb 31 78 121
        $kpi.Line.Weight = 1.5
        $kpi.TextFrame.TextRange.Text = "North`r+14%"
        $kpi.TextFrame.TextRange.Font.Name = "Aptos"
        $kpi.TextFrame.TextRange.Font.Size = 24
        $kpi.TextFrame.TextRange.Font.Color.RGB = Rgb 31 78 121

        $callout = $slide.Shapes.AddShape(9, 432, 204, 162, 126)
        $callout.Fill.ForeColor.RGB = Rgb 112 173 71
        $callout.Line.Visible = 0
        $callout.TextFrame.TextRange.Text = "Plan on track"
        $callout.TextFrame.TextRange.Font.Name = "Aptos"
        $callout.TextFrame.TextRange.Font.Size = 18
        $callout.TextFrame.TextRange.Font.Color.RGB = Rgb 255 255 255

        $corporate.SaveAs((Join-Path $cases "pptx-corporate-theme.pptx"), 24)
    }
    finally {
        $corporate.Close()
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

    $centeredExplicitTable = $powerPoint.Presentations.Add($false)
    try {
        $slide = $centeredExplicitTable.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $tableShape = $slide.Shapes.AddTable(11, 5, 65.87, 113.10, 827.99, 250.39)
        $table = $tableShape.Table
        $columnWidths = @(97.764016, 112.957717, 104.792047, 297.267323, 215.218898)
        for ($column = 1; $column -le 5; $column++) {
            $table.Columns.Item($column).Width = $columnWidths[$column - 1]
        }

        $rowHeights = @(13.989606, 9.792756, 13.989606, 22.383465, 9.792756, 13.989606, 13.989606, 18.186535, 13.989606, 13.989606, 18.186535)
        for ($row = 1; $row -le 11; $row++) {
            $table.Rows.Item($row).Height = $rowHeights[$row - 1]
        }

        for ($row = 1; $row -le 11; $row++) {
            for ($column = 1; $column -le 5; $column++) {
                $cell = $table.Cell($row, $column)
                $textFrame = $cell.Shape.TextFrame2
                $textFrame.VerticalAnchor = 3
                $textFrame.MarginLeft = 2.953701
                $textFrame.MarginRight = 2.953701
                $textFrame.MarginTop = 1.47685
                $textFrame.MarginBottom = 1.47685

                $tag = "R{0:00}C{1:00}" -f $row, $column
                if ($column -eq 4) {
                    $cell.Shape.TextFrame.TextRange.Text = "$tag neutral planning horizon wraps"
                }
                elseif ($column -eq 5) {
                    $cell.Shape.TextFrame.TextRange.Text = "$tag medium neutral note"
                }
                else {
                    $cell.Shape.TextFrame.TextRange.Text = "$tag compact"
                }

                $cell.Shape.TextFrame.TextRange.Font.Name = "Aptos"
                $cell.Shape.TextFrame.TextRange.Font.Size = 10
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = Rgb 0 0 0
                $cell.Shape.Fill.ForeColor.RGB = Rgb 240 240 240
            }
        }

        $centeredExplicitTable.SaveAs((Join-Path $cases "pptx-ladder-10-table-center-explicit-wrapped.pptx"), 24)
    }
    finally {
        $centeredExplicitTable.Close()
    }

    New-MultilineTableFixtureFromWrapped `
        (Join-Path $cases "pptx-ladder-10-table-center-explicit-wrapped.pptx") `
        (Join-Path $cases "pptx-ladder-10-table-center-explicit-multiline.pptx")

    $middleSmallInsetTable = $powerPoint.Presentations.Add($false)
    try {
        $slide = $middleSmallInsetTable.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $tableShape = $slide.Shapes.AddTable(6, 7, 65.87, 132.29, 828.00, 338.84)
        $table = $tableShape.Table
        $columnWidths = @(32.05, 76.18, 277.41, 161.71, 117.15, 87.68, 75.82)
        for ($column = 1; $column -le 7; $column++) {
            $table.Columns.Item($column).Width = $columnWidths[$column - 1]
        }

        $rowHeights = @(28.66, 65.48, 54.76, 69.70, 58.34, 61.91)
        for ($row = 1; $row -le 6; $row++) {
            $table.Rows.Item($row).Height = $rowHeights[$row - 1]
        }

        $values = @(
            @("A", "Metric", "Planning signal width check with neutral words", "Current status detail", "Value note", "Target signal", "Owner"),
            @("1", "Alpha row", "Long neutral planning phrase with enough repeated terms to wrap across several lines in the wide cell", "Short multi run note for baseline", "Stable scenario", "One line", "x"),
            @("2", "Beta row", "Capacity planning phrase with several ordinary words for wrap testing", "Another compact multi run baseline note", "Reference bucket", "One line", "x"),
            @("3", "Gamma row", "Longer operational planning sentence with repeated neutral words to exercise table cell vertical centering and wrapping", "Multi segment baseline note with suffix text", "Long status cell with neutral words", "Short note", "x"),
            @("4", "Delta row", "Forecast planning phrase with enough length for controlled wrapping", "Compact note", "Stable cell", "One line", "x"),
            @("5", "Epsilon row", "Inventory planning phrase with neutral words for wrapping", "Two line centered note with suffix", "Final value", "One line", "x")
        )

        for ($row = 1; $row -le 6; $row++) {
            for ($column = 1; $column -le 7; $column++) {
                $cell = $table.Cell($row, $column)
                $textFrame = $cell.Shape.TextFrame2
                $textFrame.VerticalAnchor = 3
                $textFrame.MarginLeft = 1.255
                $textFrame.MarginRight = 1.255
                $textFrame.MarginTop = 0.627
                $textFrame.MarginBottom = 0.627
                $textFrame.TextRange.Text = $values[$row - 1][$column - 1]
                $textFrame.TextRange.Font.Name = "Aptos"
                $textFrame.TextRange.Font.Size = 11
                $textFrame.TextRange.Font.Fill.ForeColor.RGB = Rgb 30 30 30
                if ($row -eq 1) {
                    $textFrame.TextRange.Font.Bold = $true
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 230 235 242
                }
                elseif ($column -eq 3 -or $column -eq 4) {
                    $legacyRange = $cell.Shape.TextFrame.TextRange
                    $length = $legacyRange.Text.Length
                    if ($length -gt 12) {
                        $legacyRange.Characters(1, [Math]::Min(8, $length)).Font.Bold = $true
                    }

                    if ($length -gt 24) {
                        $start = [Math]::Min(18, $length)
                        $legacyRange.Characters($start, [Math]::Min(6, $length - $start + 1)).Font.Italic = $true
                    }
                }

                if ($row -eq 1) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 230 235 242
                }
                elseif ($row % 2 -eq 0) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 248 250 252
                }
                else {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 255 255 255
                }
            }
        }

        $middleSmallInsetTable.SaveAs((Join-Path $cases "pptx-ladder-10-table-middle-small-insets.pptx"), 24)
    }
    finally {
        $middleSmallInsetTable.Close()
    }

    $fragmentedTable = $powerPoint.Presentations.Add($false)
    try {
        $slide = $fragmentedTable.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $tableShape = $slide.Shapes.AddTable(6, 7, 65.87, 132.29, 828.00, 338.84)
        $table = $tableShape.Table
        $columnWidths = @(32.05, 76.18, 277.41, 161.71, 117.15, 87.68, 75.82)
        for ($column = 1; $column -le 7; $column++) {
            $table.Columns.Item($column).Width = $columnWidths[$column - 1]
        }

        $rowHeights = @(28.66, 65.48, 54.76, 69.70, 58.34, 61.91)
        for ($row = 1; $row -le 6; $row++) {
            $table.Rows.Item($row).Height = $rowHeights[$row - 1]
        }

        $values = @(
            @("A", "Metric planning", "Planning signal width check neutral operations model", "Current status detail neutral terms", "Value note", "Target signal", "Owner"),
            @("1", "Alpha planning row", "Long neutral planning phrase with repeated operational terms to wrap across several lines in the wide cell", "Short multi segment note for baseline comparison", "Stable scenario with terms", "100.00", "x"),
            @("2", "Beta planning row", "Capacity planning phrase with several ordinary words for wrap and splitting tests", "Another compact multi segment baseline note", "Reference bucket detail", "200.00", "x"),
            @("3", "Gamma planning row", "Longer operational planning sentence with repeated neutral words to exercise table cell vertical centering wrapping and run boundaries", "Multi segment baseline note with suffix text and neutral words", "Long status cell with neutral words", "300.00", "x"),
            @("4", "Delta planning row", "Forecast planning phrase with enough length for controlled wrapping and repeated terms", "Compact note with terms", "Stable cell value", "400.00", "x"),
            @("5", "Epsilon planning row", "Inventory planning phrase with neutral words for wrapping and splitting", "Two line centered note with suffix and terms", "Final value note", "500.00", "x")
        )

        for ($row = 1; $row -le 6; $row++) {
            for ($column = 1; $column -le 7; $column++) {
                $cell = $table.Cell($row, $column)
                $textFrame = $cell.Shape.TextFrame2
                $textFrame.VerticalAnchor = 3
                $textFrame.MarginLeft = 1.255
                $textFrame.MarginRight = 1.255
                $textFrame.MarginTop = 0.627
                $textFrame.MarginBottom = 0.627
                $textFrame.TextRange.Text = $values[$row - 1][$column - 1]
                $textFrame.TextRange.Font.Name = "Cambria Math"
                $textFrame.TextRange.Font.Size = 11
                $textFrame.TextRange.Font.Fill.ForeColor.RGB = Rgb 30 30 30

                $legacyRange = $cell.Shape.TextFrame.TextRange
                $legacyRange.ParagraphFormat.Alignment = if ($column -ge 6) { 3 } else { 1 }
                $length = $legacyRange.Text.Length
                if ($row -eq 1) {
                    $legacyRange.Font.Bold = $true
                }

                if ($length -gt 10 -and $column -ge 2 -and $column -le 5) {
                    $legacyRange.Characters(1, [Math]::Min(8, $length)).Font.Bold = $true
                }

                if ($length -gt 24 -and $column -ge 2 -and $column -le 5) {
                    $start = [Math]::Min(18, $length)
                    $legacyRange.Characters($start, [Math]::Min(6, $length - $start + 1)).Font.Italic = $true
                }

                if ($length -gt 38 -and $column -ge 2 -and $column -le 5) {
                    $start = [Math]::Min(32, $length)
                    $legacyRange.Characters($start, [Math]::Min(5, $length - $start + 1)).Font.Bold = $false
                }

                if ($row -eq 1) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 230 235 242
                }
                elseif ($row % 2 -eq 0) {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 248 250 252
                }
                else {
                    $cell.Shape.Fill.ForeColor.RGB = Rgb 255 255 255
                }
            }
        }

        $fragmentedTable.SaveAs((Join-Path $cases "pptx-ladder-10-table-font-fragmentation.pptx"), 24)
    }
    finally {
        $fragmentedTable.Close()
    }
}
finally {
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
    }
    Release-ComObject $powerPoint
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
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

    $basic = $word.Documents.Add()
    try {
        $basic.PageSetup.PageWidth = 612
        $basic.PageSetup.PageHeight = 792
        $basic.PageSetup.TopMargin = 72
        $basic.PageSetup.BottomMargin = 72
        $basic.PageSetup.LeftMargin = 72
        $basic.PageSetup.RightMargin = 72
        $basic.Content.Text = "Quarterly memo`r`nRevenue grew across regions with stable margins.`r`nCentered note"

        $title = $basic.Paragraphs.Item(1).Range
        $title.Font.Name = "Arial"
        $title.Font.Size = 24
        $title.Font.Bold = $true
        $title.Font.Color = Rgb 47 128 237
        $title.ParagraphFormat.SpaceAfter = 12

        $body = $basic.Paragraphs.Item(2).Range
        $body.Font.Name = "Arial"
        $body.Font.Size = 12
        $body.Font.Color = Rgb 30 30 30
        $body.ParagraphFormat.SpaceAfter = 12

        $note = $basic.Paragraphs.Item(3).Range
        $note.Font.Name = "Arial"
        $note.Font.Size = 14
        $note.Font.Italic = $true
        $note.Font.Color = Rgb 39 174 96
        $note.ParagraphFormat.Alignment = 1

        $basic.SaveAs2((Join-Path $cases "docx-basic-paragraphs.docx"), 16)
    }
    finally {
        $basic.Close($false)
    }

    $characterSpacing = $word.Documents.Add()
    try {
        $characterSpacing.PageSetup.PageWidth = 612
        $characterSpacing.PageSetup.PageHeight = 792
        $characterSpacing.PageSetup.TopMargin = 72
        $characterSpacing.PageSetup.BottomMargin = 72
        $characterSpacing.PageSetup.LeftMargin = 72
        $characterSpacing.PageSetup.RightMargin = 72
        $characterSpacing.Content.Text = "DOCX character spacing`r`nNormal sample text`r`nWide sample text`r`nTight sample text`r`nMixed boundary text`r`n`r`n"

        $title = $characterSpacing.Paragraphs.Item(1).Range
        $title.Font.Name = "Arial"
        $title.Font.Size = 22
        $title.Font.Bold = $true
        $title.Font.Color = Rgb 47 128 237
        $title.ParagraphFormat.SpaceAfter = 12

        $normal = $characterSpacing.Paragraphs.Item(2).Range
        $normal.Font.Name = "Arial"
        $normal.Font.Size = 14
        $normal.Font.Spacing = 0

        $wide = $characterSpacing.Paragraphs.Item(3).Range
        $wide.Font.Name = "Arial"
        $wide.Font.Size = 14
        $wide.Font.Spacing = 2

        $tight = $characterSpacing.Paragraphs.Item(4).Range
        $tight.Font.Name = "Arial"
        $tight.Font.Size = 14
        $tight.Font.Spacing = -1

        $mixed = $characterSpacing.Paragraphs.Item(5).Range
        $mixed.Font.Name = "Arial"
        $mixed.Font.Size = 14
        $mixed.Font.Spacing = 0
        $characterSpacing.Range($mixed.Start, $mixed.Start + 6).Font.Spacing = 2
        $characterSpacing.Range($mixed.Start + 6, $mixed.Start + 14).Font.Spacing = -1

        $tableRange = $characterSpacing.Paragraphs.Item(7).Range
        $tbl = $characterSpacing.Tables.Add($tableRange, 2, 2)
        $tbl.Borders.Enable = $true
        $tbl.Rows.Alignment = 0
        $tbl.Columns.Item(1).Width = 180
        $tbl.Columns.Item(2).Width = 180
        $tbl.Cell(1, 1).Range.Text = "Table normal"
        $tbl.Cell(1, 2).Range.Text = "Table wide"
        $tbl.Cell(2, 1).Range.Text = "Table tight"
        $tbl.Cell(2, 2).Range.Text = "Table mixed"
        foreach ($cell in @($tbl.Cell(1, 1), $tbl.Cell(1, 2), $tbl.Cell(2, 1), $tbl.Cell(2, 2))) {
            $cell.Range.Font.Name = "Arial"
            $cell.Range.Font.Size = 12
        }
        $tbl.Cell(1, 2).Range.Font.Spacing = 2
        $tbl.Cell(2, 1).Range.Font.Spacing = -1
        $mixedCell = $tbl.Cell(2, 2).Range
        $characterSpacing.Range($mixedCell.Start, $mixedCell.Start + 6).Font.Spacing = 2
        $characterSpacing.Range($mixedCell.Start + 6, $mixedCell.Start + 12).Font.Spacing = -1

        $header = $characterSpacing.Sections.Item(1).Headers.Item(1).Range
        $header.Text = "Header wide tracking"
        $header.Font.Name = "Arial"
        $header.Font.Size = 10
        $header.Font.Spacing = 2
        $header.ParagraphFormat.Alignment = 1

        $footer = $characterSpacing.Sections.Item(1).Footers.Item(1).Range
        $footer.Text = "Footer tight tracking"
        $footer.Font.Name = "Arial"
        $footer.Font.Size = 10
        $footer.Font.Spacing = -1
        $footer.ParagraphFormat.Alignment = 1

        $characterSpacing.SaveAs2((Join-Path $cases "docx-ladder-02-character-spacing.docx"), 16)
    }
    finally {
        $characterSpacing.Close($false)
    }

    $numbering = $word.Documents.Add()
    try {
        $numbering.PageSetup.PageWidth = 612
        $numbering.PageSetup.PageHeight = 792
        $numbering.PageSetup.TopMargin = 72
        $numbering.PageSetup.BottomMargin = 72
        $numbering.PageSetup.LeftMargin = 72
        $numbering.PageSetup.RightMargin = 72
        $numbering.Content.Text = "Numbered priorities`r`nImprove renderer fidelity`r`nExpand visual cases`r`nDocument diagnostics"

        $title = $numbering.Paragraphs.Item(1).Range
        $title.Font.Name = "Arial"
        $title.Font.Size = 22
        $title.Font.Bold = $true
        $title.Font.Color = Rgb 47 128 237
        $title.ParagraphFormat.SpaceAfter = 12

        $listRange = $numbering.Range($numbering.Paragraphs.Item(2).Range.Start, $numbering.Paragraphs.Item(4).Range.End)
        $listRange.Font.Name = "Arial"
        $listRange.Font.Size = 12
        $listRange.Font.Color = Rgb 30 30 30
        $listRange.ListFormat.ApplyNumberDefault()

        $numbering.SaveAs2((Join-Path $cases "docx-numbering.docx"), 16)
    }
    finally {
        $numbering.Close($false)
    }

    Add-Type -AssemblyName System.Drawing
    $docImagePath = Join-Path $env:TEMP ("ooxpdf-docx-image-" + [guid]::NewGuid() + ".png")
    $docBitmap = [System.Drawing.Bitmap]::new(180, 90)
    try {
        $docGraphics = [System.Drawing.Graphics]::FromImage($docBitmap)
        try {
            $docGraphics.Clear([System.Drawing.Color]::FromArgb(47, 128, 237))
            $docBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(39, 174, 96))
            try {
                $docGraphics.FillEllipse($docBrush, 55, 20, 70, 50)
            }
            finally {
                $docBrush.Dispose()
            }
        }
        finally {
            $docGraphics.Dispose()
        }
        $docBitmap.Save($docImagePath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $docBitmap.Dispose()
    }

    $images = $word.Documents.Add()
    try {
        $images.PageSetup.PageWidth = 612
        $images.PageSetup.PageHeight = 792
        $images.PageSetup.TopMargin = 72
        $images.PageSetup.BottomMargin = 72
        $images.PageSetup.LeftMargin = 72
        $images.PageSetup.RightMargin = 72
        $images.Content.Text = "Inline image`r`n"

        $heading = $images.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $insertRange = $images.Paragraphs.Item(2).Range
        $insertRange.InlineShapes.AddPicture($docImagePath, $false, $true) | Out-Null
        $images.SaveAs2((Join-Path $cases "docx-images.docx"), 16)
    }
    finally {
        $images.Close($false)
        Remove-Item -LiteralPath $docImagePath -Force -ErrorAction SilentlyContinue
    }

    $tableDoc = $word.Documents.Add()
    try {
        $tableDoc.PageSetup.PageWidth = 612
        $tableDoc.PageSetup.PageHeight = 792
        $tableDoc.PageSetup.TopMargin = 72
        $tableDoc.PageSetup.BottomMargin = 72
        $tableDoc.PageSetup.LeftMargin = 72
        $tableDoc.PageSetup.RightMargin = 72
        $tableDoc.Content.Text = "Regional results`r`n"

        $heading = $tableDoc.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $tableDoc.Paragraphs.Item(2).Range
        $tbl = $tableDoc.Tables.Add($range, 3, 3)
        $values = @(
            @("Region", "Q1", "Q2"),
            @("North", "42", "48"),
            @("South", "37", "41")
        )
        for ($row = 1; $row -le 3; $row++) {
            for ($column = 1; $column -le 3; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Text = $values[$row - 1][$column - 1]
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 11
                if ($row -eq 1) {
                    $cell.Shading.BackgroundPatternColor = Rgb 47 128 237
                    $cell.Range.Font.Color = Rgb 255 255 255
                    $cell.Range.Font.Bold = $true
                }
                elseif (($row + $column) % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = Rgb 232 244 253
                }
            }
        }
        $tbl.Borders.Enable = $true

        $tableDoc.SaveAs2((Join-Path $cases "docx-tables.docx"), 16)
    }
    finally {
        $tableDoc.Close($false)
    }

    $tableTextState = $word.Documents.Add()
    try {
        $tableTextState.PageSetup.PageWidth = 612
        $tableTextState.PageSetup.PageHeight = 792
        $tableTextState.PageSetup.TopMargin = 72
        $tableTextState.PageSetup.BottomMargin = 72
        $tableTextState.PageSetup.LeftMargin = 72
        $tableTextState.PageSetup.RightMargin = 72
        $tableTextState.Content.Text = "DOCX table text state`r`n"

        $heading = $tableTextState.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $tableTextState.Paragraphs.Item(2).Range
        $tbl = $tableTextState.Tables.Add($range, 4, 4)
        $tbl.Borders.Enable = $true
        $tbl.AllowAutoFit = $false
        $tbl.Rows.Alignment = 0
        $tbl.Columns.Item(1).Width = 120
        $tbl.Columns.Item(2).Width = 95
        $tbl.Columns.Item(3).Width = 95
        $tbl.Columns.Item(4).Width = 160

        $values = @(
            @("Segment", "Value A", "Value B", "Status"),
            @("North", "42", "48", "Stable demand"),
            @("South", "37", "41", "Tight capacity"),
            @("East", "128", "133", "Longer planning note")
        )
        for ($row = 1; $row -le 4; $row++) {
            for ($column = 1; $column -le 4; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Text = $values[$row - 1][$column - 1]
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 11
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                if ($row -eq 1) {
                    $cell.Shading.BackgroundPatternColor = Rgb 47 128 237
                    $cell.Range.Font.Color = Rgb 255 255 255
                    $cell.Range.Font.Bold = $true
                }
                elseif (($row + $column) % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = Rgb 232 244 253
                }
            }
        }

        $tableTextState.SaveAs2((Join-Path $cases "docx-ladder-03-table-text-state.docx"), 16)
    }
    finally {
        $tableTextState.Close($false)
    }

    $textStateContext = $word.Documents.Add()
    try {
        $textStateContext.PageSetup.PageWidth = 612
        $textStateContext.PageSetup.PageHeight = 792
        $textStateContext.PageSetup.TopMargin = 72
        $textStateContext.PageSetup.BottomMargin = 72
        $textStateContext.PageSetup.LeftMargin = 72
        $textStateContext.PageSetup.RightMargin = 72
        $textStateContext.Content.Text = "DOCX text-state context`r`n42`r`nQ1`r`nNorth`r`n"

        $heading = $textStateContext.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        for ($paragraph = 2; $paragraph -le 4; $paragraph++) {
            $body = $textStateContext.Paragraphs.Item($paragraph).Range
            $body.Font.Name = "Arial"
            $body.Font.Size = 11
            $body.Font.Color = Rgb 30 30 30
            $body.ParagraphFormat.SpaceBefore = 0
            $body.ParagraphFormat.SpaceAfter = 0
        }

        $range = $textStateContext.Paragraphs.Item(5).Range
        $tbl = $textStateContext.Tables.Add($range, 3, 4)
        $tbl.Borders.Enable = $true
        $tbl.AllowAutoFit = $false
        $tbl.Rows.Alignment = 0
        $tbl.Columns.Item(1).Width = 95
        $tbl.Columns.Item(2).Width = 95
        $tbl.Columns.Item(3).Width = 95
        $tbl.Columns.Item(4).Width = 185

        $values = @(
            @("42", "Q1", "AB", "North"),
            @("42", "Q1", "AB", "North"),
            @("128", "R2", "X9", "East")
        )
        for ($row = 1; $row -le 3; $row++) {
            for ($column = 1; $column -le 4; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Text = $values[$row - 1][$column - 1]
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 11
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                if ($row -eq 1) {
                    $cell.Shading.BackgroundPatternColor = Rgb 47 128 237
                    $cell.Range.Font.Color = Rgb 255 255 255
                    $cell.Range.Font.Bold = $true
                }
                elseif (($row + $column) % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = Rgb 232 244 253
                }
            }
        }

        $textStateContext.SaveAs2((Join-Path $cases "docx-ladder-03-text-state-context.docx"), 16)
    }
    finally {
        $textStateContext.Close($false)
    }

    $textStateSizeMatrix = $word.Documents.Add()
    try {
        $textStateSizeMatrix.PageSetup.PageWidth = 612
        $textStateSizeMatrix.PageSetup.PageHeight = 792
        $textStateSizeMatrix.PageSetup.TopMargin = 72
        $textStateSizeMatrix.PageSetup.BottomMargin = 72
        $textStateSizeMatrix.PageSetup.LeftMargin = 72
        $textStateSizeMatrix.PageSetup.RightMargin = 72
        $textStateSizeMatrix.Content.Text = "DOCX text-state size matrix`r`n"

        $heading = $textStateSizeMatrix.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $textStateSizeMatrix.Paragraphs.Item(2).Range
        $tbl = $textStateSizeMatrix.Tables.Add($range, 6, 5)
        $tbl.Borders.Enable = $true
        $tbl.AllowAutoFit = $false
        $tbl.Rows.Alignment = 0
        $tbl.Columns.Item(1).Width = 70
        $tbl.Columns.Item(2).Width = 80
        $tbl.Columns.Item(3).Width = 80
        $tbl.Columns.Item(4).Width = 80
        $tbl.Columns.Item(5).Width = 160

        $sizes = @(8, 9, 10, 11, 12, 14)
        $values = @("42", "Q1", "AB", "North")
        for ($row = 1; $row -le $sizes.Count; $row++) {
            $fontSize = $sizes[$row - 1]
            $tbl.Cell($row, 1).Range.Text = "$fontSize pt"
            for ($column = 2; $column -le 5; $column++) {
                $tbl.Cell($row, $column).Range.Text = $values[$column - 2]
            }

            for ($column = 1; $column -le 5; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = $fontSize
                $cell.Range.Font.Color = Rgb 30 30 30
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                $cell.TopPadding = 1.5
                $cell.BottomPadding = 1.5
                $cell.LeftPadding = 5.4
                $cell.RightPadding = 5.4
                if ($row % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = Rgb 232 244 253
                }
            }
        }

        $textStateSizeMatrix.SaveAs2((Join-Path $cases "docx-ladder-03-text-state-size-matrix.docx"), 16)
    }
    finally {
        $textStateSizeMatrix.Close($false)
    }

    $textStateFontMatrix = $word.Documents.Add()
    try {
        $textStateFontMatrix.PageSetup.PageWidth = 612
        $textStateFontMatrix.PageSetup.PageHeight = 792
        $textStateFontMatrix.PageSetup.TopMargin = 72
        $textStateFontMatrix.PageSetup.BottomMargin = 72
        $textStateFontMatrix.PageSetup.LeftMargin = 72
        $textStateFontMatrix.PageSetup.RightMargin = 72
        $textStateFontMatrix.Content.Text = "DOCX text-state font matrix`r`n"

        $heading = $textStateFontMatrix.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $textStateFontMatrix.Paragraphs.Item(2).Range
        $tbl = $textStateFontMatrix.Tables.Add($range, 6, 5)
        $tbl.Borders.Enable = $true
        $tbl.AllowAutoFit = $false
        $tbl.Rows.Alignment = 0
        $tbl.Columns.Item(1).Width = 125
        $tbl.Columns.Item(2).Width = 75
        $tbl.Columns.Item(3).Width = 75
        $tbl.Columns.Item(4).Width = 75
        $tbl.Columns.Item(5).Width = 145

        $fonts = @("Arial", "Calibri", "Times New Roman", "Courier New", "Georgia", "Verdana")
        $values = @("42", "Q1", "AB", "North")
        for ($row = 1; $row -le $fonts.Count; $row++) {
            $fontName = $fonts[$row - 1]
            $tbl.Cell($row, 1).Range.Text = $fontName
            for ($column = 2; $column -le 5; $column++) {
                $tbl.Cell($row, $column).Range.Text = $values[$column - 2]
            }

            for ($column = 1; $column -le 5; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Font.Name = $fontName
                $cell.Range.Font.Size = 11
                $cell.Range.Font.Color = Rgb 30 30 30
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                $cell.TopPadding = 1.5
                $cell.BottomPadding = 1.5
                $cell.LeftPadding = 5.4
                $cell.RightPadding = 5.4
                if ($row % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = Rgb 232 244 253
                }
            }
        }

        $textStateFontMatrix.SaveAs2((Join-Path $cases "docx-ladder-03-text-state-font-matrix.docx"), 16)
    }
    finally {
        $textStateFontMatrix.Close($false)
    }

    $rowHeights = $word.Documents.Add()
    try {
        $rowHeights.PageSetup.PageWidth = 612
        $rowHeights.PageSetup.PageHeight = 792
        $rowHeights.PageSetup.TopMargin = 72
        $rowHeights.PageSetup.BottomMargin = 72
        $rowHeights.PageSetup.LeftMargin = 72
        $rowHeights.PageSetup.RightMargin = 72
        $rowHeights.Content.Text = "DOCX table row heights`r`n"

        $heading = $rowHeights.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 22
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $rowHeights.Paragraphs.Item(2).Range
        $tbl = $rowHeights.Tables.Add($range, 4, 3)
        $tbl.Borders.Enable = $true
        $tbl.Rows.Alignment = 0
        $tbl.AllowAutoFit = $false
        $tbl.Columns.Item(1).Width = 110
        $tbl.Columns.Item(2).Width = 110
        $tbl.Columns.Item(3).Width = 250

        $values = @(
            @("Auto", "short", "Natural row height"),
            @("At least", "18 pt", "Wrapped phrase with enough neutral words to require multiple table lines"),
            @("Exact", "18 pt", "Wrapped phrase with enough neutral words to expose clipping or overflow handling"),
            @("Auto empty", "", "")
        )

        for ($row = 1; $row -le 4; $row++) {
            for ($column = 1; $column -le 3; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Text = $values[$row - 1][$column - 1]
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 11
                $cell.TopPadding = 0
                $cell.BottomPadding = 0
                $cell.LeftPadding = 5.4
                $cell.RightPadding = 5.4
            }
        }

        $tbl.Rows.Item(1).HeightRule = 0
        $tbl.Rows.Item(2).HeightRule = 1
        $tbl.Rows.Item(2).Height = 18
        $tbl.Rows.Item(3).HeightRule = 2
        $tbl.Rows.Item(3).Height = 18
        $tbl.Rows.Item(4).HeightRule = 0

        $rowHeights.SaveAs2((Join-Path $cases "docx-ladder-03-table-row-heights.docx"), 16)
    }
    finally {
        $rowHeights.Close($false)
    }

    $tableParagraphAdjacency = $word.Documents.Add()
    try {
        $tableParagraphAdjacency.PageSetup.PageWidth = 612
        $tableParagraphAdjacency.PageSetup.PageHeight = 792
        $tableParagraphAdjacency.PageSetup.TopMargin = 72
        $tableParagraphAdjacency.PageSetup.BottomMargin = 72
        $tableParagraphAdjacency.PageSetup.LeftMargin = 72
        $tableParagraphAdjacency.PageSetup.RightMargin = 72
        $tableParagraphAdjacency.Content.Text = "DOCX table paragraph adjacency`r`nAfter table zero before`r`nAfter table twenty-four before"
        $tableParagraphAdjacency.Content.Font.Name = "Arial"
        $tableParagraphAdjacency.Content.Font.Size = 12
        $tableParagraphAdjacency.Content.ParagraphFormat.SpaceBefore = 0
        $tableParagraphAdjacency.Content.ParagraphFormat.SpaceAfter = 0

        $heading = $tableParagraphAdjacency.Paragraphs.Item(1).Range
        $heading.Font.Name = "Arial"
        $heading.Font.Size = 20
        $heading.Font.Bold = $true
        $heading.Font.Color = Rgb 47 128 237
        $heading.ParagraphFormat.SpaceAfter = 12

        $range = $tableParagraphAdjacency.Paragraphs.Item(2).Range
        $range.Collapse(1)
        $tbl = $tableParagraphAdjacency.Tables.Add($range, 2, 2)
        $tbl.Borders.Enable = $true
        $tbl.AllowAutoFit = $false
        $tbl.Columns.Item(1).Width = 120
        $tbl.Columns.Item(2).Width = 240

        $values = @(
            @("A1", "A2"),
            @("B1", "B2")
        )

        for ($row = 1; $row -le 2; $row++) {
            for ($column = 1; $column -le 2; $column++) {
                $cell = $tbl.Cell($row, $column)
                $cell.Range.Text = $values[$row - 1][$column - 1]
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 11
                $cell.TopPadding = 0
                $cell.BottomPadding = 0
                $cell.LeftPadding = 5.4
                $cell.RightPadding = 5.4
            }
        }

        $zeroBefore = $tableParagraphAdjacency.Content
        if ($zeroBefore.Find.Execute("After table zero before")) {
            $zeroBefore.Font.Name = "Arial"
            $zeroBefore.Font.Size = 12
            $zeroBefore.ParagraphFormat.SpaceBefore = 0
            $zeroBefore.ParagraphFormat.SpaceAfter = 0
        }

        $twentyFourBefore = $tableParagraphAdjacency.Content
        if ($twentyFourBefore.Find.Execute("After table twenty-four before")) {
            $twentyFourBefore.Font.Name = "Arial"
            $twentyFourBefore.Font.Size = 12
            $twentyFourBefore.ParagraphFormat.SpaceBefore = 24
            $twentyFourBefore.ParagraphFormat.SpaceAfter = 0
        }

        $tableParagraphAdjacency.SaveAs2((Join-Path $cases "docx-ladder-03-table-paragraph-adjacency.docx"), 16)
    }
    finally {
        $tableParagraphAdjacency.Close($false)
    }

    $headerFooter = $word.Documents.Add()
    try {
        $headerFooter.PageSetup.PageWidth = 612
        $headerFooter.PageSetup.PageHeight = 792
        $headerFooter.PageSetup.TopMargin = 72
        $headerFooter.PageSetup.BottomMargin = 72
        $headerFooter.PageSetup.LeftMargin = 72
        $headerFooter.PageSetup.RightMargin = 72
        $headerFooter.Content.Text = "Header and footer sample`r`nThis document checks static header text and footer page numbering."

        $header = $headerFooter.Sections.Item(1).Headers.Item(1).Range
        $header.Text = "Confidential memo"
        $header.Font.Name = "Arial"
        $header.Font.Size = 10
        $header.Font.Color = Rgb 90 90 90
        $header.ParagraphFormat.Alignment = 1

        $footer = $headerFooter.Sections.Item(1).Footers.Item(1).Range
        $footer.Text = "Page "
        $footer.Font.Name = "Arial"
        $footer.Font.Size = 10
        $footer.Font.Color = Rgb 90 90 90
        $footer.Collapse(0)
        $footer.Fields.Add($footer, 33) | Out-Null
        $headerFooter.Sections.Item(1).Footers.Item(1).Range.ParagraphFormat.Alignment = 1

        $body = $headerFooter.Paragraphs.Item(1).Range
        $body.Font.Name = "Arial"
        $body.Font.Size = 22
        $body.Font.Bold = $true
        $body.Font.Color = Rgb 47 128 237

        $headerFooter.SaveAs2((Join-Path $cases "docx-headers-footers.docx"), 16)
    }
    finally {
        $headerFooter.Close($false)
    }
}
finally {
    if ($word -ne $null) {
        $word.Quit()
    }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-ChildItem -LiteralPath $cases -Include "*.pptx", "*.docx" -Recurse | Select-Object FullName, Length
