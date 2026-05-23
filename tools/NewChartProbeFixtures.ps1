$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null

function Rgb($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

$output = Join-Path $cases "pptx-ladder-11-secondary-axis-overlay-probe.pptx"
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add($true)
    $slide = $presentation.Slides.Add(1, 12)
    $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

    $chartShape = $slide.Shapes.AddChart(52, 144, 90, 432, 288)
    $chart = $chartShape.Chart
    $chart.HasTitle = $false
    $chart.HasLegend = $false
    $chart.ChartData.Activate()

    $workbook = $chart.ChartData.Workbook
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Cells.Clear()
    $worksheet.Cells.Item(1, 1).Value = "Category"
    $worksheet.Cells.Item(1, 2).Value = "Inventory"
    $worksheet.Cells.Item(1, 3).Value = "Reduction"
    $worksheet.Cells.Item(1, 4).Value = "Lost sales"
    $worksheet.Cells.Item(1, 5).Value = "Efficiency"
    $worksheet.Cells.Item(2, 1).Value = "Inventory"
    $worksheet.Cells.Item(2, 2).Value = 130.0
    $worksheet.Cells.Item(2, 3).Value = 21.0
    $worksheet.Cells.Item(2, 4).Value = 0.0
    $worksheet.Cells.Item(2, 5).Value = 0.0
    $worksheet.Cells.Item(3, 1).Value = "Gain"
    $worksheet.Cells.Item(3, 2).Value = 0.0
    $worksheet.Cells.Item(3, 3).Value = 0.0
    $worksheet.Cells.Item(3, 4).Value = 2.0
    $worksheet.Cells.Item(3, 5).Value = 4.0

    $chart.SetSourceData("=Sheet1!`$A`$1:`$E`$3")
    $chart.ChartType = 52

    $series = $chart.SeriesCollection()
    $series.Item(1).Format.Fill.ForeColor.RGB = Rgb 191 191 191
    $series.Item(2).Format.Fill.ForeColor.RGB = Rgb 192 0 0
    $series.Item(3).AxisGroup = 2
    $series.Item(3).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    $series.Item(4).AxisGroup = 2
    $series.Item(4).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    $series.Item(4).Format.Fill.Patterned(9)
    $series.Item(4).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    $series.Item(4).Format.Fill.BackColor.RGB = Rgb 191 191 191

    $primaryValueAxis = $chart.Axes(2, 1)
    $primaryValueAxis.MaximumScale = 200
    $primaryValueAxis.MajorUnit = 20
    $primaryValueAxis.Format.Line.Visible = 0
    $primaryValueAxis.TickLabelPosition = -4134
    $primaryValueAxis.TickLabels.Font.Size = 9
    $primaryValueAxis.TickLabels.Font.Color = Rgb 102 102 102

    $secondaryValueAxis = $chart.Axes(2, 2)
    $secondaryValueAxis.MaximumScale = 20
    $secondaryValueAxis.MajorUnit = 2
    $secondaryValueAxis.Format.Line.Visible = 0
    $secondaryValueAxis.TickLabelPosition = -4134
    $secondaryValueAxis.TickLabels.Font.Size = 9
    $secondaryValueAxis.TickLabels.Font.Color = Rgb 102 102 102

    $categoryAxis = $chart.Axes(1, 1)
    $categoryAxis.Format.Line.ForeColor.RGB = Rgb 217 217 217
    $categoryAxis.Format.Line.Weight = 0.75
    $categoryAxis.TickLabels.Font.Size = 9
    $categoryAxis.TickLabels.Font.Color = Rgb 102 102 102

    $line = $slide.Shapes.AddConnector(1, 608, 258, 608, 330)
    $line.Line.ForeColor.RGB = Rgb 47 133 106
    $line.Line.Weight = 1.5
    $line.Line.EndArrowheadStyle = 4

    if (Test-Path -LiteralPath $output) {
        Remove-Item -LiteralPath $output -Force
    }

    $presentation.SaveAs($output, 24)
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($powerPoint -ne $null) { $powerPoint.Quit() }
}

Get-Item -LiteralPath $output
