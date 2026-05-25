$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null

function Rgb($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Release-ComObject($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function Close-ChartWorkbook($workbook) {
    $excel = $null
    try {
        if ($null -ne $workbook) {
            $excel = $workbook.Application
            $workbook.Close($false)
        }

        if ($null -ne $excel) {
            try {
                if ($excel.Workbooks.Count -eq 0) {
                    $excel.Quit()
                }
            }
            catch {
                # Some Office builds tear down the embedded workbook host during
                # Close(). The process-lifetime cleanup below is still useful.
            }
        }
    }
    finally {
        Release-ComObject $excel
        Release-ComObject $workbook
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function New-DoughnutLegendProbe($PowerPoint, $Cases, $FileName, [int]$LegendPosition, [bool]$HasLegend = $true, [bool]$IncludeInLayout = $true) {
    $output = Join-Path $Cases $FileName
    $presentation = $null
    $workbook = $null
    $worksheet = $null

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(-4120, 144, 36, 576, 432)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $HasLegend)
        if ($HasLegend) {
            [void]($chart.Legend.Position = $LegendPosition)
            [void]($chart.Legend.IncludeInLayout = $IncludeInLayout)
        }

        $chart.ChartData.Activate()
        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Share"
        $worksheet.Cells.Item(2, 1).Value = "North"
        $worksheet.Cells.Item(2, 2).Value = 65.0
        $worksheet.Cells.Item(3, 1).Value = "South"
        $worksheet.Cells.Item(3, 2).Value = 20.0
        $worksheet.Cells.Item(4, 1).Value = "West"
        $worksheet.Cells.Item(4, 2).Value = 15.0

        $chart.SetSourceData("=Sheet1!`$A`$1:`$B`$4", 2)
        [void]($chart.ChartType = -4120)
        [void]($chart.ChartGroups(1).DoughnutHoleSize = 50)

        $series = $chart.SeriesCollection().Item(1)
        try {
            $series.Points(1).Format.Fill.ForeColor.RGB = Rgb 68 114 196
            $series.Points(2).Format.Fill.ForeColor.RGB = Rgb 237 125 49
            $series.Points(3).Format.Fill.ForeColor.RGB = Rgb 165 165 165
        }
        catch {
            # Some Office builds defer doughnut point materialization until save.
            # The default Office palette is acceptable for these geometry probes.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        $presentation.SaveAs($output, 24)
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        $presentation.Close()
        $presentation = $null
    }
    finally {
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { $presentation.Close() }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application

    $output = Join-Path $cases "pptx-ladder-11-secondary-axis-overlay-probe.pptx"
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
    Release-ComObject $worksheet
    $worksheet = $null
    Close-ChartWorkbook $workbook
    $workbook = $null

    if (Test-Path -LiteralPath $output) {
        Remove-Item -LiteralPath $output -Force
    }

    $presentation.SaveAs($output, 24)
    $presentation.Close()
    $presentation = $null

    $output = Join-Path $cases "pptx-ladder-11-compact-stacked-secondary-axis-probe.pptx"
    $presentation = $powerPoint.Presentations.Add($true)
    $slide = $presentation.Slides.Add(1, 12)
    $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

    $chartShape = $slide.Shapes.AddChart(52, 430, 165, 255, 170)
    $chart = $chartShape.Chart
    $chart.HasTitle = $false
    $chart.HasLegend = $false
    $chart.ChartData.Activate()

    $workbook = $chart.ChartData.Workbook
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Cells.Clear()
    $worksheet.Cells.Item(1, 1).Value = "Category"
    $worksheet.Cells.Item(1, 2).Value = "Stock base"
    $worksheet.Cells.Item(1, 3).Value = "Stock top"
    $worksheet.Cells.Item(1, 4).Value = "Gain base"
    $worksheet.Cells.Item(1, 5).Value = "Gain middle"
    $worksheet.Cells.Item(1, 6).Value = "Gain top"
    $worksheet.Cells.Item(2, 1).Value = "Inventory"
    $worksheet.Cells.Item(2, 2).Value = 130.0
    $worksheet.Cells.Item(2, 3).Value = 21.0
    $worksheet.Cells.Item(2, 4).Value = 0.0
    $worksheet.Cells.Item(2, 5).Value = 0.0
    $worksheet.Cells.Item(2, 6).Value = 0.0
    $worksheet.Cells.Item(3, 1).Value = "EBITDA +"
    $worksheet.Cells.Item(3, 2).Value = 0.0
    $worksheet.Cells.Item(3, 3).Value = 0.0
    $worksheet.Cells.Item(3, 4).Value = 2.0
    $worksheet.Cells.Item(3, 5).Value = 4.0
    $worksheet.Cells.Item(3, 6).Value = 1.05

    $chart.SetSourceData("=Sheet1!`$A`$1:`$F`$3")
    $chart.ChartType = 52

    $series = $chart.SeriesCollection()
    $series.Item(1).Format.Fill.ForeColor.RGB = Rgb 191 191 191
    $series.Item(2).Format.Fill.ForeColor.RGB = Rgb 192 0 0
    for ($i = 3; $i -le 5; $i++) {
        $series.Item($i).AxisGroup = 2
        $series.Item($i).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    }
    $series.Item(4).Format.Fill.Patterned(9)
    $series.Item(4).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    $series.Item(4).Format.Fill.BackColor.RGB = Rgb 191 191 191
    $series.Item(5).Format.Fill.ForeColor.RGB = Rgb 192 0 0

    $primaryValueAxis = $chart.Axes(2, 1)
    $primaryValueAxis.MaximumScale = 200
    $primaryValueAxis.MajorUnit = 20
    $primaryValueAxis.Format.Line.Visible = 0
    $primaryValueAxis.MajorGridlines.Format.Line.Visible = 0
    $primaryValueAxis.TickLabelPosition = -4134
    $primaryValueAxis.TickLabels.Font.Size = 7
    $primaryValueAxis.TickLabels.Font.Color = Rgb 102 102 102

    $secondaryValueAxis = $chart.Axes(2, 2)
    $secondaryValueAxis.MaximumScale = 20
    $secondaryValueAxis.MajorUnit = 2
    $secondaryValueAxis.Format.Line.Visible = 0
    $secondaryValueAxis.MajorGridlines.Format.Line.Visible = 0
    $secondaryValueAxis.TickLabelPosition = -4134
    $secondaryValueAxis.TickLabels.Font.Size = 7
    $secondaryValueAxis.TickLabels.Font.Color = Rgb 102 102 102

    $categoryAxis = $chart.Axes(1, 1)
    $categoryAxis.Format.Line.ForeColor.RGB = Rgb 217 217 217
    $categoryAxis.Format.Line.Weight = 0.75
    $categoryAxis.TickLabels.Font.Size = 7
    $categoryAxis.TickLabels.Font.Color = Rgb 102 102 102

    $line = $slide.Shapes.AddConnector(1, 655, 230, 655, 315)
    $line.Line.ForeColor.RGB = Rgb 47 133 106
    $line.Line.Weight = 1.0
    $line.Line.EndArrowheadStyle = 4

    $line = $slide.Shapes.AddConnector(1, 566, 195, 626, 245)
    $line.Line.ForeColor.RGB = Rgb 191 191 191
    $line.Line.Weight = 0.75
    $line = $slide.Shapes.AddConnector(1, 566, 205, 626, 255)
    $line.Line.ForeColor.RGB = Rgb 191 191 191
    $line.Line.Weight = 0.75
    Release-ComObject $worksheet
    $worksheet = $null
    Close-ChartWorkbook $workbook
    $workbook = $null

    if (Test-Path -LiteralPath $output) {
        Remove-Item -LiteralPath $output -Force
    }

    $presentation.SaveAs($output, 24)
    $presentation.Close()
    $presentation = $null

    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-no-legend-probe.pptx" `
        -LegendPosition -4152 `
        -HasLegend $false `
        -IncludeInLayout $true
    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-left-legend-probe.pptx" `
        -LegendPosition -4131 `
        -HasLegend $true `
        -IncludeInLayout $true
    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-top-legend-probe.pptx" `
        -LegendPosition -4160 `
        -HasLegend $true `
        -IncludeInLayout $true
    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-bottom-legend-probe.pptx" `
        -LegendPosition -4107 `
        -HasLegend $true `
        -IncludeInLayout $true
    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-right-overlay-legend-probe.pptx" `
        -LegendPosition -4152 `
        -HasLegend $true `
        -IncludeInLayout $false
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() }
        catch {
            # COM cleanup should not mask the generation error that preceded it.
        }
    }
    if ($powerPoint -ne $null) {
        try { $powerPoint.Quit() }
        catch {
            # PowerPoint may already be unavailable after a COM failure.
        }
    }
    Release-ComObject $presentation
    Release-ComObject $powerPoint
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-Item -LiteralPath $output
