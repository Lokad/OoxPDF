param(
    [switch] $DoughnutOnly,
    [switch] $SparseOnly,
    [switch] $DataLabelsOnly,
    [switch] $AxisTitlesOnly,
    [string] $MetadataRoot
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$cases = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases"
New-Item -ItemType Directory -Force -Path $cases | Out-Null
if ([string]::IsNullOrWhiteSpace($MetadataRoot)) {
    $MetadataRoot = Join-Path $repoRoot "artifacts/office-probe-metadata"
}
else {
    $MetadataRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($MetadataRoot)
}

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
            [void]($workbook.Close($false))
        }

        if ($null -ne $excel) {
            try {
                if ($excel.Workbooks.Count -eq 0) {
                    [void]($excel.Quit())
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

function Read-ComMember($object, [string] $name) {
    if ($null -eq $object) {
        return $null
    }

    try {
        return $object.$name
    }
    catch {
        return $null
    }
}

function Round-ComNumber($value) {
    if ($null -eq $value) {
        return $null
    }

    try {
        return [Math]::Round([double]$value, 6)
    }
    catch {
        return $null
    }
}

function New-OfficeProbeMetadataPath([string] $metadataRoot, [string] $fileName) {
    $directory = Join-Path $metadataRoot ([System.IO.Path]::GetFileNameWithoutExtension($fileName))
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    return Join-Path $directory "com-metadata.json"
}

function Write-OfficeProbeMetadata([string] $metadataRoot, [string] $fileName, $metadata) {
    if ([string]::IsNullOrWhiteSpace($metadataRoot)) {
        return
    }

    $metadataPath = New-OfficeProbeMetadataPath $metadataRoot $fileName
    $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $metadataPath -Encoding UTF8
    Write-Host "Office probe metadata: $metadataPath"
}

function New-DoughnutLegendProbe($PowerPoint, $Cases, $FileName, [int]$LegendPosition, [bool]$HasLegend = $true, [bool]$IncludeInLayout = $true, [int]$Explosion = 0) {
    $output = Join-Path $Cases $FileName
    $presentation = $null
    $workbook = $null
    $worksheet = $null
    $series = $null
    $labels = $null

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

        [void]($chart.ChartData.Activate())
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

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$B`$4", 2))
        [void]($chart.ChartType = -4120)
        [void]($chart.ChartGroups(1).DoughnutHoleSize = 50)

        $series = $chart.SeriesCollection().Item(1)
        if ($Explosion -gt 0) {
            [void]($series.Explosion = $Explosion)
        }
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

        [void]($presentation.SaveAs($output, 24))
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function New-PieDataLabelLeaderLineProbe($PowerPoint, $Cases, [string] $MetadataRoot) {
    $fileName = "pptx-ladder-11-chart-pie-data-label-leader-lines-probe.pptx"
    $output = Join-Path $Cases $fileName
    $presentation = $null
    $workbook = $null
    $worksheet = $null
    $labelMetadata = @()

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(5, 144, 60, 520, 360)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $false)
        [void]($chart.ChartData.Activate())

        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Share"
        $worksheet.Cells.Item(2, 1).Value = "Alpha"
        $worksheet.Cells.Item(2, 2).Value = 48.0
        $worksheet.Cells.Item(3, 1).Value = "Beta"
        $worksheet.Cells.Item(3, 2).Value = 22.0
        $worksheet.Cells.Item(4, 1).Value = "Gamma"
        $worksheet.Cells.Item(4, 2).Value = 17.0
        $worksheet.Cells.Item(5, 1).Value = "Delta"
        $worksheet.Cells.Item(5, 2).Value = 13.0

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$B`$5", 2))
        [void]($chart.ChartType = 5)

        $series = $chart.SeriesCollection().Item(1)
        try {
            $series.Points(1).Format.Fill.ForeColor.RGB = Rgb 68 114 196
            $series.Points(2).Format.Fill.ForeColor.RGB = Rgb 237 125 49
            $series.Points(3).Format.Fill.ForeColor.RGB = Rgb 165 165 165
            $series.Points(4).Format.Fill.ForeColor.RGB = Rgb 112 173 71
        }
        catch {
            # Point materialization can be deferred by Office until save.
        }

        [void]($series.ApplyDataLabels())
        $labels = $series.DataLabels()
        $labels.ShowCategoryName = $true
        $labels.ShowValue = $false
        $labels.ShowPercentage = $true
        $labels.ShowSeriesName = $false
        $labels.ShowLegendKey = $false
        $labels.Separator = " "
        $labels.Position = 2
        try {
            $series.HasLeaderLines = $true
        }
        catch {
            # Some Office versions infer leader lines from outside labels and
            # do not expose this property until the chart is laid out.
        }
        try {
            $series.LeaderLines.Format.Line.ForeColor.RGB = Rgb 255 0 0
            $series.LeaderLines.Format.Line.Weight = 2.25
        }
        catch {
            # The generated PDF remains the source of truth for leader geometry.
        }

        $manualLabelPositions = @(
            @{ Left = 10; Top = 20 },
            @{ Left = 650; Top = 40 },
            @{ Left = 0; Top = 330 },
            @{ Left = 640; Top = 390 }
        )
        for ($i = 1; $i -le $manualLabelPositions.Count; $i++) {
            $point = $null
            $pointLabel = $null
            try {
                $point = $series.Points($i)
                $pointLabel = $point.DataLabel
                try {
                    $pointLabel.Position = 7
                }
                catch {
                    # Some Office builds keep custom placement implicit after Left/Top.
                }
                $pointLabel.Left = $manualLabelPositions[$i - 1].Left
                $pointLabel.Top = $manualLabelPositions[$i - 1].Top
                $labelMetadata += [pscustomobject]@{
                    Index = $i - 1
                    RequestedLeft = $manualLabelPositions[$i - 1].Left
                    RequestedTop = $manualLabelPositions[$i - 1].Top
                    AppliedPosition = Read-ComMember $pointLabel "Position"
                    ObservedLeft = Round-ComNumber (Read-ComMember $pointLabel "Left")
                    ObservedTop = Round-ComNumber (Read-ComMember $pointLabel "Top")
                    ObservedWidth = Round-ComNumber (Read-ComMember $pointLabel "Width")
                    ObservedHeight = Round-ComNumber (Read-ComMember $pointLabel "Height")
                    Text = Read-ComMember $pointLabel "Text"
                }
            }
            catch {
                # If Office refuses manual label coordinates, keep the automatic
                # outside-end placement and let the PDF probe expose that fact.
                $labelMetadata += [pscustomobject]@{
                    Index = $i - 1
                    RequestedLeft = $manualLabelPositions[$i - 1].Left
                    RequestedTop = $manualLabelPositions[$i - 1].Top
                    Error = $_.Exception.Message
                }
            }
            finally {
                if ($pointLabel -ne $null) { Release-ComObject $pointLabel }
                if ($point -ne $null) { Release-ComObject $point }
            }
        }
        try {
            $series.HasLeaderLines = $true
            $series.LeaderLines.Format.Line.ForeColor.RGB = Rgb 255 0 0
            $series.LeaderLines.Format.Line.Weight = 2.25
        }
        catch {
            # The generated PDF remains the source of truth for leader geometry.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        $metadata = [pscustomobject]@{
            Fixture = $fileName
            ChartShape = [pscustomobject]@{
                Left = Round-ComNumber (Read-ComMember $chartShape "Left")
                Top = Round-ComNumber (Read-ComMember $chartShape "Top")
                Width = Round-ComNumber (Read-ComMember $chartShape "Width")
                Height = Round-ComNumber (Read-ComMember $chartShape "Height")
            }
            HasLeaderLines = Read-ComMember $series "HasLeaderLines"
            DataLabels = $labelMetadata
        }
        Write-OfficeProbeMetadata $MetadataRoot $fileName $metadata
        Release-ComObject $labels
        $labels = $null
        Release-ComObject $series
        $series = $null
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($labels -ne $null) { Release-ComObject $labels }
        if ($series -ne $null) { Release-ComObject $series }
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function New-PieDataLabelLeaderLineOffsetProbe($PowerPoint, $Cases, [string] $MetadataRoot) {
    $fileName = "pptx-ladder-11-chart-pie-data-label-leader-lines-offset-probe.pptx"
    $output = Join-Path $Cases $fileName
    $presentation = $null
    $workbook = $null
    $worksheet = $null
    $labelMetadata = @()
    $labels = $null
    $series = $null

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(5, 84, 96, 420, 300)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $false)
        [void]($chart.ChartData.Activate())

        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Share"
        $worksheet.Cells.Item(2, 1).Value = "North"
        $worksheet.Cells.Item(2, 2).Value = 36.0
        $worksheet.Cells.Item(3, 1).Value = "South"
        $worksheet.Cells.Item(3, 2).Value = 28.0
        $worksheet.Cells.Item(4, 1).Value = "West"
        $worksheet.Cells.Item(4, 2).Value = 21.0
        $worksheet.Cells.Item(5, 1).Value = "East"
        $worksheet.Cells.Item(5, 2).Value = 15.0

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$B`$5", 2))
        [void]($chart.ChartType = 5)

        $series = $chart.SeriesCollection().Item(1)
        try {
            $series.Points(1).Format.Fill.ForeColor.RGB = Rgb 68 114 196
            $series.Points(2).Format.Fill.ForeColor.RGB = Rgb 237 125 49
            $series.Points(3).Format.Fill.ForeColor.RGB = Rgb 165 165 165
            $series.Points(4).Format.Fill.ForeColor.RGB = Rgb 112 173 71
        }
        catch {
            # Point materialization can be deferred by Office until save.
        }

        [void]($series.ApplyDataLabels())
        $labels = $series.DataLabels()
        $labels.ShowCategoryName = $true
        $labels.ShowValue = $false
        $labels.ShowPercentage = $true
        $labels.ShowSeriesName = $false
        $labels.ShowLegendKey = $false
        $labels.Separator = " "
        $labels.Position = 2
        try {
            $series.HasLeaderLines = $true
        }
        catch {
            # Some Office versions infer leader lines from outside labels and
            # do not expose this property until the chart is laid out.
        }
        try {
            $series.LeaderLines.Format.Line.ForeColor.RGB = Rgb 255 0 0
            $series.LeaderLines.Format.Line.Weight = 2.25
        }
        catch {
            # The generated PDF remains the source of truth for leader geometry.
        }

        $manualLabelPositions = @(
            @{ Left = -24; Top = 32 },
            @{ Left = 506; Top = 76 },
            @{ Left = 28; Top = 268 },
            @{ Left = 500; Top = 320 }
        )
        for ($i = 1; $i -le $manualLabelPositions.Count; $i++) {
            $point = $null
            $pointLabel = $null
            try {
                $point = $series.Points($i)
                $pointLabel = $point.DataLabel
                try {
                    $pointLabel.Position = 7
                }
                catch {
                    # Some Office builds keep custom placement implicit after Left/Top.
                }
                $pointLabel.Left = $manualLabelPositions[$i - 1].Left
                $pointLabel.Top = $manualLabelPositions[$i - 1].Top
                $labelMetadata += [pscustomobject]@{
                    Index = $i - 1
                    RequestedLeft = $manualLabelPositions[$i - 1].Left
                    RequestedTop = $manualLabelPositions[$i - 1].Top
                    AppliedPosition = Read-ComMember $pointLabel "Position"
                    ObservedLeft = Round-ComNumber (Read-ComMember $pointLabel "Left")
                    ObservedTop = Round-ComNumber (Read-ComMember $pointLabel "Top")
                    ObservedWidth = Round-ComNumber (Read-ComMember $pointLabel "Width")
                    ObservedHeight = Round-ComNumber (Read-ComMember $pointLabel "Height")
                    Text = Read-ComMember $pointLabel "Text"
                }
            }
            catch {
                # If Office refuses manual label coordinates, keep the automatic
                # outside-end placement and let the PDF probe expose that fact.
                $labelMetadata += [pscustomobject]@{
                    Index = $i - 1
                    RequestedLeft = $manualLabelPositions[$i - 1].Left
                    RequestedTop = $manualLabelPositions[$i - 1].Top
                    Error = $_.Exception.Message
                }
            }
            finally {
                if ($pointLabel -ne $null) { Release-ComObject $pointLabel }
                if ($point -ne $null) { Release-ComObject $point }
            }
        }
        try {
            $series.HasLeaderLines = $true
            $series.LeaderLines.Format.Line.ForeColor.RGB = Rgb 255 0 0
            $series.LeaderLines.Format.Line.Weight = 2.25
        }
        catch {
            # The generated PDF remains the source of truth for leader geometry.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        $metadata = [pscustomobject]@{
            Fixture = $fileName
            ChartShape = [pscustomobject]@{
                Left = Round-ComNumber (Read-ComMember $chartShape "Left")
                Top = Round-ComNumber (Read-ComMember $chartShape "Top")
                Width = Round-ComNumber (Read-ComMember $chartShape "Width")
                Height = Round-ComNumber (Read-ComMember $chartShape "Height")
            }
            HasLeaderLines = Read-ComMember $series "HasLeaderLines"
            DataLabels = $labelMetadata
        }
        Write-OfficeProbeMetadata $MetadataRoot $fileName $metadata
        Release-ComObject $labels
        $labels = $null
        Release-ComObject $series
        $series = $null
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($labels -ne $null) { Release-ComObject $labels }
        if ($series -ne $null) { Release-ComObject $series }
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function Populate-SparseChartWorksheet($Worksheet) {
    $Worksheet.Cells.Clear()
    $Worksheet.Cells.Item(1, 1).Value = "Category"
    $Worksheet.Cells.Item(1, 2).Value = "Actual"
    $Worksheet.Cells.Item(1, 3).Value = "Plan"
    $Worksheet.Cells.Item(2, 1).Value = "Jan"
    $Worksheet.Cells.Item(2, 2).Value = 10.0
    $Worksheet.Cells.Item(2, 3).Value = 8.0
    $Worksheet.Cells.Item(3, 1).Value = "Feb"
    $Worksheet.Cells.Item(3, 2).ClearContents()
    $Worksheet.Cells.Item(3, 3).Value = 12.0
    $Worksheet.Cells.Item(4, 1).Value = "Mar"
    $Worksheet.Cells.Item(4, 2).Value = "N/A"
    $Worksheet.Cells.Item(4, 3).Value = 14.0
    $Worksheet.Cells.Item(5, 1).Value = "Apr"
    $Worksheet.Cells.Item(5, 2).Value = 18.0
    $Worksheet.Cells.Item(5, 3).ClearContents()
    $Worksheet.Cells.Item(6, 1).Value = "May"
    $Worksheet.Cells.Item(6, 2).Value = 24.0
    $Worksheet.Cells.Item(6, 3).Value = 21.0
}

function Configure-SparseProbeChart($Chart, [int]$ChartType) {
    [void]($Chart.HasTitle = $false)
    [void]($Chart.HasLegend = $true)
    [void]($Chart.ChartType = $ChartType)
    try {
        [void]($Chart.DisplayBlanksAs = 1)
    }
    catch {
        # Some Office builds reject DisplayBlanksAs for a subset of chart kinds.
    }

    $valueAxis = $Chart.Axes(2, 1)
    $valueAxis.MinimumScale = 0
    $valueAxis.MaximumScale = 30
    $valueAxis.MajorUnit = 10
    $valueAxis.TickLabels.Font.Size = 8
    $valueAxis.Format.Line.Visible = 0
    $valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = Rgb 217 217 217
    $valueAxis.MajorGridlines.Format.Line.Weight = 0.75

    $categoryAxis = $Chart.Axes(1, 1)
    $categoryAxis.TickLabels.Font.Size = 8
    $categoryAxis.Format.Line.ForeColor.RGB = Rgb 217 217 217
    $categoryAxis.Format.Line.Weight = 0.75

    $series = $Chart.SeriesCollection()
    $series.Item(1).Format.Line.ForeColor.RGB = Rgb 68 114 196
    $series.Item(1).Format.Fill.ForeColor.RGB = Rgb 68 114 196
    $series.Item(2).Format.Line.ForeColor.RGB = Rgb 237 125 49
    $series.Item(2).Format.Fill.ForeColor.RGB = Rgb 237 125 49
}

function New-DefaultAxisTitleProbe($PowerPoint, $Cases) {
    $output = Join-Path $Cases "pptx-ladder-11-chart-default-axis-titles-probe.pptx"
    $presentation = $null
    $workbook = $null
    $worksheet = $null

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(51, 120, 66, 540, 366)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $false)
        [void]($chart.ChartData.Activate())

        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Primary"
        $worksheet.Cells.Item(1, 3).Value = "Secondary"
        $worksheet.Cells.Item(2, 1).Value = "North"
        $worksheet.Cells.Item(2, 2).Value = 42.0
        $worksheet.Cells.Item(2, 3).Value = 35.0
        $worksheet.Cells.Item(3, 1).Value = "South"
        $worksheet.Cells.Item(3, 2).Value = 68.0
        $worksheet.Cells.Item(3, 3).Value = 44.0
        $worksheet.Cells.Item(4, 1).Value = "West"
        $worksheet.Cells.Item(4, 2).Value = 31.0
        $worksheet.Cells.Item(4, 3).Value = 52.0
        $worksheet.Cells.Item(5, 1).Value = "East"
        $worksheet.Cells.Item(5, 2).Value = 55.0
        $worksheet.Cells.Item(5, 3).Value = 39.0

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$C`$5"))
        [void]($chart.ChartType = 51)

        $categoryAxis = $chart.Axes(1, 1)
        [void]($categoryAxis.HasTitle = $true)
        $categoryAxis.AxisTitle.Text = "Category Axis"
        $categoryAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $categoryAxis.TickLabels.Font.Size = 9
        $categoryAxis.TickLabels.Font.Color = Rgb 89 89 89

        $valueAxis = $chart.Axes(2, 1)
        [void]($valueAxis.HasTitle = $true)
        $valueAxis.AxisTitle.Text = "Value Axis"
        $valueAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $valueAxis.TickLabels.Font.Size = 9
        $valueAxis.TickLabels.Font.Color = Rgb 89 89 89

        try {
            [void]($workbook.Application.CalculateFull())
            [void]($chart.Refresh())
        }
        catch {
            # Some Office builds do not expose chart refresh for embedded hosts;
            # saved chart caches are still verified by the visual harness.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function New-BarDefaultAxisTitleProbe($PowerPoint, $Cases) {
    $output = Join-Path $Cases "pptx-ladder-11-chart-bar-default-axis-titles-probe.pptx"
    $presentation = $null
    $workbook = $null
    $worksheet = $null

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(57, 120, 66, 540, 366)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $false)
        [void]($chart.ChartData.Activate())

        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Actual"
        $worksheet.Cells.Item(2, 1).Value = "North"
        $worksheet.Cells.Item(2, 2).Value = 42.0
        $worksheet.Cells.Item(3, 1).Value = "South"
        $worksheet.Cells.Item(3, 2).Value = 68.0
        $worksheet.Cells.Item(4, 1).Value = "West"
        $worksheet.Cells.Item(4, 2).Value = 31.0
        $worksheet.Cells.Item(5, 1).Value = "East"
        $worksheet.Cells.Item(5, 2).Value = 55.0

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$B`$5"))
        [void]($chart.ChartType = 57)

        $categoryAxis = $chart.Axes(1, 1)
        [void]($categoryAxis.HasTitle = $true)
        $categoryAxis.AxisTitle.Text = "Category Axis"
        $categoryAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $categoryAxis.TickLabels.Font.Size = 9
        $categoryAxis.TickLabels.Font.Color = Rgb 89 89 89

        $valueAxis = $chart.Axes(2, 1)
        [void]($valueAxis.HasTitle = $true)
        $valueAxis.AxisTitle.Text = "Value Axis"
        $valueAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $valueAxis.TickLabels.Font.Size = 9
        $valueAxis.TickLabels.Font.Color = Rgb 89 89 89
        $valueAxis.MinimumScale = 0
        $valueAxis.MaximumScale = 80
        $valueAxis.MajorUnit = 20

        try {
            [void]($workbook.Application.CalculateFull())
            [void]($chart.Refresh())
        }
        catch {
            # Some Office builds do not expose chart refresh for embedded hosts;
            # saved chart caches are still verified by the visual harness.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function New-TopRightDefaultAxisTitleProbe($PowerPoint, $Cases) {
    $output = Join-Path $Cases "pptx-ladder-11-chart-top-right-default-axis-titles-probe.pptx"
    $presentation = $null
    $workbook = $null
    $worksheet = $null

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartShape = $slide.Shapes.AddChart(51, 120, 66, 540, 366)
        $chart = $chartShape.Chart
        [void]($chart.HasTitle = $false)
        [void]($chart.HasLegend = $false)
        [void]($chart.ChartData.Activate())

        $workbook = $chart.ChartData.Workbook
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value = "Category"
        $worksheet.Cells.Item(1, 2).Value = "Primary"
        $worksheet.Cells.Item(1, 3).Value = "Secondary"
        $worksheet.Cells.Item(2, 1).Value = "North"
        $worksheet.Cells.Item(2, 2).Value = 42.0
        $worksheet.Cells.Item(2, 3).Value = 35.0
        $worksheet.Cells.Item(3, 1).Value = "South"
        $worksheet.Cells.Item(3, 2).Value = 68.0
        $worksheet.Cells.Item(3, 3).Value = 44.0
        $worksheet.Cells.Item(4, 1).Value = "West"
        $worksheet.Cells.Item(4, 2).Value = 31.0
        $worksheet.Cells.Item(4, 3).Value = 52.0
        $worksheet.Cells.Item(5, 1).Value = "East"
        $worksheet.Cells.Item(5, 2).Value = 55.0
        $worksheet.Cells.Item(5, 3).Value = 39.0

        [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$C`$5"))
        [void]($chart.ChartType = 51)

        $categoryAxis = $chart.Axes(1, 1)
        [void]($categoryAxis.HasTitle = $true)
        $categoryAxis.AxisTitle.Text = "Category Axis"
        $categoryAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $categoryAxis.TickLabels.Font.Size = 9
        $categoryAxis.TickLabels.Font.Color = Rgb 89 89 89

        $valueAxis = $chart.Axes(2, 1)
        [void]($valueAxis.HasTitle = $true)
        $valueAxis.AxisTitle.Text = "Value Axis"
        $valueAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $valueAxis.TickLabels.Font.Size = 9
        $valueAxis.TickLabels.Font.Color = Rgb 89 89 89
        $valueAxis.MinimumScale = 0
        $valueAxis.MaximumScale = 80
        $valueAxis.MajorUnit = 20

        # xlMaximum = 2. These crossings make Office save the category axis at the top
        # and the value axis at the right without manually positioning the titles.
        $valueAxis.Crosses = 2
        $categoryAxis.Crosses = 2

        try {
            [void]($workbook.Application.CalculateFull())
            [void]($chart.Refresh())
        }
        catch {
            # Some Office builds do not expose chart refresh for embedded hosts;
            # saved chart caches are still verified by the visual harness.
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        Release-ComObject $worksheet
        $worksheet = $null
        Close-ChartWorkbook $workbook
        $workbook = $null

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        if ($worksheet -ne $null) { Release-ComObject $worksheet }
        if ($workbook -ne $null) { Close-ChartWorkbook $workbook }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function New-SparseBlankChartProbe($PowerPoint, $Cases) {
    $output = Join-Path $Cases "pptx-ladder-11-chart-sparse-blank-points-probe.pptx"
    $presentation = $null
    $workbooks = @()
    $worksheets = @()

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartSpecs = @(
            @{ Type = 65; X = 42;  Y = 36;  W = 300; H = 210 },
            @{ Type = 51; X = 378; Y = 36;  W = 300; H = 210 },
            @{ Type = 1;  X = 210; Y = 276; W = 300; H = 210 }
        )

        foreach ($spec in $chartSpecs) {
            $chartShape = $slide.Shapes.AddChart($spec.Type, $spec.X, $spec.Y, $spec.W, $spec.H)
            $chart = $chartShape.Chart
            [void]($chart.ChartData.Activate())
            $workbook = $chart.ChartData.Workbook
            $worksheet = $workbook.Worksheets.Item(1)
            $workbooks += $workbook
            $worksheets += $worksheet
            Populate-SparseChartWorksheet $worksheet
            [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$C`$6"))
            Configure-SparseProbeChart $chart $spec.Type
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        foreach ($worksheet in $worksheets) {
            Release-ComObject $worksheet
        }
        $worksheets = @()
        foreach ($workbook in $workbooks) {
            Close-ChartWorkbook $workbook
        }
        $workbooks = @()

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        foreach ($worksheet in $worksheets) {
            Release-ComObject $worksheet
        }
        foreach ($workbook in $workbooks) {
            Close-ChartWorkbook $workbook
        }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
            catch {
                # PowerPoint can already have torn down a failed presentation.
            }
        }
        Release-ComObject $presentation
    }

    return $output
}

function Populate-LegendKeyDataLabelWorksheet($Worksheet) {
    $Worksheet.Cells.Clear()
    $Worksheet.Cells.Item(1, 1).Value = "Category"
    $Worksheet.Cells.Item(1, 2).Value = "Demand"
    $Worksheet.Cells.Item(1, 3).Value = "Supply"
    $Worksheet.Cells.Item(2, 1).Value = "A"
    $Worksheet.Cells.Item(2, 2).Value = 4.0
    $Worksheet.Cells.Item(2, 3).Value = 7.0
    $Worksheet.Cells.Item(3, 1).Value = "B"
    $Worksheet.Cells.Item(3, 2).Value = 9.0
    $Worksheet.Cells.Item(3, 3).Value = 5.0
    $Worksheet.Cells.Item(4, 1).Value = "C"
    $Worksheet.Cells.Item(4, 2).Value = 6.0
    $Worksheet.Cells.Item(4, 3).Value = 11.0
    $Worksheet.Cells.Item(5, 1).Value = "D"
    $Worksheet.Cells.Item(5, 2).Value = 12.0
    $Worksheet.Cells.Item(5, 3).Value = 8.0
}

function Configure-LegendKeyDataLabelProbeChart($Chart, [int]$ChartType, [int]$DataLabelPosition) {
    [void]($Chart.HasTitle = $false)
    [void]($Chart.HasLegend = $false)
    [void]($Chart.ChartType = $ChartType)

    try {
        $valueAxis = $Chart.Axes(2, 1)
        $valueAxis.MinimumScale = 0
        $valueAxis.MaximumScale = 14
        $valueAxis.MajorUnit = 2
        $valueAxis.TickLabels.Font.Size = 7
        $valueAxis.Format.Line.ForeColor.RGB = Rgb 217 217 217
        $valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = Rgb 230 230 230
        $valueAxis.MajorGridlines.Format.Line.Weight = 0.5
    }
    catch {
        # Some chart kinds materialize axes lazily; the saved Office PDF is the oracle.
    }

    try {
        $categoryAxis = $Chart.Axes(1, 1)
        $categoryAxis.TickLabels.Font.Size = 7
        $categoryAxis.Format.Line.ForeColor.RGB = Rgb 217 217 217
    }
    catch {
        # Scatter axes do not expose category-axis settings in the same way.
    }

    $seriesCollection = $Chart.SeriesCollection()
    for ($i = 1; $i -le $seriesCollection.Count; $i++) {
        $series = $seriesCollection.Item($i)
        if ($i -eq 1) {
            $series.Format.Fill.ForeColor.RGB = Rgb 68 114 196
            $series.Format.Line.ForeColor.RGB = Rgb 68 114 196
            try { $series.MarkerStyle = 8 } catch {}
            try { $series.MarkerSize = 7 } catch {}
            try { $series.MarkerForegroundColor = Rgb 68 114 196 } catch {}
            try { $series.MarkerBackgroundColor = Rgb 68 114 196 } catch {}
        }
        else {
            $series.Format.Fill.ForeColor.RGB = Rgb 237 125 49
            $series.Format.Line.ForeColor.RGB = Rgb 237 125 49
            try { $series.MarkerStyle = 8 } catch {}
            try { $series.MarkerSize = 7 } catch {}
            try { $series.MarkerForegroundColor = Rgb 237 125 49 } catch {}
            try { $series.MarkerBackgroundColor = Rgb 237 125 49 } catch {}
        }

        try {
            [void]($series.ApplyDataLabels())
            $labels = $series.DataLabels()
            $labels.ShowValue = $true
            $labels.ShowSeriesName = $false
            $labels.ShowCategoryName = $false
            $labels.ShowLegendKey = $true
            $labels.Position = $DataLabelPosition
            $labels.Font.Size = 8
            Release-ComObject $labels
        }
        catch {
            # If a specific Office build refuses one label property, keep the
            # generated PPTX as evidence of that build's observable behavior.
        }

        Release-ComObject $series
    }
    Release-ComObject $seriesCollection
}

function New-CartesianLegendKeyDataLabelProbe($PowerPoint, $Cases) {
    $output = Join-Path $Cases "pptx-ladder-11-chart-data-label-legend-keys-probe.pptx"
    $presentation = $null
    $workbooks = @()
    $worksheets = @()

    try {
        $presentation = $PowerPoint.Presentations.Add($true)
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Background.Fill.ForeColor.RGB = Rgb 255 255 255

        $chartSpecs = @(
            @{ Type = 51; X = 144; Y = 78; W = 432; H = 324; Position = 2 }
        )

        foreach ($spec in $chartSpecs) {
            $chartShape = $slide.Shapes.AddChart($spec.Type, $spec.X, $spec.Y, $spec.W, $spec.H)
            $chart = $chartShape.Chart
            [void]($chart.ChartData.Activate())
            $workbook = $chart.ChartData.Workbook
            $worksheet = $workbook.Worksheets.Item(1)
            $workbooks += $workbook
            $worksheets += $worksheet
            Populate-LegendKeyDataLabelWorksheet $worksheet
            [void]($chart.SetSourceData("=Sheet1!`$A`$1:`$C`$5"))
            Configure-LegendKeyDataLabelProbeChart $chart $spec.Type $spec.Position
        }

        if (Test-Path -LiteralPath $output) {
            Remove-Item -LiteralPath $output -Force
        }

        [void]($presentation.SaveAs($output, 24))
        foreach ($worksheet in $worksheets) {
            Release-ComObject $worksheet
        }
        $worksheets = @()
        foreach ($workbook in $workbooks) {
            Close-ChartWorkbook $workbook
        }
        $workbooks = @()

        [void]($presentation.Close())
        $presentation = $null
    }
    finally {
        foreach ($worksheet in $worksheets) {
            Release-ComObject $worksheet
        }
        foreach ($workbook in $workbooks) {
            Close-ChartWorkbook $workbook
        }
        if ($presentation -ne $null) {
            try { [void]($presentation.Close()) }
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

    if ($AxisTitlesOnly) {
        $output = New-DefaultAxisTitleProbe `
            -PowerPoint $powerPoint `
            -Cases $cases
        $output = New-BarDefaultAxisTitleProbe `
            -PowerPoint $powerPoint `
            -Cases $cases
        $output = New-TopRightDefaultAxisTitleProbe `
            -PowerPoint $powerPoint `
            -Cases $cases
    }
    elseif ($DataLabelsOnly) {
        $output = New-PieDataLabelLeaderLineProbe `
            -PowerPoint $powerPoint `
            -Cases $cases `
            -MetadataRoot $MetadataRoot
        $output = New-PieDataLabelLeaderLineOffsetProbe `
            -PowerPoint $powerPoint `
            -Cases $cases `
            -MetadataRoot $MetadataRoot
        $output = New-CartesianLegendKeyDataLabelProbe `
            -PowerPoint $powerPoint `
            -Cases $cases
    }
    elseif ($SparseOnly) {
        $output = New-SparseBlankChartProbe `
            -PowerPoint $powerPoint `
            -Cases $cases
    }
    elseif (-not $DoughnutOnly) {
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
    [void]($series.Item(3).AxisGroup = 2)
    $series.Item(3).Format.Fill.ForeColor.RGB = Rgb 47 133 106
    [void]($series.Item(4).AxisGroup = 2)
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

    [void]($presentation.SaveAs($output, 24))
    [void]($presentation.Close())
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
        [void]($series.Item($i).AxisGroup = 2)
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

    [void]($presentation.SaveAs($output, 24))
    [void]($presentation.Close())
    $presentation = $null
    }

    if (-not $SparseOnly -and -not $DataLabelsOnly -and -not $AxisTitlesOnly) {
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
        -FileName "pptx-ladder-11-chart-doughnut-no-legend-exploded-probe.pptx" `
        -LegendPosition -4152 `
        -HasLegend $false `
        -IncludeInLayout $true `
        -Explosion 25
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
        -FileName "pptx-ladder-11-chart-doughnut-left-legend-exploded-probe.pptx" `
        -LegendPosition -4131 `
        -HasLegend $true `
        -IncludeInLayout $true `
        -Explosion 25
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
        -FileName "pptx-ladder-11-chart-doughnut-top-legend-exploded-probe.pptx" `
        -LegendPosition -4160 `
        -HasLegend $true `
        -IncludeInLayout $true `
        -Explosion 25
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
        -FileName "pptx-ladder-11-chart-doughnut-bottom-legend-exploded-probe.pptx" `
        -LegendPosition -4107 `
        -HasLegend $true `
        -IncludeInLayout $true `
        -Explosion 25
    $output = New-DoughnutLegendProbe `
        -PowerPoint $powerPoint `
        -Cases $cases `
        -FileName "pptx-ladder-11-chart-doughnut-right-overlay-legend-probe.pptx" `
        -LegendPosition -4152 `
        -HasLegend $true `
        -IncludeInLayout $false
    }
}
finally {
    if ($presentation -ne $null) {
        try { [void]($presentation.Close()) }
        catch {
            # COM cleanup should not mask the generation error that preceded it.
        }
    }
    if ($powerPoint -ne $null) {
        try { [void]($powerPoint.Quit()) }
        catch {
            # PowerPoint may already be unavailable after a COM failure.
        }
    }
    Release-ComObject $presentation
    Release-ComObject $powerPoint
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-Item -LiteralPath @($output)[-1]
