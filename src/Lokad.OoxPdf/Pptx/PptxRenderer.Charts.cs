using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderCharts(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        foreach (XElement frame in context.SlideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            XElement? graphicData = frame
                .Element(DrawingNamespace + "graphic")
                ?.Element(DrawingNamespace + "graphicData");
            if (graphicData?.Attribute("uri") is not { } uri ||
                !uri.Value.Contains("drawingml/2006/chart", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
            string? relationshipId = (string?)graphicData
                .Element(ChartNamespace + "chart")
                ?.Attribute(RelationshipsNamespace + "id");
            if (bounds is null || relationshipId is null || !context.SlideRelationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", context.Slide.PartName, context.SlideNumber, "Ignored");
                continue;
            }

            OoxPart? chartPart = context.Package.GetPart(relationship.ResolvedTarget);
            if (chartPart is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing and was ignored.", relationship.ResolvedTarget, context.SlideNumber, "Ignored");
                continue;
            }

            using Stream chartStream = chartPart.OpenRead();
            XDocument chartXml = SafeXml.Load(chartStream);
            if (TryRenderChartFallback(graphics, context.Document, context.Theme, bounds.Value, chartXml))
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_STATIC_FALLBACK", OoxPdfSeverity.Info, "PPTX chart was rendered with an approximate static chart fallback.", chartPart.Name, context.SlideNumber, "Static chart fallback");
                continue;
            }

            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, and pie chart cached numeric values have a static fallback.", chartPart.Name, context.SlideNumber, "Ignored");
        }
    }

    private static bool TryRenderChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml)
    {
        XElement? barChart = chartXml.Descendants(ChartNamespace + "barChart").FirstOrDefault();
        if (barChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> barSeries = ReadChartSeries(barChart);
            if (barSeries.Count != 0)
            {
                bool horizontalBars = string.Equals((string?)barChart.Element(ChartNamespace + "barDir")?.Attribute("val"), "bar", StringComparison.Ordinal);
                string grouping = (string?)barChart.Element(ChartNamespace + "grouping")?.Attribute("val") ?? "clustered";
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadChartSeriesFills(barChart, theme);
                RenderBarChartFallback(graphics, document, bounds, barSeries, horizontalBars, grouping, seriesFills);
                return true;
            }
        }

        XElement? lineChart = chartXml.Descendants(ChartNamespace + "lineChart").FirstOrDefault();
        if (lineChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> lineSeries = ReadChartSeries(lineChart);
            if (lineSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadChartSeriesStrokes(lineChart, theme);
                RenderLineChartFallback(graphics, document, bounds, lineSeries, seriesStrokes);
                return true;
            }
        }

        XElement? areaChart = chartXml.Descendants(ChartNamespace + "areaChart").FirstOrDefault();
        if (areaChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> areaSeries = ReadChartSeries(areaChart);
            if (areaSeries.Count != 0)
            {
                string grouping = (string?)areaChart.Element(ChartNamespace + "grouping")?.Attribute("val") ?? "standard";
                bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
                    string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
                RenderAreaChartFallback(graphics, document, bounds, areaSeries, stacked);
                return true;
            }
        }

        XElement? scatterChart = chartXml.Descendants(ChartNamespace + "scatterChart").FirstOrDefault();
        if (scatterChart is not null)
        {
            IReadOnlyList<ScatterSeries> scatterSeries = ReadScatterSeries(scatterChart, readBubbleSize: false);
            if (scatterSeries.Count != 0)
            {
                bool connectLines = ((string?)scatterChart.Element(ChartNamespace + "scatterStyle")?.Attribute("val"))?.Contains("Line", StringComparison.OrdinalIgnoreCase) == true;
                RenderScatterChartFallback(graphics, document, bounds, scatterSeries, connectLines, bubble: false);
                return true;
            }
        }

        XElement? bubbleChart = chartXml.Descendants(ChartNamespace + "bubbleChart").FirstOrDefault();
        if (bubbleChart is not null)
        {
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadScatterSeries(bubbleChart, readBubbleSize: true);
            if (bubbleSeries.Count != 0)
            {
                RenderScatterChartFallback(graphics, document, bounds, bubbleSeries, connectLines: false, bubble: true);
                return true;
            }
        }

        IReadOnlyList<IReadOnlyList<double>> radarSeries = ReadChartSeries(chartXml, "radarChart");
        if (radarSeries.Count != 0)
        {
            RenderRadarChartFallback(graphics, document, bounds, radarSeries);
            return true;
        }

        IReadOnlyList<IReadOnlyList<double>> pieSeries = ReadChartSeries(chartXml, "pieChart");
        if (pieSeries.Count != 0)
        {
            RenderPieChartFallback(graphics, document, bounds, pieSeries[0]);
            return true;
        }

        IReadOnlyList<IReadOnlyList<double>> doughnutSeries = ReadChartSeries(chartXml, "doughnutChart");
        if (doughnutSeries.Count != 0)
        {
            RenderDoughnutChartFallback(graphics, document, bounds, doughnutSeries[0]);
            return true;
        }

        return false;
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XDocument chartXml, string chartElementName)
    {
        XElement? chartElement = chartXml.Descendants(ChartNamespace + chartElementName).FirstOrDefault();
        return chartElement is null ? [] : ReadChartSeries(chartElement);
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XElement chartElement)
    {
        var series = new List<IReadOnlyList<double>>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            double[] values = element
                .Elements(ChartNamespace + "val")
                .Descendants(ChartNamespace + "pt")
                .Select(point => (string?)point.Element(ChartNamespace + "v"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
                .Where(value => !double.IsNaN(value))
                .ToArray();
            if (values.Length > 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static IReadOnlyList<ChartSeriesFill?> ReadChartSeriesFills(XElement chartElement, PptxTheme theme)
    {
        var fills = new List<ChartSeriesFill?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            fills.Add(TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha)
                ? new ChartSeriesFill(color, alpha)
                : null);
        }

        return fills;
    }

    private static IReadOnlyList<ChartSeriesStroke?> ReadChartSeriesStrokes(XElement chartElement, PptxTheme theme)
    {
        var strokes = new List<ChartSeriesStroke?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            strokes.Add(shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha)
                    ? new ChartSeriesStroke(color, alpha, width)
                    : null);
        }

        return strokes;
    }

    private static IReadOnlyList<ScatterSeries> ReadScatterSeries(XElement chartElement, bool readBubbleSize)
    {
        var series = new List<ScatterSeries>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            double[] xValues = ReadCachedNumbers(element, "xVal");
            double[] yValues = ReadCachedNumbers(element, "yVal");
            double[] bubbleSizes = readBubbleSize ? ReadCachedNumbers(element, "bubbleSize") : [];
            int count = Math.Min(xValues.Length, yValues.Length);
            if (count == 0)
            {
                continue;
            }

            var points = new ScatterPoint[count];
            for (int i = 0; i < count; i++)
            {
                double size = i < bubbleSizes.Length ? bubbleSizes[i] : 1d;
                points[i] = new ScatterPoint(xValues[i], yValues[i], size);
            }

            series.Add(new ScatterSeries(points));
        }

        return series;
    }

    private static double[] ReadCachedNumbers(XElement element, string containerName)
    {
        return element
            .Elements(ChartNamespace + containerName)
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
            .Where(value => !double.IsNaN(value))
            .ToArray();
    }

    private static RgbColor ChartPalette(int index)
    {
        RgbColor[] palette =
        [
            new RgbColor(68, 114, 196),
            new RgbColor(237, 125, 49),
            new RgbColor(165, 165, 165),
            new RgbColor(255, 192, 0),
            new RgbColor(91, 155, 213),
            new RgbColor(112, 173, 71)
        ];
        return palette[index % palette.Length];
    }

    private static void RenderBarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool horizontalBars, string grouping, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.1d;
        double plotY = y + height * 0.14d;
        double plotWidth = width * 0.82d;
        double plotHeight = height * 0.72d;
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
            string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        bool percentStacked = string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        (double minValue, double maxValue) = stacked
            ? GetStackedValueExtents(series, categoryCount, percentStacked)
            : GetClusteredValueExtents(series);
        double valueRange = Math.Max(1d, maxValue - minValue);

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        double zeroX = plotX + (-minValue) / valueRange * plotWidth;
        double zeroY = plotY + (-minValue) / valueRange * plotHeight;
        graphics.StrokeLine(plotX, zeroY, plotX + plotWidth, zeroY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        if (horizontalBars)
        {
            if (stacked)
            {
                RenderStackedHorizontalBars(graphics, plotY, plotWidth, plotHeight, series, categoryCount, valueRange, zeroX, percentStacked, seriesFills);
            }
            else
            {
                RenderClusteredHorizontalBars(graphics, plotY, plotWidth, plotHeight, series, categoryCount, valueRange, zeroX, seriesFills);
            }

            return;
        }

        if (stacked)
        {
            RenderStackedColumns(graphics, plotX, plotWidth, plotHeight, series, categoryCount, valueRange, zeroY, percentStacked, seriesFills);
            return;
        }

        double categoryWidth = plotWidth / categoryCount;
        double barSlot = categoryWidth * 0.82d / Math.Max(1, series.Count);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + categoryWidth * 0.09d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = values[category];
                double barHeight = Math.Abs(value) / valueRange * plotHeight;
                ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
                FillChartRectangle(graphics, categoryX + seriesIndex * barSlot, value >= 0d ? zeroY : zeroY - barHeight, Math.Max(0.5d, barSlot * 0.86d), barHeight, fill);
            }
        }
    }

    private static ChartSeriesFill ChartSeriesColor(int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(seriesIndex), 1d);
    }

    private static void FillChartRectangle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
        graphics.FillRectangle(x, y, width, height);

        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static (double Min, double Max) GetClusteredValueExtents(IReadOnlyList<IReadOnlyList<double>> series)
    {
        double maxValue = Math.Max(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        return (minValue, maxValue);
    }

    private static (double Min, double Max) GetStackedValueExtents(IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, bool percentStacked)
    {
        if (percentStacked)
        {
            return (0d, 1d);
        }

        double minValue = 0d;
        double maxValue = 0d;
        for (int category = 0; category < categoryCount; category++)
        {
            double positive = 0d;
            double negative = 0d;
            foreach (IReadOnlyList<double> values in series)
            {
                if (category >= values.Count)
                {
                    continue;
                }

                if (values[category] >= 0d)
                {
                    positive += values[category];
                }
                else
                {
                    negative += values[category];
                }
            }

            maxValue = Math.Max(maxValue, positive);
            minValue = Math.Min(minValue, negative);
        }

        return (minValue, maxValue);
    }

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroX, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barSlot = categoryHeight * 0.82d / Math.Max(1, series.Count);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + plotHeight - (category + 1) * categoryHeight + categoryHeight * 0.09d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = values[category];
                double barWidth = Math.Abs(value) / valueRange * plotWidth;
                ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
                FillChartRectangle(graphics, value >= 0d ? zeroX : zeroX - barWidth, categoryY + seriesIndex * barSlot, barWidth, Math.Max(0.5d, barSlot * 0.86d), fill);
            }
        }
    }

    private static void RenderStackedColumns(PdfGraphicsBuilder graphics, double plotX, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroY, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        double categoryWidth = plotWidth / categoryCount;
        double barWidth = Math.Max(0.5d, categoryWidth * 0.58d);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + (categoryWidth - barWidth) / 2d;
            double positiveY = zeroY;
            double negativeY = zeroY;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                double segmentHeight = Math.Abs(value) / valueRange * plotHeight;
                ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
                if (value >= 0d)
                {
                    FillChartRectangle(graphics, categoryX, positiveY, barWidth, segmentHeight, fill);
                    positiveY += segmentHeight;
                }
                else
                {
                    negativeY -= segmentHeight;
                    FillChartRectangle(graphics, categoryX, negativeY, barWidth, segmentHeight, fill);
                }
            }
        }
    }

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroX, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = Math.Max(0.5d, categoryHeight * 0.58d);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + plotHeight - (category + 1) * categoryHeight + (categoryHeight - barHeight) / 2d;
            double positiveX = zeroX;
            double negativeX = zeroX;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                double segmentWidth = Math.Abs(value) / valueRange * plotWidth;
                ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
                if (value >= 0d)
                {
                    FillChartRectangle(graphics, positiveX, categoryY, segmentWidth, barHeight, fill);
                    positiveX += segmentWidth;
                }
                else
                {
                    negativeX -= segmentWidth;
                    FillChartRectangle(graphics, negativeX, categoryY, segmentWidth, barHeight, fill);
                }
            }
        }
    }

    private static double GetCategoryPositiveTotal(IReadOnlyList<IReadOnlyList<double>> series, int category, bool percentStacked)
    {
        if (!percentStacked)
        {
            return 1d;
        }

        return Math.Max(1d, series
            .Where(values => category < values.Count)
            .Select(values => Math.Max(0d, values[category]))
            .Sum());
    }

    private static double NormalizeStackedValue(double value, double positiveTotal, bool percentStacked)
    {
        return percentStacked && value > 0d ? value / positiveTotal : value;
    }

    private static void RenderLineChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        double valueRange = Math.Max(1d, maxValue - minValue);

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.5d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            double previousX = 0d;
            double previousY = 0d;
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double pointY = plotY + (values[i] - minValue) / valueRange * plotHeight;
                if (i > 0)
                {
                    graphics.StrokeLine(previousX, previousY, pointX, pointY);
                }

                graphics.SetFillRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
                graphics.FillEllipse(pointX - 2d, pointY - 2d, 4d, 4d);
                previousX = pointX;
                previousY = pointY;
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ChartSeriesStroke ChartSeriesStrokeColor(int seriesIndex, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, double defaultWidth)
    {
        return seriesIndex < seriesStrokes.Count && seriesStrokes[seriesIndex] is { } stroke
            ? stroke
            : new ChartSeriesStroke(ChartPalette(seriesIndex), 1d, defaultWidth);
    }

    private static void SetChartStroke(PdfGraphicsBuilder graphics, ChartSeriesStroke stroke)
    {
        graphics.SetStrokeRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
        graphics.SetLineWidth(stroke.Width);
    }

    private static void RenderAreaChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool stacked)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = stacked ? GetMaxStackedLineValue(series, pointCount) : Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        double valueRange = Math.Max(1d, maxValue - minValue);
        double zeroY = plotY + (-minValue) / valueRange * plotHeight;

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, zeroY, plotX + plotWidth, zeroY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            var upperPoints = new (double X, double Y)[pointCount];
            var lowerPoints = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double value = i < values.Count ? values[i] : 0d;
                double lowerValue = stacked ? lower[i] : 0d;
                double upperValue = stacked ? lower[i] + value : value;
                upperPoints[i] = (pointX, plotY + (upperValue - minValue) / valueRange * plotHeight);
                lowerPoints[i] = (pointX, plotY + (lowerValue - minValue) / valueRange * plotHeight);
                if (stacked)
                {
                    lower[i] = upperValue;
                }
            }

            var polygon = new (double X, double Y)[pointCount * 2];
            for (int i = 0; i < pointCount; i++)
            {
                polygon[i] = upperPoints[i];
                polygon[polygon.Length - i - 1] = lowerPoints[i];
            }

            RgbColor fill = ChartPalette(seriesIndex);
            graphics.SaveState();
            graphics.SetAlpha(0.62d, 1d);
            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            graphics.FillPolygon(polygon);
            graphics.RestoreState();
            graphics.SetStrokeRgb(fill.Red, fill.Green, fill.Blue);
            graphics.SetLineWidth(1.2d);
            for (int i = 1; i < upperPoints.Length; i++)
            {
                graphics.StrokeLine(upperPoints[i - 1].X, upperPoints[i - 1].Y, upperPoints[i].X, upperPoints[i].Y);
            }
        }
    }

    private static double GetMaxStackedLineValue(IReadOnlyList<IReadOnlyList<double>> series, int pointCount)
    {
        double maxValue = 1d;
        for (int i = 0; i < pointCount; i++)
        {
            double sum = 0d;
            foreach (IReadOnlyList<double> values in series)
            {
                if (i < values.Count)
                {
                    sum += Math.Max(0d, values[i]);
                }
            }

            maxValue = Math.Max(maxValue, sum);
        }

        return maxValue;
    }

    private static void RenderScatterChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        double minX = series.SelectMany(item => item.Points).Min(point => point.X);
        double maxX = series.SelectMany(item => item.Points).Max(point => point.X);
        double minY = Math.Min(0d, series.SelectMany(item => item.Points).Min(point => point.Y));
        double maxY = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Y));
        double xRange = Math.Max(1d, maxX - minX);
        double yRange = Math.Max(1d, maxY - minY);
        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Size));

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            RgbColor color = ChartPalette(seriesIndex);
            graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.SetLineWidth(1.2d);
            (double X, double Y)? previous = null;
            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                double pointX = plotX + (point.X - minX) / xRange * plotWidth;
                double pointY = plotY + (point.Y - minY) / yRange * plotHeight;
                if (connectLines && previous is { } prior)
                {
                    graphics.StrokeLine(prior.X, prior.Y, pointX, pointY);
                }

                double radius = bubble ? 3d + Math.Sqrt(Math.Max(0d, point.Size) / maxBubbleSize) * 8d : 3d;
                graphics.FillEllipse(pointX - radius, pointY - radius, radius * 2d, radius * 2d);
                previous = (pointX, pointY);
            }
        }
    }

    private static void RenderRadarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double centerX = x + width * 0.5d;
        double centerY = y + height * 0.52d;
        double radius = Math.Min(width, height) * 0.32d;
        int pointCount = Math.Max(3, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());

        graphics.SetStrokeRgb(180, 180, 180);
        graphics.SetLineWidth(0.5d);
        for (int i = 0; i < pointCount; i++)
        {
            double angle = -Math.PI / 2d + i * Math.PI * 2d / pointCount;
            graphics.StrokeLine(centerX, centerY, centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
        }

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count ? Math.Max(0d, values[i]) : 0d;
                double pointRadius = value / maxValue * radius;
                double angle = -Math.PI / 2d + i * Math.PI * 2d / pointCount;
                points[i] = (centerX + Math.Cos(angle) * pointRadius, centerY + Math.Sin(angle) * pointRadius);
            }

            RgbColor color = ChartPalette(seriesIndex);
            graphics.SaveState();
            graphics.SetAlpha(series.Count == 1 ? 0.40d : 0.18d, 1d);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillPolygon(points);
            graphics.RestoreState();
            graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
            graphics.SetLineWidth(1.2d);
            for (int i = 0; i < points.Length; i++)
            {
                (double X, double Y) a = points[i];
                (double X, double Y) b = points[(i + 1) % points.Length];
                graphics.StrokeLine(a.X, a.Y, b.X, b.Y);
            }
        }
    }

    private static void RenderPieChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values)
    {
        double total = values.Where(value => value > 0d).Sum();
        if (total <= 0d)
        {
            return;
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double angle = -Math.PI / 2d;

        for (int i = 0; i < values.Count; i++)
        {
            double value = Math.Max(0d, values[i]);
            if (value <= 0d)
            {
                continue;
            }

            double sweep = value / total * Math.PI * 2d;
            int segments = Math.Max(2, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 18d)));
            var points = new (double X, double Y)[segments + 2];
            points[0] = (centerX, centerY);
            for (int segment = 0; segment <= segments; segment++)
            {
                double segmentAngle = angle + sweep * segment / segments;
                points[segment + 1] = (
                    centerX + Math.Cos(segmentAngle) * radius,
                    centerY + Math.Sin(segmentAngle) * radius);
            }

            RgbColor fill = ChartPalette(i);
            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            graphics.FillPolygon(points);
            angle += sweep;
        }
    }

    private static void RenderDoughnutChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values)
    {
        RenderPieChartFallback(graphics, document, bounds, values);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double innerRadius = radius * 0.56d;
        graphics.SetFillRgb(255, 255, 255);
        graphics.FillEllipse(centerX - innerRadius, centerY - innerRadius, innerRadius * 2d, innerRadius * 2d);
    }

    private static void EmitChartDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, string id, OoxPdfSeverity severity, string message, string? partName, int slideIndex, string fallback)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            severity,
            message,
            partName,
            SlideIndex: slideIndex,
            Feature: "chart",
            Fallback: fallback));
    }

    private readonly record struct ScatterSeries(IReadOnlyList<ScatterPoint> Points);

    private readonly record struct ScatterPoint(double X, double Y, double Size);

    private readonly record struct ChartSeriesFill(RgbColor Color, double Alpha);

    private readonly record struct ChartSeriesStroke(RgbColor Color, double Alpha, double Width);
}
