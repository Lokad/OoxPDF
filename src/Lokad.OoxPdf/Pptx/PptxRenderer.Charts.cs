using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PdfFontResource> RenderCharts(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        var fonts = new List<PdfFontResource>();
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
            if (TryRenderChartFallback(graphics, context.Document, context.Theme, bounds.Value, chartXml, fonts))
            {
                fonts.AddRange(RenderChartTitle(context.Document, context.Theme, graphics, bounds.Value, chartXml));
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_STATIC_FALLBACK", OoxPdfSeverity.Info, "PPTX chart was rendered with an approximate static chart fallback.", chartPart.Name, context.SlideNumber, "Static chart fallback");
                continue;
            }

            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut chart cached numeric values have a static fallback.", chartPart.Name, context.SlideNumber, "Ignored");
        }

        return fonts;
    }

    private static bool TryRenderChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, List<PdfFontResource> fonts)
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
                ChartAxesStyle axesStyle = ReadChartAxesStyle(chartXml, theme);
                ChartShapeStyle plotAreaStyle = ReadChartPlotAreaStyle(chartXml, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderBarChartFallback(graphics, document, bounds, barSeries, horizontalBars, grouping, seriesFills, HasMajorGridlines(chartXml), HasMinorGridlines(chartXml), axesStyle, plotAreaStyle);
                fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, GetBarChartPlotBox(document, bounds), ReadChartCategoryLabels(barChart), horizontalBars));
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
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadChartMarkerStyles(lineChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadChartSeriesSmooth(lineChart);
                ChartAxesStyle axesStyle = ReadChartAxesStyle(chartXml, theme);
                ChartShapeStyle plotAreaStyle = ReadChartPlotAreaStyle(chartXml, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderLineChartFallback(graphics, document, bounds, lineSeries, seriesStrokes, markerStyles, smoothSeries, HasMajorGridlines(chartXml), HasMinorGridlines(chartXml), axesStyle, plotAreaStyle);
                fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, GetLineChartPlotBox(document, bounds), ReadChartCategoryLabels(lineChart), horizontalBars: false));
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
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadChartSeriesFills(areaChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadChartSeriesStrokes(areaChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderAreaChartFallback(graphics, document, bounds, areaSeries, stacked, seriesFills, seriesStrokes);
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
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadChartSeriesFills(scatterChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadChartSeriesStrokes(scatterChart, theme);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadChartMarkerStyles(scatterChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadChartSeriesSmooth(scatterChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderScatterChartFallback(graphics, document, bounds, scatterSeries, connectLines, bubble: false, seriesFills, seriesStrokes, markerStyles, smoothSeries);
                return true;
            }
        }

        XElement? bubbleChart = chartXml.Descendants(ChartNamespace + "bubbleChart").FirstOrDefault();
        if (bubbleChart is not null)
        {
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadScatterSeries(bubbleChart, readBubbleSize: true);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadChartSeriesFills(bubbleChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadChartSeriesStrokes(bubbleChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderScatterChartFallback(graphics, document, bounds, bubbleSeries, connectLines: false, bubble: true, seriesFills, seriesStrokes, [], []);
                return true;
            }
        }

        XElement? radarChart = chartXml.Descendants(ChartNamespace + "radarChart").FirstOrDefault();
        if (radarChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> radarSeries = ReadChartSeries(radarChart);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadChartSeriesFills(radarChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadChartSeriesStrokes(radarChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderRadarChartFallback(graphics, document, bounds, radarSeries, seriesFills, seriesStrokes);
                return true;
            }
        }

        XElement? pieChart = chartXml.Descendants(ChartNamespace + "pieChart").FirstOrDefault();
        if (pieChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> pieSeries = ReadChartSeries(pieChart);
            if (pieSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadChartPointFills(pieChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadChartPointStrokes(pieChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadChartPointExplosions(pieChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderPieChartFallback(graphics, document, bounds, pieSeries[0], pointFills, pointStrokes, pointExplosions);
                return true;
            }
        }

        XElement? doughnutChart = chartXml.Descendants(ChartNamespace + "doughnutChart").FirstOrDefault();
        if (doughnutChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> doughnutSeries = ReadChartSeries(doughnutChart);
            if (doughnutSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadChartPointFills(doughnutChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadChartPointStrokes(doughnutChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadChartPointExplosions(doughnutChart);
                double holeSize = ReadDoughnutHoleSize(doughnutChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, theme);
                RenderDoughnutChartFallback(graphics, document, bounds, doughnutSeries[0], pointFills, pointStrokes, pointExplosions, holeSize);
                return true;
            }
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

    private static void RenderChartAreaStyle(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxTheme theme)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        RenderChartShapeStyle(graphics, x, y, width, height, ReadChartShapeStyle(chartXml.Root?.Element(ChartNamespace + "spPr"), theme));
    }

    private static ChartShapeStyle ReadChartPlotAreaStyle(XDocument chartXml, PptxTheme theme)
    {
        XElement? shapeProperties = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "spPr");
        return ReadChartShapeStyle(shapeProperties, theme);
    }

    private static ChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme)
    {
        if (shapeProperties is null)
        {
            return ChartShapeStyle.Empty;
        }

        ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new ChartSeriesFill(fillColor, fillAlpha)
            : null;
        ChartSeriesStroke? stroke = TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double lineWidth, out double strokeAlpha)
            ? new ChartSeriesStroke(strokeColor, strokeAlpha, lineWidth)
            : null;
        return new ChartShapeStyle(fill, stroke);
    }

    private static void RenderChartShapeStyle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartShapeStyle style)
    {
        if (style.Fill is { } fill)
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

        if (style.Stroke is { } stroke)
        {
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            graphics.SetStrokeRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
            graphics.SetLineWidth(stroke.Width);
            graphics.StrokeRectangle(x, y, width, height);
            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static IReadOnlyList<PdfFontResource> RenderChartTitle(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, XDocument chartXml)
    {
        string? title = ReadChartTitleText(chartXml);
        if (string.IsNullOrWhiteSpace(title))
        {
            return [];
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double fontSize = 12d;
        RgbColor color = theme.TryResolveColor("tx1", out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        var run = new TextRun(
            title.Trim(),
            x + width * 0.08d,
            y + height * 0.88d,
            width * 0.84d,
            fontSize * 1.4d,
            x,
            y,
            width,
            height,
            fontSize,
            0d,
            0d,
            color,
            1d,
            null,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            KerningEnabled: true,
            TextAlignment.Center,
            FontFamily: null,
            RotationDegrees: 0d,
            RotationCenterX: 0d,
            RotationCenterY: 0d,
            FlipHorizontal: false,
            FlipVertical: false);
        return RenderTextRuns([run], graphics);
    }

    private static string? ReadChartTitleText(XDocument chartXml)
    {
        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        if (title is null)
        {
            return null;
        }

        string text = string.Concat(title
            .Descendants(DrawingNamespace + "t")
            .Select(element => (string?)element)
            .Where(value => !string.IsNullOrEmpty(value)));
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return (string?)title
            .Descendants(ChartNamespace + "v")
            .FirstOrDefault();
    }

    private static IReadOnlyList<string> ReadChartCategoryLabels(XElement chartElement)
    {
        XElement? categories = chartElement
            .Element(ChartNamespace + "ser")
            ?.Element(ChartNamespace + "cat");
        if (categories is null)
        {
            return [];
        }

        return categories
            .Descendants(ChartNamespace + "pt")
            .Select(point => ((string?)point.Element(ChartNamespace + "v"))?.Trim())
            .Where(value => !string.IsNullOrEmpty(value))
            .Cast<string>()
            .ToArray();
    }

    private static IReadOnlyList<PdfFontResource> RenderChartCategoryLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<string> labels, bool horizontalBars)
    {
        if (labels.Count == 0)
        {
            return [];
        }

        double fontSize = 9d;
        RgbColor color = theme.TryResolveColor("tx1", out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            double x;
            double y;
            double width;
            double height = fontSize * 1.35d;
            TextAlignment alignment;
            if (horizontalBars)
            {
                double slotHeight = plotBox.Height / labels.Count;
                x = Math.Max(0d, plotBox.X - plotBox.Width * 0.18d);
                y = plotBox.Y + slotHeight * (i + 0.5d) - height * 0.45d;
                width = plotBox.Width * 0.16d;
                alignment = TextAlignment.Right;
            }
            else
            {
                double slotWidth = plotBox.Width / labels.Count;
                x = plotBox.X + slotWidth * i;
                y = plotBox.Y - height * 1.25d;
                width = slotWidth;
                alignment = TextAlignment.Center;
            }

            runs.Add(new TextRun(
                labels[i],
                x,
                y,
                Math.Max(1d, width),
                height,
                plotBox.X,
                plotBox.Y,
                plotBox.Width,
                plotBox.Height,
                fontSize,
                0d,
                0d,
                color,
                1d,
                null,
                Bold: false,
                Italic: false,
                Underline: false,
                Strike: false,
                KerningEnabled: true,
                alignment,
                FontFamily: null,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics);
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

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadChartPointFills(XElement chartElement, PptxTheme theme)
    {
        var fills = new Dictionary<int, ChartSeriesFill>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return fills;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha))
            {
                fills[index] = new ChartSeriesFill(color, alpha);
            }
        }

        return fills;
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadChartPointStrokes(XElement chartElement, PptxTheme theme)
    {
        var strokes = new Dictionary<int, ChartSeriesStroke>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return strokes;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha))
            {
                strokes[index] = new ChartSeriesStroke(color, alpha, width);
            }
        }

        return strokes;
    }

    private static IReadOnlyDictionary<int, double> ReadChartPointExplosions(XElement chartElement)
    {
        var explosions = new Dictionary<int, double>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return explosions;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            if (point.Element(ChartNamespace + "explosion")?.Attribute("val") is { } explosionAttribute &&
                double.TryParse(explosionAttribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double explosion))
            {
                explosions[index] = Math.Clamp(explosion / 100d, 0d, 1d);
            }
        }

        return explosions;
    }

    private static double ReadDoughnutHoleSize(XElement doughnutChart)
    {
        if (doughnutChart.Element(ChartNamespace + "holeSize")?.Attribute("val") is { } value &&
            double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return Math.Clamp(parsed / 100d, 0.1d, 0.9d);
        }

        return 0.56d;
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

    private static IReadOnlyList<ChartMarkerStyle> ReadChartMarkerStyles(XElement chartElement, PptxTheme theme)
    {
        var styles = new List<ChartMarkerStyle>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? marker = element.Element(ChartNamespace + "marker");
            string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ?? "circle";
            double size = marker?.Element(ChartNamespace + "size")?.Attribute("val") is { } value &&
                double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                    ? Math.Clamp(parsed, 2d, 30d)
                    : 4d;
            XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
            ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
                ? new ChartSeriesFill(fillColor, fillAlpha)
                : null;
            ChartSeriesStroke? stroke = shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double strokeWidth, out double strokeAlpha)
                    ? new ChartSeriesStroke(strokeColor, strokeAlpha, strokeWidth)
                    : null;
            styles.Add(new ChartMarkerStyle(symbol, size, fill, stroke));
        }

        return styles;
    }

    private static IReadOnlyList<bool> ReadChartSeriesSmooth(XElement chartElement)
    {
        var smooth = new List<bool>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            string? value = (string?)element.Element(ChartNamespace + "smooth")?.Attribute("val");
            smooth.Add(value is "1" or "true");
        }

        return smooth;
    }

    private static bool HasMajorGridlines(XDocument chartXml)
    {
        return chartXml.Descendants(ChartNamespace + "majorGridlines").Any();
    }

    private static bool HasMinorGridlines(XDocument chartXml)
    {
        return chartXml.Descendants(ChartNamespace + "minorGridlines").Any();
    }

    private static ChartAxesStyle ReadChartAxesStyle(XDocument chartXml, PptxTheme theme)
    {
        ChartSeriesStroke? valueAxis = ReadChartAxisStroke(chartXml, "valAx", theme);
        ChartSeriesStroke? categoryAxis = ReadChartAxisStroke(chartXml, "catAx", theme);
        return new ChartAxesStyle(valueAxis, categoryAxis);
    }

    private static ChartSeriesStroke? ReadChartAxisStroke(XDocument chartXml, string axisName, PptxTheme theme)
    {
        XElement? shapeProperties = chartXml
            .Descendants(ChartNamespace + axisName)
            .FirstOrDefault()
            ?.Element(ChartNamespace + "spPr");
        return shapeProperties is not null &&
            TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha)
                ? new ChartSeriesStroke(color, alpha, width)
                : null;
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

    private static void RenderBarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool horizontalBars, string grouping, IReadOnlyList<ChartSeriesFill?> seriesFills, bool majorGridlines, bool minorGridlines, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle)
    {
        ChartPlotBox plotBox = GetBarChartPlotBox(document, bounds);
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotX, plotY, plotWidth, plotHeight, plotAreaStyle);
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
            string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        bool percentStacked = string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        (double minValue, double maxValue) = stacked
            ? GetStackedValueExtents(series, categoryCount, percentStacked)
            : GetClusteredValueExtents(series);
        double valueRange = Math.Max(1d, maxValue - minValue);

        double zeroX = plotX + (-minValue) / valueRange * plotWidth;
        double zeroY = plotY + (-minValue) / valueRange * plotHeight;
        if (minorGridlines)
        {
            if (horizontalBars)
            {
                DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: false);
            }
            else
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: false);
            }
        }

        if (majorGridlines)
        {
            if (horizontalBars)
            {
                DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: true);
            }
            else
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: true);
            }
        }

        ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        SetChartStroke(graphics, horizontalBars ? valueAxisStroke : categoryAxisStroke);
        graphics.StrokeLine(plotX, zeroY, plotX + plotWidth, zeroY);
        SetChartStroke(graphics, horizontalBars ? categoryAxisStroke : valueAxisStroke);
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

    private static ChartPlotBox GetBarChartPlotBox(PptxDocument document, ShapeBounds bounds)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        return new ChartPlotBox(x + width * 0.1d, y + height * 0.14d, width * 0.82d, height * 0.72d);
    }

    private static void DrawHorizontalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, bool major)
    {
        graphics.SetStrokeRgb(major ? (byte)217 : (byte)235, major ? (byte)217 : (byte)235, major ? (byte)217 : (byte)235);
        graphics.SetLineWidth(major ? 0.5d : 0.25d);
        int divisions = major ? 4 : 8;
        for (int i = 1; i < divisions; i++)
        {
            if (!major && i % 2 == 0)
            {
                continue;
            }

            double y = plotY + plotHeight * i / divisions;
            graphics.StrokeLine(plotX, y, plotX + plotWidth, y);
        }
    }

    private static void DrawVerticalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, bool major)
    {
        graphics.SetStrokeRgb(major ? (byte)217 : (byte)235, major ? (byte)217 : (byte)235, major ? (byte)217 : (byte)235);
        graphics.SetLineWidth(major ? 0.5d : 0.25d);
        int divisions = major ? 4 : 8;
        for (int i = 1; i < divisions; i++)
        {
            if (!major && i % 2 == 0)
            {
                continue;
            }

            double x = plotX + plotWidth * i / divisions;
            graphics.StrokeLine(x, plotY, x, plotY + plotHeight);
        }
    }

    private static ChartSeriesFill ChartSeriesColor(int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha = 1d)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(seriesIndex), defaultAlpha);
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

    private static void RenderLineChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries, bool majorGridlines, bool minorGridlines, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle)
    {
        ChartPlotBox plotBox = GetLineChartPlotBox(document, bounds);
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotX, plotY, plotWidth, plotHeight, plotAreaStyle);
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        double valueRange = Math.Max(1d, maxValue - minValue);

        if (minorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: false);
        }

        if (majorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, major: true);
        }

        ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        SetChartStroke(graphics, categoryAxisStroke);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        SetChartStroke(graphics, valueAxisStroke);
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
            var points = new List<(double X, double Y)>(values.Count);
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double pointY = plotY + (values[i] - minValue) / valueRange * plotHeight;
                points.Add((pointX, pointY));
            }

            if (IsSmoothSeries(seriesIndex, smoothSeries))
            {
                StrokeSmoothChartPath(graphics, points);
            }
            else
            {
                StrokeStraightChartPath(graphics, points);
            }

            foreach ((double pointX, double pointY) in points)
            {
                graphics.SetFillRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
                DrawChartMarker(graphics, pointX, pointY, ChartMarker(seriesIndex, markerStyles), stroke.Color, stroke.Color);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ChartPlotBox GetLineChartPlotBox(PptxDocument document, ShapeBounds bounds)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        return new ChartPlotBox(x + width * 0.12d, y + height * 0.16d, width * 0.76d, height * 0.68d);
    }

    private static bool IsSmoothSeries(int seriesIndex, IReadOnlyList<bool> smoothSeries)
    {
        return seriesIndex < smoothSeries.Count && smoothSeries[seriesIndex];
    }

    private static void StrokeStraightChartPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        for (int i = 1; i < points.Count; i++)
        {
            graphics.StrokeLine(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
        }
    }

    private static void StrokeSmoothChartPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count < 2)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 0; i < points.Count - 1; i++)
        {
            (double X, double Y) p0 = i == 0 ? points[i] : points[i - 1];
            (double X, double Y) p1 = points[i];
            (double X, double Y) p2 = points[i + 1];
            (double X, double Y) p3 = i + 2 < points.Count ? points[i + 2] : points[i + 1];
            graphics.CurveTo(
                p1.X + (p2.X - p0.X) / 6d,
                p1.Y + (p2.Y - p0.Y) / 6d,
                p2.X - (p3.X - p1.X) / 6d,
                p2.Y - (p3.Y - p1.Y) / 6d,
                p2.X,
                p2.Y);
        }

        graphics.StrokeCurrentPath();
    }

    private static ChartMarkerStyle ChartMarker(int seriesIndex, IReadOnlyList<ChartMarkerStyle> markerStyles)
    {
        return seriesIndex < markerStyles.Count ? markerStyles[seriesIndex] : ChartMarkerStyle.Default;
    }

    private static void DrawChartMarker(PdfGraphicsBuilder graphics, double x, double y, ChartMarkerStyle marker, RgbColor defaultFill, RgbColor defaultStroke)
    {
        if (string.Equals(marker.Symbol, "none", StringComparison.Ordinal))
        {
            return;
        }

        double size = marker.Size;
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d);
        ChartSeriesStroke? stroke = marker.Stroke;
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
        switch (marker.Symbol)
        {
            case "square":
                graphics.FillRectangle(x - size / 2d, y - size / 2d, size, size);
                break;
            case "diamond":
                graphics.FillPolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y),
                    (x, y - size / 2d),
                    (x - size / 2d, y)
                ]);
                break;
            case "triangle":
                graphics.FillPolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y - size / 2d),
                    (x - size / 2d, y - size / 2d)
                ]);
                break;
            default:
                graphics.FillEllipse(x - size / 2d, y - size / 2d, size, size);
                break;
        }

        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        if (stroke is not { } markerStroke)
        {
            return;
        }

        if (markerStroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, markerStroke.Alpha);
        }

        SetChartStroke(graphics, markerStroke);
        switch (marker.Symbol)
        {
            case "square":
                graphics.StrokeRectangle(x - size / 2d, y - size / 2d, size, size);
                break;
            case "diamond":
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y),
                    (x, y - size / 2d),
                    (x - size / 2d, y)
                ]);
                break;
            case "triangle":
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y - size / 2d),
                    (x - size / 2d, y - size / 2d)
                ]);
                break;
            default:
                graphics.StrokeEllipse(x - size / 2d, y - size / 2d, size, size);
                break;
        }

        if (markerStroke.Alpha < 1d)
        {
            graphics.RestoreState();
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

    private static void RenderAreaChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool stacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
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

            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills, 0.62d);
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(polygon);
            graphics.RestoreState();
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            for (int i = 1; i < upperPoints.Length; i++)
            {
                graphics.StrokeLine(upperPoints[i - 1].X, upperPoints[i - 1].Y, upperPoints[i].X, upperPoints[i].Y);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
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

    private static void RenderScatterChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries)
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
            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (fill.Alpha < 1d || stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            var points = new List<(double X, double Y)>(series[seriesIndex].Points.Count);
            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                double pointX = plotX + (point.X - minX) / xRange * plotWidth;
                double pointY = plotY + (point.Y - minY) / yRange * plotHeight;
                points.Add((pointX, pointY));
            }

            if (connectLines)
            {
                if (IsSmoothSeries(seriesIndex, smoothSeries))
                {
                    StrokeSmoothChartPath(graphics, points);
                }
                else
                {
                    StrokeStraightChartPath(graphics, points);
                }
            }

            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                double pointX = plotX + (point.X - minX) / xRange * plotWidth;
                double pointY = plotY + (point.Y - minY) / yRange * plotHeight;
                double radius = bubble ? 3d + Math.Sqrt(Math.Max(0d, point.Size) / maxBubbleSize) * 8d : 3d;
                if (bubble)
                {
                    graphics.FillEllipse(pointX - radius, pointY - radius, radius * 2d, radius * 2d);
                }
                else
                {
                    DrawChartMarker(graphics, pointX, pointY, ChartMarker(seriesIndex, markerStyles), fill.Color, stroke.Color);
                }

            }

            if (fill.Alpha < 1d || stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static void RenderRadarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
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

            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills, series.Count == 1 ? 0.40d : 0.18d);
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(points);
            graphics.RestoreState();
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            for (int i = 0; i < points.Length; i++)
            {
                (double X, double Y) a = points[i];
                (double X, double Y) b = points[(i + 1) % points.Length];
                graphics.StrokeLine(a.X, a.Y, b.X, b.Y);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static void RenderPieChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions)
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
            double midpointAngle = angle + sweep / 2d;
            double explosionOffset = pointExplosions.TryGetValue(i, out double explosion) ? radius * explosion : 0d;
            double sliceCenterX = centerX + Math.Cos(midpointAngle) * explosionOffset;
            double sliceCenterY = centerY + Math.Sin(midpointAngle) * explosionOffset;
            int segments = Math.Max(2, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 18d)));
            var points = new (double X, double Y)[segments + 2];
            points[0] = (sliceCenterX, sliceCenterY);
            for (int segment = 0; segment <= segments; segment++)
            {
                double segmentAngle = angle + sweep * segment / segments;
                points[segment + 1] = (
                    sliceCenterX + Math.Cos(segmentAngle) * radius,
                    sliceCenterY + Math.Sin(segmentAngle) * radius);
            }

            ChartSeriesFill fill = pointFills.TryGetValue(i, out ChartSeriesFill explicitFill)
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(i), 1d);
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(points);
            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }

            if (pointStrokes.TryGetValue(i, out ChartSeriesStroke stroke))
            {
                if (stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }

                SetChartStroke(graphics, stroke);
                graphics.StrokePolygon(points);
                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }

            angle += sweep;
        }
    }

    private static void RenderDoughnutChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize)
    {
        RenderPieChartFallback(graphics, document, bounds, values, pointFills, pointStrokes, pointExplosions);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double innerRadius = radius * holeSize;
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

    private static ChartSeriesStroke ChartAxisDefaultStroke { get; } = new(new RgbColor(90, 90, 90), 1d, 0.75d);

    private readonly record struct ChartAxesStyle(ChartSeriesStroke? ValueAxis, ChartSeriesStroke? CategoryAxis);

    private readonly record struct ChartPlotBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartShapeStyle(ChartSeriesFill? Fill, ChartSeriesStroke? Stroke)
    {
        public static ChartShapeStyle Empty { get; } = new(null, null);
    }

    private readonly record struct ChartMarkerStyle(string Symbol, double Size, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke)
    {
        public static ChartMarkerStyle Default { get; } = new("circle", 4d, null, null);
    }
}
