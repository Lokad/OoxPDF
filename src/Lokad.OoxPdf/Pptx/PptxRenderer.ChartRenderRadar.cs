using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderRadarChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartRadarLayout layout, IReadOnlyList<ChartRadarSeries> series, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents extents, ChartAxisUnits axisUnits, int? chartStyleId)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        int pointCount = layout.PointCount;

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double?> values = series[seriesIndex].Values;
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count && values[i] is { } pointValue ? Math.Max(0d, pointValue) : 0d;
                double pointRadius = GetChartValuePlotRatio(extents, value, false) * geometry.Radius;
                double angle = GetRadarPointAngle(i, pointCount);
                points[i] = (geometry.CenterX + Math.Cos(angle) * pointRadius, geometry.CenterY + Math.Sin(angle) * pointRadius);
            }

            if (layout.IsFilled)
            {
                // RV04: effective-style-18 filled-radar series paint a vertical theme-relative
                // gradient sampled from the Office reference; other styles and explicit
                // fills keep the flat polygon.
                ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
                if ((seriesIndex < seriesFills.Count ? seriesFills[seriesIndex] : null) is null &&
                    TryReadStyle18SeriesGradient(chartStyleId, fill.Color, out RgbColor gradientTop, out RgbColor gradientBottom))
                {
                    PaintRadarSeriesGradient(graphics, points, gradientTop, gradientBottom);
                }
                else
                {
                    graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
                    graphics.FillPolygon(points);
                }
                graphics.RestoreState();
            }
        }

        // RV04: series fills paint before grid spokes so spokes stay visible over filled polygons.

        SetChartStroke(graphics, RadarGridlineDefaultStroke);
        bool hasGridPath = false;
        foreach (double tickValue in GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount))
        {
            double tickRatio = GetChartValuePlotRatio(extents, tickValue, false);
            if (tickRatio <= PptxChartMetricRules.AxisValueEpsilon)
            {
                continue;
            }

            AppendRadarPolygonEdgeSegments(geometry.Radius * tickRatio);
            hasGridPath = true;
        }
        if (hasGridPath)
        {
            graphics.StrokeCurrentPath();
        }

        AppendRadarRadialSegments(repeatFirstSpoke: true);
        graphics.StrokeCurrentPath();

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double?> values = series[seriesIndex].Values;
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count && values[i] is { } pointValue ? Math.Max(0d, pointValue) : 0d;
                double pointRadius = GetChartValuePlotRatio(extents, value, false) * geometry.Radius;
                double angle = GetRadarPointAngle(i, pointCount);
                points[i] = (geometry.CenterX + Math.Cos(angle) * pointRadius, geometry.CenterY + Math.Sin(angle) * pointRadius);
            }

            // Office omits the polygon outline on filled radars without an explicit line;
            // marker radars always outline (3.75pt default measured on the ladder reference).
            bool hasExplicitStroke = seriesIndex < seriesStrokes.Count && seriesStrokes[seriesIndex] is not null;
            if (!layout.IsFilled || hasExplicitStroke)
            {
                ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, PptxChartMetricRules.RadarSeriesOutlineWidth);
                if (!hasExplicitStroke)
                {
                    stroke = stroke with { Color = ApplyUnstyledLineStrokeTint(stroke.Color) };
                }
                if (stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }

                SetChartStroke(graphics, stroke);
                AppendClosedPolylinePath(points);
                graphics.StrokeCurrentPath();

                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }
        }

        void AppendClosedPolylinePath(IReadOnlyList<(double X, double Y)> points)
        {
            if (points.Count == 0)
            {
                return;
            }

            graphics.MoveTo(points[0].X, points[0].Y);
            for (int i = 1; i < points.Count; i++)
            {
                graphics.LineTo(points[i].X, points[i].Y);
            }

            graphics.LineTo(points[0].X, points[0].Y);
        }

        void AppendRadarPolygonEdgeSegments(double radius)
        {
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double angle = GetRadarPointAngle(i, pointCount);
                points[i] = (geometry.CenterX + Math.Cos(angle) * radius, geometry.CenterY + Math.Sin(angle) * radius);
            }

            for (int i = 0; i < points.Length; i++)
            {
                (double X, double Y) start = points[i];
                (double X, double Y) end = points[(i + 1) % points.Length];
                graphics.MoveTo(start.X, start.Y);
                graphics.LineTo(end.X, end.Y);
            }
        }

        void AppendRadarRadialSegments(bool repeatFirstSpoke)
        {
            int segmentCount = repeatFirstSpoke ? pointCount + 1 : pointCount;
            for (int i = 0; i < segmentCount; i++)
            {
                double angle = GetRadarPointAngle(i % pointCount, pointCount);
                graphics.MoveTo(geometry.CenterX, geometry.CenterY);
                graphics.LineTo(
                    geometry.CenterX + Math.Cos(angle) * geometry.Radius,
                    geometry.CenterY + Math.Sin(angle) * geometry.Radius);
            }
        }
    }

    // Reads the effective chart style id from the markup-compatibility-resolved tree:
    // a spliced c14 style when that namespace is understood, else the transitional
    // style. In this corpus 18 always arrives as the c14:118 fallback; a native-2007
    // style-18 reference would confirm or split the recipe below.
    private static int? ReadChartStyleId(XDocument chartXml)
    {
        XElement? root = chartXml.Root;
        XElement? styled = root?.Element(ChartStyle2010Namespace + "style")
            ?? root?.Element(ChartNamespace + "style");
        return styled is not null && TryReadStyleValue(styled, out int styleId)
            ? styleId
            : null;

        static bool TryReadStyleValue(XElement element, out int value)
        {
            value = 0;
            return element.Attribute("val") is { } attribute &&
                int.TryParse(attribute.Value, out value);
        }
    }

    // Effective-style-18 series recipe calibrated from cached Office references: a vertical
    // light-to-dark gradient expressed as linear transforms of the series base color
    // so themed documents keep their hue. The same recipe fits both ladder series.
    private static bool TryReadStyle18SeriesGradient(int? chartStyleId, RgbColor baseColor, out RgbColor top, out RgbColor bottom)
    {
        top = default;
        bottom = default;
        if (chartStyleId != 18)
        {
            return false;
        }

        top = ApplyChartStyleGradientStop(baseColor, 0.88d, 0.315d);
        bottom = ApplyChartStyleGradientStop(baseColor, 1.2832d, -0.1427d);
        return true;
    }

    private static RgbColor ApplyChartStyleGradientStop(RgbColor baseColor, double multiplier, double offset)
    {
        return new RgbColor(
            GradientStopByte(baseColor.Red * multiplier + 255d * offset),
            GradientStopByte(baseColor.Green * multiplier + 255d * offset),
            GradientStopByte(baseColor.Blue * multiplier + 255d * offset));

        static byte GradientStopByte(double value) => (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static void PaintRadarSeriesGradient(PdfGraphicsBuilder graphics, (double X, double Y)[] points, RgbColor top, RgbColor bottom)
    {
        if (points.Length == 0)
        {
            return;
        }

        double minY = points[0].Y;
        double maxY = points[0].Y;
        double centerX = 0d;
        foreach (var (x, y) in points)
        {
            minY = Math.Min(minY, y);
            maxY = Math.Max(maxY, y);
            centerX += x;
        }

        if (maxY - minY <= 0.001d)
        {
            graphics.SetFillRgb(top.Red, top.Green, top.Blue);
            graphics.FillPolygon(points);
            return;
        }

        centerX /= points.Length;
        graphics.ClipPolygon(points);
        graphics.PaintAxialShading(centerX, maxY, centerX, minY, top.Red, top.Green, top.Blue, bottom.Red, bottom.Green, bottom.Blue);
    }

    private static ChartValueExtents GetAreaChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, bool stacked, bool percentStacked)
    {
        if (percentStacked)
        {
            return new ChartValueExtents(0d, 1d);
        }

        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
        int pointCount = Math.Max(1, denseSeries.Max(values => values.Count));
        (double minValue, double maxValue) = stacked
            ? GetStackedPointValueExtents(denseSeries, pointCount, percentStacked)
            : GetClusteredPointValueExtents(denseSeries);
        return new ChartValueExtents(minValue, maxValue);
    }

    private static double GetRadarPointAngle(int index, int pointCount)
    {
        return Math.PI / 2d - index * Math.PI * 2d / pointCount;
    }

    private static void RenderRadarCategoryLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartRadarLayout layout,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartAxis? sceneAxis,
        XElement? categoryAxis,
        IReadOnlyList<ChartIndexedTextPoint?> labels,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (labels.Count == 0)
        {
            return;
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        ChartPlotBox plotBox = layout.PlotBox;
        int pointCount = Math.Max(labels.Count, layout.PointCount);
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            string? label = labels[i]?.Text;
            if (string.IsNullOrWhiteSpace(label))
            {
                continue;
            }

            ChartRadarLabelFrame frame = ResolveRadarCategoryLabelFrame(layout, label, style, textMeasurer, i, pointCount);
            runs.Add(CreateChartLabelRun(label, frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment, kerningEnabled: false));
        }

        RenderChartTextRuns(runs, graphics, chartFonts, "RCA", fontResolver, diagnosticSink);
    }

    private static ChartRadarLabelFrame ResolveRadarCategoryLabelFrame(ChartRadarLayout layout, string label, ChartTextStyle style, ChartTextMeasurer textMeasurer, int index, int pointCount)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        ChartRadarLabelRules labelRules = layout.LabelRules;
        // Gaps key off the web-side length (exact inverse of the web construction);
        // manual-plot webs reuse their legacy radius here (Office manual layout open).
        double webSide = 2d * (geometry.Radius + PptxChartMetricRules.RadarWebRadiusPenHalf);
        double verticalGap = webSide * labelRules.CategoryVerticalGapSideFactor + fontSize * labelRules.CategoryVerticalGapFontFactor;
        double horizontalGap = webSide * labelRules.CategoryHorizontalGapSideFactor;
        double angle = GetRadarPointAngle(index, pointCount);
        double cosine = Math.Cos(angle);
        double anchorX = geometry.CenterX + cosine * (geometry.Radius + horizontalGap);
        double anchorY = geometry.CenterY + Math.Sin(angle) * (geometry.Radius + verticalGap);
        double width = Math.Max(fontSize * 2d, textMeasurer.Measure(label, style) + fontSize);
        TextAlignment alignment = cosine > 0.25d
            ? TextAlignment.Left
            : cosine < -0.25d
                ? TextAlignment.Right
                : TextAlignment.Center;
        double x = alignment switch
        {
            TextAlignment.Left => anchorX,
            TextAlignment.Right => anchorX - width,
            _ => anchorX - width / 2d
        };
        double y = ResolveRadarCategoryLabelBaselineY(anchorY, angle, height, labelRules);
        return new ChartRadarLabelFrame(x, y, width, height, alignment);
    }

    private static double ResolveRadarCategoryLabelBaselineY(double anchorY, double angle, double labelHeight, ChartRadarLabelRules labelRules)
    {
        double sine = Math.Sin(angle);
        double baselineFactor = labelRules.CategoryBaselineBaseFactor +
            labelRules.CategoryBaselineSineFactor * sine +
            labelRules.CategoryBaselineSineSquaredFactor * sine * sine;
        return anchorY + labelHeight * baselineFactor;
    }

    private static void RenderRadarValueAxisLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartRadarLayout layout,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        XElement? valueAxis,
        PptxSceneChartAxis? sceneAxis,
        ChartValueExtents extents,
        ChartAxisUnits axisUnits,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        ChartPlotBox plotBox = layout.PlotBox;
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        var runs = new List<TextRun>();
        foreach (double tickValue in GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount))
        {
            double ratio = GetChartValuePlotRatio(extents, tickValue, false);
            string label = FormatSceneOrXmlChartAxisLabel(tickValue, sceneAxis, valueAxis, null);
            ChartRadarLabelFrame frame = ResolveRadarValueAxisLabelFrame(layout, label, style, textMeasurer, ratio);
            runs.Add(CreateChartLabelRun(label, frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment, kerningEnabled: false));
        }

        RenderChartTextRuns(runs, graphics, chartFonts, "RVA", fontResolver, diagnosticSink);
    }

    private static ChartRadarLabelFrame ResolveRadarValueAxisLabelFrame(ChartRadarLayout layout, string label, ChartTextStyle style, ChartTextMeasurer textMeasurer, double ratio)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        ChartRadarLabelRules labelRules = layout.LabelRules;
        double width = Math.Max(
            fontSize * labelRules.ValueWidthFactor,
            textMeasurer.Measure(label, style) + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        double x = geometry.CenterX - width - fontSize * labelRules.ValueGapFactor;
        double y = ResolveRadarValueAxisLabelBaselineY(layout, ratio, height);
        return new ChartRadarLabelFrame(x, y, width, height, TextAlignment.Right);
    }

    private static double ResolveRadarValueAxisLabelBaselineY(ChartRadarLayout layout, double ratio, double labelHeight)
    {
        return layout.Geometry.CenterY + layout.Geometry.Radius * ratio - labelHeight * layout.LabelRules.ValueBaselineOffsetFactor;
    }
}
