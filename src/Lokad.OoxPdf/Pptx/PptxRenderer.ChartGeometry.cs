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
    private static IReadOnlyList<ChartIndexedNumberVector> ReadChartSeriesVectors(XElement chartElement, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        var series = new List<ChartIndexedNumberVector>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            ChartIndexedNumberVector values = ReadChartNumberVector(element.Element(ChartNamespace + "val"), workbook, plotVisibleOnly);
            if (values.Points.Count > 0 || values.PointCount is not null || values.DensePoints().Count != 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static ChartIndexedNumberVector ReadChartNumberVector(XElement? container, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        if (container is null)
        {
            return default;
        }

        XElement? cache = container
            .Descendants(ChartNamespace + "numCache")
            .Concat(container.Descendants(ChartNamespace + "numLit"))
            .FirstOrDefault();
        PptxSceneChartDataSource source = PptxSceneBuilder.ReadChartDataSource(container, "numRef");
        if (cache is null)
        {
            return new ChartIndexedNumberVector(
                [],
                null,
                source.Formula,
                null,
                source,
                ReadWorkbookNumberPoints(workbook, source),
                plotVisibleOnly);
        }

        ChartIndexedNumberPoint[] points = PptxSceneBuilder
            .ReadChartNumberPoints(cache.Elements(ChartNamespace + "pt"), requireNonNegativeIndex: true)
            .Select(ToChartIndexedNumberPoint)
            .ToArray();
        return new ChartIndexedNumberVector(
            points,
            PptxSceneBuilder.ReadChartCachePointCount(cache).Value ?? InferPointCount(points),
            source.Formula,
            (string?)cache.Element(ChartNamespace + "formatCode"),
            source,
            ReadWorkbookNumberPoints(workbook, source),
            plotVisibleOnly);
    }

    private static void HydrateChartReferenceCaches(ChartWorkbookData workbook, XDocument chartXml)
    {
        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "numRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookNumericValue[] values = workbook.ReadNumericRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "numCache",
                new XElement(ChartNamespace + "formatCode", ReadChartReferenceNumberCacheFormatCode(reference)),
                BuildChartCachePointsFromWorkbookValues(values)));
        }

        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "strRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookTextValue[] values = workbook.ReadTextRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "strCache",
                BuildChartCachePointsFromWorkbookValues(values)));
        }

        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "multiLvlStrRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookTextValue[] values = workbook.ReadTextRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "multiLvlStrCache",
                new XElement(ChartNamespace + "ptCount", new XAttribute("val", (values.Max(value => value.Cell.Index) + 1).ToString(CultureInfo.InvariantCulture))),
                new XElement(ChartNamespace + "lvl", values.Select(value => BuildChartCachePoint(value.Cell.Index, value.Text)))));
        }
    }

    private static void InsertChartReferenceCache(XElement reference, XElement cache)
    {
        reference.Elements(cache.Name).Remove();
        XElement? formula = reference.Element(ChartNamespace + "f");
        if (formula is not null)
        {
            formula.AddAfterSelf(cache);
            return;
        }

        reference.AddFirst(cache);
    }

    private static string ReadChartReferenceNumberCacheFormatCode(XElement reference)
    {
        return (string?)reference
            .Element(ChartNamespace + "numCache")
            ?.Element(ChartNamespace + "formatCode") ?? "General";
    }

    private static object[] BuildChartCachePointsFromWorkbookValues(IReadOnlyList<ChartWorkbookNumericValue> values)
    {
        var elements = new object[values.Count + 1];
        int pointCount = values.Count == 0 ? 0 : values.Max(value => value.Cell.Index) + 1;
        elements[0] = new XElement(ChartNamespace + "ptCount", new XAttribute("val", pointCount.ToString(CultureInfo.InvariantCulture)));
        for (int index = 0; index < values.Count; index++)
        {
            ChartWorkbookNumericValue value = values[index];
            elements[index + 1] = BuildChartCachePoint(value.Cell.Index, value.Cell.Text);
        }

        return elements;
    }

    private static object[] BuildChartCachePointsFromWorkbookValues(IReadOnlyList<ChartWorkbookTextValue> values)
    {
        var elements = new object[values.Count + 1];
        int pointCount = values.Count == 0 ? 0 : values.Max(value => value.Cell.Index) + 1;
        elements[0] = new XElement(ChartNamespace + "ptCount", new XAttribute("val", pointCount.ToString(CultureInfo.InvariantCulture)));
        for (int index = 0; index < values.Count; index++)
        {
            ChartWorkbookTextValue value = values[index];
            elements[index + 1] = BuildChartCachePoint(value.Cell.Index, value.Text);
        }

        return elements;
    }

    private static XElement BuildChartCachePoint(int index, string value)
    {
        return new XElement(
            ChartNamespace + "pt",
            new XAttribute("idx", index.ToString(CultureInfo.InvariantCulture)),
            new XElement(ChartNamespace + "v", value));
    }

    private static ChartFrameBox GetChartFrameBox(PptxDocument document, ShapeBounds bounds)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        return new ChartFrameBox(x, y, width, height);
    }

    private static ChartPlotBox GetDefaultChartPlotBox(ChartFrameBox frame)
    {
        return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.DefaultCartesian);
    }

    private static ChartPlotBox GetChartPlotBoxPreset(ChartFrameBox frame, ChartPlotBoxPreset preset)
    {
        ChartPlotBoxRatios ratios = preset switch
        {
            ChartPlotBoxPreset.DefaultCartesian => new ChartPlotBoxRatios(
                PptxChartMetricRules.DefaultPlotBoxXRatio,
                PptxChartMetricRules.DefaultPlotBoxYRatio,
                PptxChartMetricRules.DefaultPlotBoxWidthRatio,
                PptxChartMetricRules.DefaultPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarDefault => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarDefaultPlotBoxXRatio,
                PptxChartMetricRules.BarDefaultPlotBoxYRatio,
                PptxChartMetricRules.BarDefaultPlotBoxWidthRatio,
                PptxChartMetricRules.BarDefaultPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarOverlayOnly => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarOverlayOnlyPlotBoxXRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxYRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxWidthRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarNoTitleBottomLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxXRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxYRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxWidthRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarTitleNoLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarTitleNoLegendPlotBoxXRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxYRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxWidthRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarTitleNoLegendInsideCrossing => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxXRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxYRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxWidthRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxHeightRatio),
            ChartPlotBoxPreset.HorizontalBarTitleNoLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxXRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxYRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxWidthRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.LineNoTitleRightLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxXRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxYRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxWidthRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.LineTitleRightLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.LineTitleRightLegendPlotBoxXRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxYRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxWidthRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxHeightRatio),
            _ => throw new ArgumentOutOfRangeException(nameof(preset), preset, null)
        };

        return GetChartPlotBox(frame, ratios);
    }

    private static ChartPlotBox GetChartPlotBox(ChartFrameBox frame, ChartPlotBoxRatios ratios)
    {
        return new ChartPlotBox(
            frame.X + frame.Width * ratios.Left,
            frame.Y + frame.Height * ratios.Top,
            frame.Width * ratios.Width,
            frame.Height * ratios.Height);
    }

    private static ChartPlotBox GetPolarChartPlotBox(PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        ChartPlotBox defaultPlotBox = new(frame.X, frame.Y, frame.Width, frame.Height);
        return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
            ? manualPlotLayout.PlotBox
            : defaultPlotBox;
    }

    private static ChartPolarGeometry GetPieOrDoughnutGeometry(ChartPolarKind kind, ChartPlotBox plotBox, double explosionReserve, ChartLegendLayout legend)
    {
        double radius = Math.Min(plotBox.Width, plotBox.Height) * GetPieOrDoughnutRadiusRatio(kind, legend);
        if (explosionReserve > 0d)
        {
            radius /= 1d + explosionReserve;
        }

        double centerXRatio = GetPieOrDoughnutCenterXRatio(kind, legend);
        double centerYRatio = GetPieOrDoughnutCenterYRatio(kind, legend);
        double centerXOffset = GetPieOrDoughnutCenterXOffset(kind, radius, explosionReserve, legend);
        return new ChartPolarGeometry(
            plotBox.X + plotBox.Width * centerXRatio + centerXOffset,
            plotBox.Y + plotBox.Height * centerYRatio,
            radius);
    }

    private static double GetPieOrDoughnutRadiusRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Pie => PptxChartMetricRules.PieRadiusRatio,
            ChartPolarKind.Doughnut when !hasLegend => PptxChartMetricRules.DoughnutNoLegendRadiusRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutNoLegendRadiusRatio,
            ChartPolarKind.Doughnut when IsHorizontalLegendPosition(legend.PositionKind) => PptxChartMetricRules.DoughnutHorizontalLegendRadiusRatio,
            ChartPolarKind.Doughnut => PptxChartMetricRules.PieRadiusRatio,
            _ => PptxChartMetricRules.PieRadiusRatio
        };
    }

    private static double GetPieOrDoughnutCenterXRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Pie => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && legend.PositionKind == PptxSceneChartLegendPosition.Right => PptxChartMetricRules.DoughnutRightLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutLeftLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && IsHorizontalLegendPosition(legend.PositionKind) => PptxChartMetricRules.DoughnutHorizontalLegendCenterXRatio,
            ChartPolarKind.Doughnut => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio,
            _ => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio
        };
    }

    private static double GetPieOrDoughnutCenterYRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Doughnut when !hasLegend => PptxChartMetricRules.DoughnutNoLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutNoLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Top => PptxChartMetricRules.DoughnutTopLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Bottom => PptxChartMetricRules.DoughnutBottomLegendCenterYRatio,
            _ => PptxChartMetricRules.PieCenterYRatio
        };
    }

    private static bool IsHorizontalLegendPosition(PptxSceneChartLegendPosition position)
    {
        return position is PptxSceneChartLegendPosition.Top or PptxSceneChartLegendPosition.Bottom;
    }

    private static double GetPieOrDoughnutCenterXOffset(ChartPolarKind kind, double radius, double explosionReserve, ChartLegendLayout legend)
    {
        if (kind == ChartPolarKind.Doughnut &&
            explosionReserve > 0d &&
            legend.Visible &&
            !legend.Overlay &&
            legend.PositionKind == PptxSceneChartLegendPosition.Right)
        {
            return radius * explosionReserve * PptxChartMetricRules.DoughnutExplosionCenterOffsetRatio;
        }

        return 0d;
    }

    private static ChartPolarLayout ResolvePieOrDoughnutLayout(ChartPolarKind kind, ChartPlotBox plotBox, IReadOnlyDictionary<int, double> pointExplosions, ChartLegendLayout legend)
    {
        double explosionReserve = pointExplosions.Count == 0 ? 0d : pointExplosions.Values.Max();
        bool hasLegend = legend.Visible && !legend.Overlay;
        return new ChartPolarLayout(
            kind,
            plotBox,
            GetPieOrDoughnutGeometry(kind, plotBox, explosionReserve, legend),
            explosionReserve,
            hasLegend);
    }

    private static ChartRadarLayout ResolveRadarLayout(ChartPlotBox plotBox, PptxSceneChartRadarStyle radarStyle, IReadOnlyList<ChartRadarSeries> series)
    {
        ChartRadarStyle style = radarStyle == PptxSceneChartRadarStyle.Filled
            ? ChartRadarStyle.Filled
            : ChartRadarStyle.Marker;
        return new ChartRadarLayout(
            plotBox,
            GetRadarChartGeometry(plotBox, style),
            style,
            Math.Max(3, series.Max(item => item.Points.Count)),
            ResolveRadarLabelRules(style));
    }

    private static ChartPolarGeometry GetRadarChartGeometry(ChartPlotBox plotBox, ChartRadarStyle style)
    {
        ChartRadarGeometryRule rule = ResolveRadarGeometryRule(style);
        return new ChartPolarGeometry(
            plotBox.X + plotBox.Width * rule.CenterXRatio,
            plotBox.Y + plotBox.Height * rule.CenterYRatio,
            Math.Min(plotBox.Width, plotBox.Height) * rule.RadiusRatio);
    }

    private static ChartRadarGeometryRule ResolveRadarGeometryRule(ChartRadarStyle style)
    {
        return style == ChartRadarStyle.Filled
            ? new ChartRadarGeometryRule(CenterXRatio: 0.5d, CenterYRatio: 0.4583333333333333d, RadiusRatio: 0.3825d)
            : new ChartRadarGeometryRule(CenterXRatio: 0.5d, CenterYRatio: 0.5d, RadiusRatio: 0.4226d);
    }

    private static ChartRadarLabelRules ResolveRadarLabelRules(ChartRadarStyle style)
    {
        double categoryHorizontalGapFactor = style == ChartRadarStyle.Filled ? 0.35d : 0.41d;
        return new ChartRadarLabelRules(
            CategoryVerticalGapFactor: 0.65d,
            CategoryHorizontalGapFactor: categoryHorizontalGapFactor,
            CategoryBaselineBaseFactor: -0.309d,
            CategoryBaselineSineFactor: -0.005d,
            CategoryBaselineSineSquaredFactor: 0.397d,
            ValueGapFactor: 1.01d,
            ValueBaselineOffsetFactor: 0.25d,
            ValueWidthFactor: 3.0d);
    }

    private static void RenderChartAreaStyle(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxTheme theme, PptxColorMap colorMap)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        RenderChartShapeStyle(graphics, frame.X, frame.Y, frame.Width, frame.Height, ReadSceneOrXmlChartAreaStyle(sceneChart, chartXml, theme, colorMap));
    }

    private static ChartShapeStyle ReadSceneOrXmlChartAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        return sceneChart is null
            ? ToChartShapeStyle(PptxSceneBuilder.ReadChartShapeStyle(chartXml.Root?.Element(ChartNamespace + "spPr"), theme, colorMap))
            : ToChartShapeStyle(sceneChart.ChartAreaStyle);
    }

    private static ChartShapeStyle ReadSceneOrXmlChartPlotAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        return sceneChart is null
            ? ReadChartPlotAreaStyle(chartXml, theme, colorMap)
            : ToChartShapeStyle(sceneChart.PlotAreaStyle);
    }

    private static ChartShapeStyle ReadChartPlotAreaStyle(XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "spPr");
        return ToChartShapeStyle(PptxSceneBuilder.ReadChartShapeStyle(shapeProperties, theme, colorMap));
    }

    private static bool TryReadSceneOrXmlManualPlotLayout(PptxSceneChart? sceneChart, XDocument chartXml, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        if (sceneChart is not null)
        {
            return TryBuildManualPlotLayout(sceneChart.PlotAreaLayout, frame, defaultPlotBox, out plotLayout);
        }

        return TryReadManualPlotLayout(chartXml, frame, defaultPlotBox, out plotLayout);
    }

    private static bool TryReadManualPlotLayout(XDocument chartXml, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        return TryBuildManualPlotLayout(PptxSceneBuilder.ReadChartPlotAreaManualLayout(chartXml), frame, defaultPlotBox, out plotLayout);
    }

    private static bool TryBuildManualPlotLayout(PptxSceneChartManualLayout layout, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        plotLayout = default;
        if (!TryBuildManualLayoutBox(layout, frame, new ChartLayoutBox(defaultPlotBox.X, defaultPlotBox.Y, defaultPlotBox.Width, defaultPlotBox.Height), out ChartLayoutBox layoutBox, true, false))
        {
            return false;
        }

        ChartPlotBox plotBox = new(layoutBox.X, layoutBox.Y, layoutBox.Width, layoutBox.Height);
        plotLayout = new ChartPlotLayout(layoutBox, plotBox, layout.LayoutTargetKind);
        return true;
    }

    private static bool TryBuildManualLayoutBox(PptxSceneChartManualLayout layout, ChartFrameBox frame, ChartLayoutBox defaultBox, out ChartLayoutBox box, bool clampToFrame, bool missingPositionModesAreFactor)
    {
        box = default;
        if (!layout.HasLayout)
        {
            return false;
        }

        ChartPlotBoxRatios defaults = GetLayoutBoxRatios(frame, defaultBox);
        double left = layout.X is { } x
            ? ClampManualLayoutRatio(ResolveManualLayoutStartRatio(x, layout.XModeKind, layout.XMode, defaultBox.X, frame.X, frame.Width, missingPositionModesAreFactor), clampToFrame)
            : defaults.Left;
        double top = layout.Y is { } y
            ? ClampManualLayoutRatio(ResolveManualLayoutStartRatio(y, layout.YModeKind, layout.YMode, frame.Y + frame.Height - defaultBox.Y - defaultBox.Height, 0d, frame.Height, missingPositionModesAreFactor), clampToFrame)
            : defaults.Top;
        double width = layout.Width is { } layoutWidth
            ? Math.Clamp(layoutWidth, 0.02d, 1d)
            : defaults.Width;
        double height = layout.Height is { } layoutHeight
            ? Math.Clamp(layoutHeight, 0.02d, 1d)
            : defaults.Height;
        double right = IsManualLayoutEdgeMode(layout.WidthModeKind)
            ? ClampManualLayoutEdgeRatio(layout.Width ?? defaults.Right, left, clampToFrame)
            : left + width;
        double bottom = IsManualLayoutEdgeMode(layout.HeightModeKind)
            ? ClampManualLayoutEdgeRatio(layout.Height ?? defaults.Bottom, top, clampToFrame)
            : top + height;
        double boxWidth = Math.Max(0d, right - left) * frame.Width;
        double boxHeight = Math.Max(0d, bottom - top) * frame.Height;
        double boxX = frame.X + left * frame.Width;
        double boxY = frame.Y + frame.Height - bottom * frame.Height;
        box = new ChartLayoutBox(boxX, boxY, boxWidth, boxHeight);
        return boxWidth > 0d && boxHeight > 0d;
    }

    private static double ClampManualLayoutRatio(double value, bool clampToFrame)
    {
        return clampToFrame ? Math.Clamp(value, 0d, 1d) : value;
    }

    private static double ClampManualLayoutEdgeRatio(double value, double minimum, bool clampToFrame)
    {
        return clampToFrame ? Math.Clamp(value, minimum, 1d) : Math.Max(minimum, value);
    }

    private static ChartPlotBoxRatios GetPlotBoxRatios(ChartFrameBox frame, ChartPlotBox plotBox)
    {
        return GetLayoutBoxRatios(frame, new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height));
    }

    private static ChartPlotBoxRatios GetLayoutBoxRatios(ChartFrameBox frame, ChartLayoutBox box)
    {
        if (frame.Width <= 0d || frame.Height <= 0d)
        {
            return new ChartPlotBoxRatios(0d, 0d, 1d, 1d);
        }

        double left = (box.X - frame.X) / frame.Width;
        double top = (frame.Y + frame.Height - box.Y - box.Height) / frame.Height;
        double width = box.Width / frame.Width;
        double height = box.Height / frame.Height;
        return new ChartPlotBoxRatios(left, top, width, height);
    }

    private static double ResolveManualLayoutStartRatio(double value, PptxSceneChartManualLayoutMode mode, string modeValue, double defaultStart, double frameStart, double frameLength, bool missingModeIsFactor)
    {
        if (IsManualLayoutFactorMode(mode, modeValue, missingModeIsFactor) && frameLength > 0d)
        {
            return (defaultStart - frameStart) / frameLength + value;
        }

        return value;
    }

    private static bool IsManualLayoutEdgeMode(PptxSceneChartManualLayoutMode mode)
    {
        return mode == PptxSceneChartManualLayoutMode.Edge;
    }

    private static bool IsManualLayoutFactorMode(PptxSceneChartManualLayoutMode mode)
    {
        return mode == PptxSceneChartManualLayoutMode.Factor;
    }

    private static bool IsManualLayoutFactorMode(PptxSceneChartManualLayoutMode mode, string modeValue, bool missingModeIsFactor)
    {
        return IsManualLayoutFactorMode(mode) || (missingModeIsFactor && string.IsNullOrEmpty(modeValue));
    }
}
