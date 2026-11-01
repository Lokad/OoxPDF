using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartLinePointAppearance(
        ChartSeriesStroke Stroke, ChartMarkerStyle Marker, RgbColor MarkerColor, RgbColor MarkerOutlineColor);

    // Resolve styles by source point index, shared by paths, markers and keys.
    // The category palette is independent of an authored series line color.
    private static Func<int, ChartLinePointAppearance>? ReadLinePointAppearance(
        PptxSceneChart chart, PptxSceneChartPlot? plot, XElement xml,
        PptxTheme theme, PptxColorMap colorMap, ChartColorPalette palette,
        IReadOnlyList<ChartIndexedNumberVector> vectors, ChartLinePlotOptions options,
        IReadOnlyList<ChartSeriesStroke?> strokes, IReadOnlyList<ChartMarkerStyle> markers, int? styleId)
    {
        // Office varies by point only for a single-series chart. Composed
        // charts retain their series owners.
        if (chart.Plots.Count != 1 || vectors.Count != 1 || !options.VaryColors.Value)
            return null;
        int count = vectors[0].DensePointCount();
        IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(plot, xml, theme, colorMap)[0];
        ChartSeriesStroke? authoredStroke = strokes.Count > 0 ? strokes[0] : null;
        ChartMarkerStyle seriesMarker = ChartMarker(0, markers);
        bool rawStroke = palette.SingleAccent is not null || styleId is 18 or 118 or 26 or 126;
        double width = ResolveStyleLineSeriesWidth(styleId, smoothed: false);
        XElement? lineColor = theme.TryGetLineStyle(1, out var themeLine)
            ? themeLine.Element(DrawingNamespace + "solidFill")?.Elements().FirstOrDefault(PptxColorResolver.IsDrawingColorElement)
            : null;
        return index =>
        {
            RgbColor color = ChartPalette(palette, theme, colorMap, index, count);
            if (palette.SingleAccent is null) color = ResolveAutomaticCategoryColor(color, index, count);
            RgbColor outline = rawStroke ? color :
                lineColor?.Name == DrawingNamespace + "schemeClr" && (string?)lineColor.Attribute("val") == "phClr"
                    ? PptxColorResolver.ApplyColorTransforms(lineColor, color)
                    : lineColor is not null && PptxColorResolver.TryReadSolidColorWithAlpha(lineColor, theme, colorMap, out var resolved, out _)
                        ? resolved : ApplyUnstyledLineStrokeTint(color);
            ChartSeriesStroke stroke = pointStrokes.TryGetValue(index, out var pointStroke) ? pointStroke :
                authoredStroke ?? new ChartSeriesStroke(outline, 1d, width) with { Cap = 1, Join = 1 };
            ChartMarkerStyle marker = seriesMarker;
            if (!marker.IsDefined || marker.Symbol == "auto")
                marker = marker with { SymbolKind = PptxChartMarkerMetricRules.ResolveAutomaticLinePointSymbol(index) };
            return new(stroke, marker, color, outline);
        };
    }

    private static IReadOnlyList<ChartLegendEntry> BuildLinePointLegendEntries(
        PptxSceneChartPlot? plot, XElement xml, ChartWorkbookData? workbook, bool plotVisibleOnly,
        int count, int? styleId, Func<int, ChartLinePointAppearance> appearance)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = ReadSharedCategoryLabels(plot, xml, workbook, plotVisibleOnly);
        var entries = new List<ChartLegendEntry>(count);
        for (int index = 0; index < count; index++)
        {
            ChartLinePointAppearance point = appearance(index);
            ChartMarkerStyle marker = ResolveStrokeLegendMarker(point.Marker, index, styleId) ?? point.Marker;
            marker = marker with { Stroke = marker.Stroke ?? new ChartSeriesStroke(point.MarkerOutlineColor, 1d,
                IsGalleryLineChartStyle(styleId) ? PptxChartMarkerMetricRules.StyleLineMarkerOutlineWidth : PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth) with { Join = 1 } };
            entries.Add(new ChartLegendEntry(index < labels.Count ? labels[index]?.Text ?? string.Empty : string.Empty,
                null, point.Stroke, marker, null, false, styleId, point.MarkerColor));
        }
        return entries;
    }
}
