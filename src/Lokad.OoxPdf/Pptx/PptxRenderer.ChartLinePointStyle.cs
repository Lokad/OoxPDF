using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartLineAppearance(
        ChartSeriesStroke Stroke, ChartMarkerStyle Marker, RgbColor MarkerColor, RgbColor MarkerOutlineColor, double? MarkerOutlineWidth = null);

    // Resolve styles by source point index, shared by paths, markers and keys.
    // The category palette is independent of an authored series line color.
    private static Func<int, ChartLineAppearance>? ReadLinePointAppearance(
        PptxSceneChart chart, PptxSceneChartPlot? plot, XElement xml,
        PptxTheme theme, PptxColorMap colorMap, ChartColorPalette palette,
        IReadOnlyList<ChartIndexedNumberVector> vectors, ChartLinePlotOptions options,
        IReadOnlyList<PptxSceneChartSeries> sources, IReadOnlyList<ChartLineAppearance> seriesAppearance, int? styleId, double? themeWidth, CancellationToken cancellationToken)
    {
        // Office varies by point only for a single-series chart. Composed
        // charts retain their series owners.
        if (chart.Plots.Count != 1 || vectors.Count != 1 || !options.VaryColors.Value)
            return null;
        int count = vectors[0].DensePointCount();
        IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(plot, xml, theme, colorMap)[0];
        var sourceLine = sources[0].Line;
        var seriesStyle = seriesAppearance[0];
        ChartMarkerStyle seriesMarker = seriesStyle.Marker;
        bool rawStroke = palette.SingleAccent is not null || IsGalleryLineChartStyle(styleId);
        return index =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            RgbColor color = ChartPalette(palette, theme, colorMap, index, count);
            if (palette.SingleAccent is null) color = ResolveAutomaticCategoryColor(color, index, count);
            RgbColor outline = themeWidth is not null ? ResolveThemeLineColor(theme, colorMap, color) :
                rawStroke ? color : ApplyUnstyledLineStrokeTint(color);
            ChartSeriesStroke stroke = pointStrokes.TryGetValue(index, out var pointStroke) ? pointStroke :
                seriesStyle.Stroke with { Color = sourceLine.HasLine ? seriesStyle.Stroke.Color : outline };
            ChartMarkerStyle marker = seriesMarker;
            if (!marker.IsDefined || marker.Symbol == "auto")
                marker = marker with { SymbolKind = PptxChartMarkerMetricRules.ResolveAutomaticLinePointSymbol(index) };
            return new(stroke, marker, color, outline, seriesStyle.MarkerOutlineWidth);
        };
    }

    private static IReadOnlyList<ChartLegendEntry> BuildLinePointLegendEntries(
        PptxSceneChartPlot? plot, XElement xml, ChartWorkbookData? workbook, bool plotVisibleOnly,
        int count, int? styleId, Func<int, ChartLineAppearance> appearance)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = ReadSharedCategoryLabels(plot, xml, workbook, plotVisibleOnly);
        var entries = new List<ChartLegendEntry>(count);
        for (int index = 0; index < count; index++)
        {
            ChartLineAppearance point = appearance(index);
            ChartMarkerStyle marker = ResolveStrokeLegendMarker(point.Marker, index, styleId) ?? point.Marker;
            marker = marker with { Stroke = marker.Stroke ?? new ChartSeriesStroke(point.MarkerOutlineColor, 1d,
                point.MarkerOutlineWidth ?? (IsGalleryLineChartStyle(styleId) ? PptxChartMarkerMetricRules.StyleLineMarkerOutlineWidth : PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth)) with { Join = 1 } };
            entries.Add(new ChartLegendEntry(index < labels.Count ? labels[index]?.Text ?? string.Empty : string.Empty,
                null, point.Stroke, marker, null, false, styleId, point.MarkerColor));
        }
        return entries;
    }
}
