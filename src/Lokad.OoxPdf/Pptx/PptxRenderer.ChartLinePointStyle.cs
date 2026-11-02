using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartLineAppearance(
        ChartSeriesStroke Stroke, ChartMarkerStyle Marker, RgbColor MarkerColor, RgbColor MarkerOutlineColor, double? MarkerOutlineWidth = null);

    // Resolve styles by source point index, shared by paths, markers and keys.
    // The category palette is independent of an authored series line color.
    private static Func<int, int, ChartLineAppearance>? ReadLinePointAppearance(
        PptxTheme theme, PptxColorMap colorMap, ChartColorPalette palette,
        IReadOnlyList<ChartIndexedNumberVector> vectors, bool varyPointColors,
        IReadOnlyList<PptxSceneChartSeries> sources, IReadOnlyList<ChartLineAppearance> seriesAppearance, int? styleId, double? themeWidth, CancellationToken cancellationToken)
    {
        var pointMarkers = new Dictionary<int, PptxSceneChartMarker>?[sources.Count];
        bool hasPointAppearance = false;
        for (int seriesIndex = 0; seriesIndex < sources.Count; seriesIndex++)
        {
            int pointLimit = sources[seriesIndex].PointStyles.Count == 0 ? 0 : vectors[seriesIndex].DensePointCount();
            foreach (var point in sources[seriesIndex].PointStyles)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (point.Index >= pointLimit) continue;
                if (point.HasMarkerAppearance)
                    (pointMarkers[seriesIndex] ??= new())[point.Index] = point.Marker!;
                hasPointAppearance |= point.HasShapeProperties || point.HasMarkerAppearance;
            }
        }
        if (!varyPointColors && !hasPointAppearance) return null;
        int count = varyPointColors ? vectors[0].DensePointCount() : 0;
        IReadOnlyDictionary<int, ChartSeriesStroke>? pointStrokes = varyPointColors ? ReadSceneChartPointStrokes(sources[0]) : null;
        bool rawStroke = palette.SingleAccent is not null || IsGalleryLineChartStyle(styleId);
        return (seriesIndex, index) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            ChartLineAppearance result = seriesAppearance[seriesIndex];
            if (varyPointColors)
            {
                RgbColor color = ChartPalette(palette, theme, colorMap, index, count);
                if (palette.SingleAccent is null) color = ResolveAutomaticCategoryColor(color, index, count);
                RgbColor outline = themeWidth is not null ? ResolveThemeLineColor(theme, colorMap, color) :
                    rawStroke ? color : ApplyUnstyledLineStrokeTint(color);
                ChartSeriesStroke stroke = pointStrokes!.TryGetValue(index, out var pointStroke) ? pointStroke :
                    result.Stroke with { Color = sources[seriesIndex].Line.HasLine ? result.Stroke.Color : outline };
                ChartMarkerStyle marker = result.Marker;
                if (!marker.IsDefined || marker.Symbol == "auto")
                    marker = marker with { SymbolKind = PptxChartMarkerMetricRules.ResolveAutomaticLinePointSymbol(index) };
                result = new(stroke, marker, color, outline, result.MarkerOutlineWidth);
            }
            if (pointMarkers[seriesIndex] is { } markers && markers.TryGetValue(index, out var ownMarker))
                result = result with
                {
                    Marker = ToChartMarkerStyle(ownMarker),
                    MarkerOutlineWidth = ownMarker.Line.WidthSpecified ? ownMarker.Line.Width : result.MarkerOutlineWidth
                };
            return result;
        };
    }

    private static IReadOnlyList<ChartLegendEntry> BuildLinePointLegendEntries(
        PptxSceneChartPlot? plot, XElement xml, ChartWorkbookData? workbook, bool plotVisibleOnly,
        int count, int? styleId, Func<int, int, ChartLineAppearance> appearance)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = ReadSharedCategoryLabels(plot, xml, workbook, plotVisibleOnly);
        var entries = new List<ChartLegendEntry>(count);
        for (int index = 0; index < count; index++)
        {
            ChartLineAppearance point = appearance(0, index);
            ChartMarkerStyle marker = ResolveStrokeLegendMarker(point.Marker, index, styleId) ?? point.Marker;
            marker = marker with { Stroke = marker.Stroke ?? new ChartSeriesStroke(point.MarkerOutlineColor, 1d,
                point.MarkerOutlineWidth ?? (IsGalleryLineChartStyle(styleId) ? PptxChartMarkerMetricRules.StyleLineMarkerOutlineWidth : PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth)) with { Join = 1 } };
            entries.Add(new ChartLegendEntry(index < labels.Count ? labels[index]?.Text ?? string.Empty : string.Empty,
                null, point.Stroke, marker, null, false, styleId, point.MarkerColor));
        }
        return entries;
    }
}
