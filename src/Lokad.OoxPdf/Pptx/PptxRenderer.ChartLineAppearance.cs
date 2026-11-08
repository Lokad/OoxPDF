using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static double? ReadChartLineThemeWidth(PptxTheme theme) =>
        theme.TryGetLineStyle(1, out var line) &&
        long.TryParse((string?)line.Attribute("w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long width) && width >= 0
            ? OoxUnits.EmuToPoints(width) : null;

    private static double ResolveThemeLineWidth(double themeWidth, int? styleId) =>
        themeWidth * (styleId is 26 or 126 ? 7d : styleId is 18 or 118 ? 5d : 3d);

    private static RgbColor ResolveThemeLineColor(PptxTheme theme, PptxColorMap colorMap, RgbColor color)
    {
        XElement? source = theme.TryGetLineStyle(1, out var line)
            ? line.Element(DrawingNamespace + "solidFill")?.Elements().FirstOrDefault(PptxColorResolver.IsDrawingColorElement)
            : null;
        return source?.Name == DrawingNamespace + "schemeClr" && (string?)source.Attribute("val") == "phClr"
            ? PptxColorResolver.ApplyColorTransforms(source, color)
            : source is not null && PptxColorResolver.TryReadSolidColorWithAlpha(source.Parent, theme, colorMap, out var resolved, out _)
                ? resolved : color;
    }

    private static IReadOnlyList<ChartLineAppearance> ReadLineSeriesAppearance(
        IReadOnlyList<PptxSceneChartSeries> sources, ChartLinePlotOptions options,
        PptxTheme theme, PptxColorMap colorMap, ChartColorPalette palette, int? styleId,
        double? themeWidth, bool pointColors, CancellationToken cancellationToken)
    {
        var result = new ChartLineAppearance[sources.Count];
        bool gallery = IsGalleryLineChartStyle(styleId);
        bool plotSmoothed = options.SmoothSeries.Any(static option => option.Value);
        for (int i = 0; i < sources.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var source = sources[i];
            var line = source.Line;
            bool straight = i < options.SmoothSeries.Count && options.SmoothSeries[i].IsDefined && !options.SmoothSeries[i].Value;
            bool legacySmooth = themeWidth is null && plotSmoothed && !straight && !gallery;
            RgbColor baseColor = ChartPalette(palette, theme, colorMap, i);
            RgbColor outline = themeWidth is not null ? ResolveThemeLineColor(theme, colorMap, baseColor) :
                palette.SingleAccent is not null || gallery || legacySmooth ? baseColor : ApplyUnstyledLineStrokeTint(baseColor);
            double inheritedWidth = themeWidth is { } basis ? ResolveThemeLineWidth(basis, styleId) :
                line.HasLine ? ResolveLineKindInheritedStrokeWidth(styleId) : ResolveStyleLineSeriesWidth(styleId, plotSmoothed && !straight);
            ChartSeriesStroke stroke = ToChartSeriesStroke(line, inheritedWidth)
                ?? new ChartSeriesStroke(outline, line.NoFill ? 0d : 1d, line.WidthSpecified ? line.Width : inheritedWidth,
                    line.DashPattern, line.Cap ?? 1, line.Join ?? 1, line.Compound);
            ChartMarkerStyle marker = ToChartMarkerStyle(source.Marker);
            if (themeWidth is not null && !marker.IsDefined)
                marker = marker with { Size = PptxChartMarkerMetricRules.QuantizeMarkerExtent(Math.Clamp(Math.Floor(2d * stroke.Width + 3d), 2d, 72d)) };
            else if (!marker.IsDefined && (styleId is 18 or 118 || legacySmooth && !pointColors))
                marker = marker with
                {
                    SymbolKind = marker.SymbolKind == PptxSceneChartMarkerSymbol.None
                        ? PptxChartMarkerMetricRules.ResolveForcedLineMarkerSymbol(i) : marker.SymbolKind,
                    Size = styleId is 18 or 118 ? PptxChartMarkerMetricRules.StyleLineMarkerSize : PptxChartMarkerMetricRules.SmoothLineMarkerSize
                };
            RgbColor markerColor = themeWidth is null && line.HasLine ? stroke.Color : baseColor;
            RgbColor markerOutline = themeWidth is not null ? ResolveThemeLineColor(theme, colorMap, baseColor) :
                line.HasLine || palette.SingleAccent is not null || styleId is 18 or 118 || legacySmooth && !pointColors
                    ? markerColor : ApplyUnstyledLineStrokeTint(markerColor);
            double? markerWidth = themeWidth ?? (styleId is 18 or 118 || legacySmooth && !pointColors
                ? PptxChartMarkerMetricRules.StyleLineMarkerOutlineWidth : null);
            result[i] = new(stroke, marker, markerColor, markerOutline, markerWidth);
        }
        return result;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildLineSeriesLegendEntries(
        PptxSceneChartPlot plot, XElement chartElement, ChartWorkbookData? workbook,
        IReadOnlyList<ChartLineAppearance> appearance, IReadOnlyList<PptxSceneChartSeries> sources, double? themeWidth, bool reverseOrder, int? styleId)
    {
        var names = ReadSharedChartSeriesNames(plot, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            var paint = appearance[i];
            ChartSeriesStroke stroke = paint.Stroke;
            if (themeWidth is null && !IsGalleryLineChartStyle(styleId) && !sources[i].Line.HasLine && !sources[i].Line.WidthSpecified)
                stroke = stroke with { Color = ApplyUnstyledLineStrokeTint(paint.MarkerColor), Width = ChartLineDefaultStrokeWidth };
            var marker = themeWidth is null ? ResolveStrokeLegendMarker(paint.Marker, i, styleId)!.Value : paint.Marker;
            marker = marker with { Stroke = marker.Stroke ?? new ChartSeriesStroke(paint.MarkerOutlineColor, 1d,
                paint.MarkerOutlineWidth ?? (IsGalleryLineChartStyle(styleId) ? PptxChartMarkerMetricRules.StyleLineMarkerOutlineWidth : PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth)) with { Join = 1 } };
            entries.Add(new(names[i].ActiveName, null, stroke, marker, names[i], paint.Stroke.Alpha <= .001d, styleId, paint.MarkerColor));
        }
        if (reverseOrder) entries.Reverse();
        return entries;
    }
}
