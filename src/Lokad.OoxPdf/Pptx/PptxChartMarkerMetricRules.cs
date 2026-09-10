using System.Globalization;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxChartMarkerMetricRules
{
    private static readonly string[] AutoLineChartMarkerSymbols = ["diamond", "square", "triangle", "x", "star", "circle"];

    public const double DefaultChartMarkerSize = 4d;
    public const double AutoLineChartMarkerSize = 7d;
    public const double StyledLineChartMarkerSize = 9d;
    // Office default scatter markers step with per-series point density. Series with
    // fewer than DenseScatterMarkerMinimumPointCount points draw ~9.9pt markers (9.84
    // diamond diagonals and 9.96 square sides across the cached 11-reference probe set;
    // 7sqrt(2) = 9.8995 sits midrange, mechanism unclaimed), while denser series draw
    // the 7pt auto size. The cut sits exactly between the 5-point (sparse) and 6-point
    // (dense) probes; symbol order, style, lines, and series count were all killed as
    // discriminators by dedicated probes.
    public const double SparseScatterDefaultMarkerSize = 9.9d;
    // Office default marker outline is 0.75pt at every measured size (7pt dense and
    // 9.9pt sparse diamonds alike); the old size-scaled fallback drew 1.1-1.6pt rims.
    public const double DefaultMarkerOutlineWidth = 0.75d;
    public const int DenseScatterMarkerMinimumPointCount = 6;

    public static string ResolveDefaultSymbol(PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        if (plotKind == PptxSceneChartPlotKind.Line)
        {
            return chartMarkersEnabled ? AutoLineChartMarkerSymbols[seriesIndex % AutoLineChartMarkerSymbols.Length] : "none";
        }

        // Scatter shares the line auto-symbol order (cached Office refs show diamond,
        // square, triangle, x across series 0-3) unconditionally: scatter files without a
        // plot-level marker switch still render markers on both sides, so gating on
        // chartMarkersEnabled would newly suppress them. Bubble keeps the circle below
        // since its markers render as data-sized ellipses with the symbol unused.
        if (plotKind == PptxSceneChartPlotKind.Scatter)
        {
            return AutoLineChartMarkerSymbols[seriesIndex % AutoLineChartMarkerSymbols.Length];
        }

        return "circle";
    }

    public static double ResolveSize(string? sizeValue, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, bool markerDefined, bool hasShapeProperties, int scatterPointCount)
    {
        if (sizeValue is not null &&
            double.TryParse(sizeValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return Math.Clamp(parsed, 2d, 30d);
        }

        if (plotKind == PptxSceneChartPlotKind.Scatter && !markerDefined)
        {
            return scatterPointCount < DenseScatterMarkerMinimumPointCount
                ? SparseScatterDefaultMarkerSize
                : AutoLineChartMarkerSize;
        }

        if (plotKind != PptxSceneChartPlotKind.Line || !chartMarkersEnabled)
        {
            return DefaultChartMarkerSize;
        }

        if (!markerDefined)
        {
            return AutoLineChartMarkerSize;
        }

        if (hasShapeProperties)
        {
            return StyledLineChartMarkerSize;
        }

        return DefaultChartMarkerSize;
    }
}
