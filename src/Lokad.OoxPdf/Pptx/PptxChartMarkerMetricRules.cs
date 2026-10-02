using System.Globalization;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxChartMarkerMetricRules
{
    private static readonly string[] AutoLineChartMarkerSymbols = ["diamond", "square", "triangle", "x", "star", "circle"];

    public const double DefaultChartMarkerSize = 4d;
    // Explicit markers without a size element draw at the nominal size-5 geometry (COM
    // decks: no-size dot/star/plus/x/circle/diamond/square/triangle markers all match
    // explicit size 5 digit-for-digit); bare auto markers and styled markers keep their
    // own defaults.
    public const double UnsetMarkerSize = 5d;
    // Marker rect extents land on a 0.12pt grid with ties down (COM dot/dash ladders:
    // all 20 measured rect dims fit; the ties at 37.5 and 12.5 go to 37 and 12).
    public const double MarkerExtentQuantum = 0.12d;
    public static double QuantizeMarkerExtent(double value)
    {
        decimal scaled = (decimal)value / 0.12m;
        decimal floored = decimal.Floor(scaled);
        if (scaled - floored > 0.5m)
        {
            floored += 1m;
        }

        return (double)(floored * 0.12m);
    }
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
    // Style-18 line gallery markers without marker markup draw at 12.96pt with 1pt
    // raw-base outlines (COM-built 1- and 3-series style-18 decks: 28 diamond/square/
    // triangle markers all measure 12.96, including on a doubled-height frame, so the
    // size is absolute, not plot-relative; symbol order follows the auto line order).
    public const double StyleLineMarkerSize = 12.96d;
    public const double StyleLineMarkerOutlineWidth = 1d;
    // Gallery legend keys draw markers at a fixed smaller size (COM-built style-18
    // and style-26 legend decks: diamonds/triangles 9.84, squares 9.96; shipped as a
    // uniform 9.9 since the per-shape spread is sub-pixel dust on single-deck evidence).
    public const double StyleLegendMarkerSize = 9.9d;
    // Smoothed default-style markers without marker markup draw flat at 9pt with raw
    // 1pt rims (COM-built default-style smooth deck: squares measure 9.0; diamond and
    // triangle paths measure 8.88, recorded as a watch item rather than a per-shape rule
    // on single-deck evidence).
    public const double SmoothLineMarkerSize = 9d;
    public const int DenseScatterMarkerMinimumPointCount = 6;

    public static PptxSceneChartMarkerSymbol ResolveForcedLineMarkerSymbol(int seriesIndex)
    {
        return PptxSceneBuilder.ParseChartMarkerSymbol(AutoLineChartMarkerSymbols[seriesIndex % AutoLineChartMarkerSymbols.Length]);
    }

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

    public static double ResolveSize(string? sizeValue, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, bool markerDefined, bool hasShapeProperties, int scatterPointCount, bool symbolDefined)
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

        // Explicit dot/star markers without a size element draw at the nominal size-5
        // geometry (COM decks: no-size dot rects and star diameters match explicit size 5
        // digit-for-digit), except styled (spPr) markers which keep the style-9 size; other
        // markers without an explicit symbol keep the size-4 fallback and scatter markers
        // keep density sizing.
        if (sizeValue is null && markerDefined && symbolDefined && !hasShapeProperties
            && plotKind == PptxSceneChartPlotKind.Line)
        {
            return UnsetMarkerSize;
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
