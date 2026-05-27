using System.Globalization;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxChartMarkerMetricRules
{
    public const double DefaultChartMarkerSize = 4d;
    public const double AutoLineChartMarkerSize = 7d;
    public const double StyledLineChartMarkerSize = 9d;

    public static double ResolveSize(string? sizeValue, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, bool markerDefined, bool hasShapeProperties)
    {
        if (sizeValue is not null &&
            double.TryParse(sizeValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return Math.Clamp(parsed, 2d, 30d);
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
