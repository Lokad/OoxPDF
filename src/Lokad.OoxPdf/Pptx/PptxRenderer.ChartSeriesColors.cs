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
    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, seriesIndex), defaultAlpha, null, null);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, categoryIndex), 1d, null, null)
            : ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d, null, null);
        }

        return ChartPointCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
    }

    private static bool HasExplicitChartPointFill(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int seriesIndex, int categoryIndex)
    {
        return seriesIndex < pointFills.Count && pointFills[seriesIndex].ContainsKey(categoryIndex);
    }

    private static ChartSeriesStroke? ResolveNegativeBarFallbackStroke(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, int seriesIndex, int categoryIndex, double value)
    {
        return value < 0d && !HasExplicitChartPointStroke()
            ? ChartNegativeBarDefaultStroke
            : null;

        bool HasExplicitChartPointStroke()
        {
            return seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        }
    }
}
