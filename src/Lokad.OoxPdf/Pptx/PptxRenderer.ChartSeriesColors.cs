using System.Globalization;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // A palette is an inherited color recipe, separate from authored series/point fills.
    private readonly record struct ChartColorPalette(
        IReadOnlyList<RgbColor>? Colors, RgbColor? SingleAccent, int SeriesCount);

    private static ChartColorPalette ResolveChartColorPalette(
        PptxSceneChart chart, IReadOnlyList<RgbColor>? colors, PptxTheme theme, PptxColorMap colorMap)
    {
        int seriesCount = chart.Plots.Sum(plot => plot.Series.Count);
        if (colors is { Count: > 0 }) { return new(colors, null, seriesCount); }

        int? preferred = OoxMarkupCompatibility.ReadPreferredChartGalleryStyle(chart.ChartXml?.Root);
        int style = preferred is int modern ? modern - 100 :
            int.TryParse(chart.StyleId, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int legacy) ? legacy : 0;
        int slot = (style - 1) % 8;
        if (style is >= 1 and <= 48 && slot >= 2 &&
            theme.TryResolveColor("accent" + (slot - 1).ToString(CultureInfo.InvariantCulture), colorMap, out RgbColor accent))
        {
            return new(null, accent, seriesCount);
        }

        return new(colors, null, seriesCount);
    }

    // Office's single-accent palette samples one continuous shade/tint span.
    // Integer percentages are shared by series and varying-point colors.
    private static RgbColor ResolveSingleAccentColor(RgbColor color, int index, int count)
    {
        int percentage = (int)(140L * (index + 1L) / (count + 1L));
        return percentage < 70
            ? LinearLightColor.ShadeTowardBlack(color, (30 + percentage) / 100d)
            : LinearLightColor.TintTowardWhite(color, (percentage - 70) / 100d);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, seriesIndex), defaultAlpha, null, null);
    }

    private static RgbColor ApplyUnstyledLineStrokeTint(RgbColor color)
    {
        (double hue, double saturation, double luminosity) = ToHsl(color);
        (double red, double green, double blue) = FromHsl(hue, saturation, luminosity * PptxChartMetricRules.UnstyledLineStrokeLuminanceFactor);
        return new RgbColor(ToTintByte(red), ToTintByte(green), ToTintByte(blue));
    }

    private static (double Hue, double Saturation, double Luminosity) ToHsl(RgbColor color)
    {
        double red = color.Red / 255d;
        double green = color.Green / 255d;
        double blue = color.Blue / 255d;
        double max = System.Math.Max(red, System.Math.Max(green, blue));
        double min = System.Math.Min(red, System.Math.Min(green, blue));
        double luminosity = (max + min) / 2d;
        if (max == min)
        {
            return (0d, 0d, luminosity);
        }

        double delta = max - min;
        double saturation = luminosity > 0.5d ? delta / (2d - max - min) : delta / (max + min);
        double hue = max == red
            ? PositiveModulo((green - blue) / delta, 6d)
            : max == green ? (blue - red) / delta + 2d : (red - green) / delta + 4d;
        return (hue * 60d, saturation, luminosity);
    }

    private static double PositiveModulo(double value, double modulus)
    {
        double remainder = value % modulus;
        return remainder < 0d ? remainder + modulus : remainder;
    }

    private static (double Red, double Green, double Blue) FromHsl(double hue, double saturation, double luminosity)
    {
        if (saturation == 0d)
        {
            return (luminosity, luminosity, luminosity);
        }

        double h = hue / 360d;
        double q = luminosity < 0.5d ? luminosity * (1d + saturation) : luminosity + saturation - luminosity * saturation;
        double p = 2d * luminosity - q;
        return (HueToRgb(p, q, h + 1d / 3d), HueToRgb(p, q, h), HueToRgb(p, q, h - 1d / 3d));

        static double HueToRgb(double p, double q, double t)
        {
            t %= 1d;
            if (t < 0d)
            {
                t += 1d;
            }

            if (t < 1d / 6d)
            {
                return p + (q - p) * 6d * t;
            }

            if (t < 0.5d)
            {
                return q;
            }

            if (t < 2d / 3d)
            {
                return p + (q - p) * (2d / 3d - t) * 6d;
            }

            return p;
        }
    }

    private static byte ToTintByte(double value)
    {
        return (byte)System.Math.Clamp((int)System.Math.Round(value * 255d, System.MidpointRounding.AwayFromZero), 0, 255);
    }

    // Automatic Office bar palettes use the six theme accents. Each repeated
    // cycle receives a shade/tint spaced over (-70%, +70%); Office truncates
    // the resulting percentage before applying it in linear light. Counts at
    // exact cycle boundaries include the next cycle in the spacing denominator.
    private static RgbColor ResolveAutomaticBarColor(RgbColor color, int categoryIndex, int valuePointCount)
    {
        const int accentCount = 6;
        int percentage = (int)(140L * (categoryIndex / accentCount + 1) / (valuePointCount / accentCount + 2));
        return percentage < 70
            ? LinearLightColor.ShadeTowardBlack(color, (30 + percentage) / 100d)
            : LinearLightColor.TintTowardWhite(color, (percentage - 70) / 100d);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, int valuePointCount)
    {
        if (varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null))
        {
            RgbColor paletteColor = ChartPalette(chartPalette, theme, colorMap, categoryIndex, valuePointCount);
            if (chartPalette.SingleAccent is null)
                paletteColor = ResolveAutomaticBarColor(paletteColor, categoryIndex, valuePointCount);

            return new ChartSeriesFill(paletteColor, 1d, null, null);
        }

        return ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int valuePointCount)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, valuePointCount);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int valuePointCount, double value, bool invertNegative)
    {
        if (value < 0d && invertNegative)
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d, null, null);
        }

        return ChartPointCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills, valuePointCount);
    }

    private static bool HasExplicitChartPointFill(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int seriesIndex, int categoryIndex)
    {
        return seriesIndex < pointFills.Count && pointFills[seriesIndex].ContainsKey(categoryIndex);
    }

    private static ChartSeriesStroke? ResolveNegativeBarFallbackStroke(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, int seriesIndex, int categoryIndex, double value, bool invertNegative)
    {
        return value < 0d && invertNegative && !HasExplicitChartPointStroke()
            ? ChartNegativeBarDefaultStroke
            : null;

        bool HasExplicitChartPointStroke()
        {
            return seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        }
    }
}
