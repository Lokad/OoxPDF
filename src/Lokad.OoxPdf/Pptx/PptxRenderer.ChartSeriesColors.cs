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

    private static bool ShouldShadeSingleSeriesVaryColors(int valuePointCount, bool verticalBars)
    {
        return verticalBars && valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsShadePointThreshold;
    }

    // Second-regime gate: twelve-plus points take the dark-plus-raw variation row.
    private static bool UseSecondVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsSecondRegimePointThreshold;
    }

    // Third-regime gate: eighteen-plus points take the dark/mid/light rows.

    // Fourth-regime gate: twenty-four-plus points take the dark/0.88/raw/replay rows.

    // Fifth-regime gate: thirty-plus points take the dark/0.85/0.95/fixed/palest rows.

    // Sixth-regime gate: thirty-six-plus points take the dark/0.82/0.91/raw/fixed/replay rows.
    private static bool UseSixthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimePointThreshold;
    }
    private static bool UseFifthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimePointThreshold;
    }
    private static bool UseFourthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimePointThreshold;
    }
    private static bool UseThirdVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimePointThreshold;
    }

    private static RgbColor ShadeThirdRegimeDarkSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeDarkShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeDarkShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeDarkShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeThirdRegimeMidSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveThirdRegimeLightFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 13)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot14Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot15Fill;
            return true;
        }

        if (categoryIndex == 15)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot16Fill;
            return true;
        }

        if (categoryIndex == 16)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 17)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot18Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourthRegimeReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 22)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeFifthRegimeMidSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeFifthRegimeLightSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero));
    }


    private static bool TryResolveFifthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifthRegimeFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 22)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifthRegimePalestFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeSixthRegimeLightSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero));
    }


    private static bool TryResolveSixthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixthRegimeFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixthRegimeReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeSecondRegimeSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsSecondRegimeShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsSecondRegimeShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsSecondRegimeShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    // Fixed overflow tints (dash7 through dash35 Office fills agree),
    // also serving as the fallback past the second-regime raw window, the third-regime light row and the fourth-regime replay row:
    // first-regime slots 7-10 (dash7/dash8/dash9/dash10 pairs agree), regime-2 tail slots 13-16 (dash13-dash17 pairs agree),
    // third-regime tail slots 19-22 (dash19-dash23 pairs agree), post-replay slots 25-28 (dash25 through dash29 agree)
    // and post-palest slots 31-34 (dash31 through dash35 agree); slot-35-plus keeps shaded cycling (sky-cyan single sample).
    private static bool TryResolveSingleSeriesVaryColorsOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot9Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot10Fill;
            return true;
        }

        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot13Fill;
            return true;
        }

        if (categoryIndex == 13)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot14Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot15Fill;
            return true;
        }

        if (categoryIndex == 15)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot19Fill;
            return true;
        }

        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot20Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot21Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot22Fill;
            return true;
        }

        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot28Fill;
            return true;
        }

        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot34Fill;
            return true;
        }

        fill = default;
        return false;
    }

    // Unstyled series strokes take the palette color through a 97.5% HSL luminance
    // modulation (Office 9/9 bytes exact on line series; same bytes on scatter and
    // radar outlines); fills keep raw colors.
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

    // Regime router: thirty-six-plus points take dark/0.82/0.91/raw/fixed/replay rows, thirty-plus points take dark/0.85/0.95/fixed/palest rows, twenty-four-plus points take dark/0.88/raw/replay rows, eighteen-plus points take dark/mid/light rows, twelve-plus points
    // shade slots 1-6 at 0.82 and leave slots 7-12 raw; slot-13-plus falls back through the overflow table (fixed through slot-22) and 0.88 cycling.
    private static RgbColor ResolveShadedSingleSeriesVaryColorsFill(RgbColor paletteColor, int categoryIndex, int valuePointCount)
    {
        if (UseSixthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveSixthRegimeDarkFill(categoryIndex, out RgbColor sixthDarkFill))
            {
                return sixthDarkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeSecondRegimeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return ShadeSixthRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return paletteColor;
            }

            if (TryResolveSixthRegimeFixedFill(categoryIndex, out RgbColor sixthFixedFill))
            {
                return sixthFixedFill;
            }

            if (TryResolveSixthRegimeReplayFill(categoryIndex, out RgbColor sixthReplayFill))
            {
                return sixthReplayFill;
            }
        }

        if (UseFifthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveFifthRegimeDarkFill(categoryIndex, out RgbColor fifthDarkFill))
            {
                return fifthDarkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeFifthRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return ShadeFifthRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveFifthRegimeFixedFill(categoryIndex, out RgbColor fifthFixedFill))
            {
                return fifthFixedFill;
            }

            if (TryResolveFifthRegimePalestFill(categoryIndex, out RgbColor fifthPalestFill))
            {
                return fifthPalestFill;
            }
        }

        if (UseFourthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveFourthRegimeDarkFill(categoryIndex, out RgbColor darkFill))
            {
                return darkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return paletteColor;
            }

            if (TryResolveFourthRegimeReplayFill(categoryIndex, out RgbColor replayFill))
            {
                return replayFill;
            }
        }

        if (UseThirdVaryColorsRegime(valuePointCount))
        {
            if (categoryIndex < 6)
            {
                return ShadeThirdRegimeDarkSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 12)
            {
                return ShadeThirdRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveThirdRegimeLightFill(categoryIndex, out RgbColor lightFill))
            {
                return lightFill;
            }
        }

        if (UseSecondVaryColorsRegime(valuePointCount))
        {
            if (categoryIndex < 6)
            {
                return ShadeSecondRegimeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 12)
            {
                return paletteColor;
            }
        }

        return TryResolveSingleSeriesVaryColorsOverflowFill(categoryIndex, out RgbColor overflowFill)
            ? overflowFill
            : ShadeSingleSeriesVaryColorsFill(paletteColor);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, int valuePointCount, bool shadeSingleSeriesVaryColors)
    {
        if (varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null))
        {
            RgbColor paletteColor = ChartPalette(chartPalette, theme, colorMap, categoryIndex);
            if (ShouldShadeSingleSeriesVaryColors(valuePointCount, shadeSingleSeriesVaryColors))
            {
                paletteColor = ResolveShadedSingleSeriesVaryColorsFill(paletteColor, categoryIndex, valuePointCount);
            }

            return new ChartSeriesFill(paletteColor, 1d, null, null);
        }

        return ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int valuePointCount, bool shadeSingleSeriesVaryColors)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, valuePointCount, shadeSingleSeriesVaryColors);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int valuePointCount, bool shadeSingleSeriesVaryColors, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d, null, null);
        }

        return ChartPointCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills, valuePointCount, shadeSingleSeriesVaryColors);
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
