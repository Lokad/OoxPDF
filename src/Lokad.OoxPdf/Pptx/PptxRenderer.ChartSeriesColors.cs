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

    // Seventh-regime gate: forty-two-plus points take the dark/mid/0.88/0.96/fixed/replay/palest rows.

    // Tenth-regime gate: sixty-plus points take the dark/fourth-replay/0.82/0.88/0.94/raw/fixed/seventh-replay/sixth-replay/overflow-replay rows.
    private static bool UseTenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimePointThreshold;
    }

    // Ninth-regime gate: fifty-four-plus points take the dark/0.76/0.84/0.90/0.96/fixed/fixed/palest/ninth rows.
    private static bool UseNinthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimePointThreshold;
    }

    // Eighth-regime gate: forty-eight-plus points take the dark/0.78/0.86/0.93/raw/light/fixed/palest rows.
    private static bool UseEighthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimePointThreshold;
    }
    private static bool UseSeventhVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimePointThreshold;
    }
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

    private static RgbColor ShadeSeventhRegimeLightSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeLightShadeFactor, System.MidpointRounding.AwayFromZero));
    }


    private static bool TryResolveSeventhRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventhRegimeMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventhRegimeFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventhRegimeReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventhRegimePalestFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeEighthRegimeMidSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeMidShadeFactor, System.MidpointRounding.AwayFromZero));
    }


    private static bool TryResolveEighthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighthRegimeLightFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighthRegimeFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighthRegimePalestFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeNinthRegimeSecondRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeNinthRegimeThirdRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeNinthRegimeFourthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveNinthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNinthRegimeFixed30Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNinthRegimeFixed36Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNinthRegimePalestFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNinthRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static bool TryResolveNinthRegimeReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot55Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot56Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot57Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot58Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeTenthRegimeFifthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveTenthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTenthRegimeFourthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTenthRegimeFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTenthRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTenthRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTenthRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot26Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot28Fill;
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

    // Fixed overflow tints (dash7 through dash53 Office fills agree),
    // also serving as the fallback past every regime row: first-regime slots 7-10, regime-2 tail slots 13-16,
    // third-regime tail slots 19-22, fourth-regime post-replay slots 25-28, fifth-regime post-palest slots 31-34
    // and sixth-regime post-replay slots 37-40 plus post-palest slots 43-46 plus post-palest slots 49-52 (dash37 through dash53 agree); slot-53-plus keeps shaded cycling (sky single sample).
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

        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot40Fill;
            return true;
        }

        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot46Fill;
            return true;
        }

        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot52Fill;
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

    // Regime router: forty-eight-plus points take dark/0.78/0.86/0.93/raw/light/fixed/palest rows, forty-two-plus points take dark/mid/0.88/0.96/fixed/replay/palest rows, thirty-six-plus points take dark/0.82/0.91/raw/fixed/replay rows, thirty-plus points take dark/0.85/0.95/fixed/palest rows, twenty-four-plus points take dark/0.88/raw/replay rows, eighteen-plus points take dark/mid/light rows, twelve-plus points
    // shade slots 1-6 at 0.82 and leave slots 7-12 raw; slot-13-plus falls back through the overflow table (fixed through slot-22) and 0.88 cycling.
    private static RgbColor ResolveShadedSingleSeriesVaryColorsFill(RgbColor paletteColor, int categoryIndex, int valuePointCount)
    {
        if (UseTenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTenthRegimeDarkFill(categoryIndex, out RgbColor tenthDarkFill))
            {
                return tenthDarkFill;
            }

            if (TryResolveTenthRegimeFourthFill(categoryIndex, out RgbColor tenthFourthFill))
            {
                return tenthFourthFill;
            }

            if (categoryIndex < 18)
            {
                return ShadeSecondRegimeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeTenthRegimeFifthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return paletteColor;
            }

            if (TryResolveTenthRegimeFixedFill(categoryIndex, out RgbColor tenthFixedFill))
            {
                return tenthFixedFill;
            }

            if (TryResolveTenthRegimeSeventhFill(categoryIndex, out RgbColor tenthSeventhFill))
            {
                return tenthSeventhFill;
            }

            if (TryResolveTenthRegimeSixthFill(categoryIndex, out RgbColor tenthSixthFill))
            {
                return tenthSixthFill;
            }

            if (TryResolveTenthRegimeOverflowFill(categoryIndex, out RgbColor tenthOverflowFill))
            {
                return tenthOverflowFill;
            }
        }
        if (UseNinthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveNinthRegimeDarkFill(categoryIndex, out RgbColor ninthDarkFill))
            {
                return ninthDarkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeNinthRegimeSecondRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return ShadeNinthRegimeThirdRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeNinthRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeSeventhRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveNinthRegimeFixed30Fill(categoryIndex, out RgbColor ninthFixed30Fill))
            {
                return ninthFixed30Fill;
            }

            if (TryResolveNinthRegimeFixed36Fill(categoryIndex, out RgbColor ninthFixed36Fill))
            {
                return ninthFixed36Fill;
            }

            if (TryResolveNinthRegimePalestFill(categoryIndex, out RgbColor ninthPalestFill))
            {
                return ninthPalestFill;
            }

            if (TryResolveNinthRegimeNinthFill(categoryIndex, out RgbColor ninthNinthFill))
            {
                return ninthNinthFill;
            }

            if (TryResolveNinthRegimeReplayFill(categoryIndex, out RgbColor ninthReplayFill))
            {
                return ninthReplayFill;
            }
        }
        if (UseEighthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveEighthRegimeDarkFill(categoryIndex, out RgbColor eighthDarkFill))
            {
                return eighthDarkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeThirdRegimeDarkSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return ShadeEighthRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeThirdRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return paletteColor;
            }

            if (TryResolveEighthRegimeLightFill(categoryIndex, out RgbColor eighthLightFill))
            {
                return eighthLightFill;
            }

            if (TryResolveEighthRegimeFixedFill(categoryIndex, out RgbColor eighthFixedFill))
            {
                return eighthFixedFill;
            }

            if (TryResolveEighthRegimePalestFill(categoryIndex, out RgbColor eighthPalestFill))
            {
                return eighthPalestFill;
            }
        }

        if (UseSeventhVaryColorsRegime(valuePointCount))
        {
            if (TryResolveSeventhRegimeDarkFill(categoryIndex, out RgbColor seventhDarkFill))
            {
                return seventhDarkFill;
            }

            if (TryResolveSeventhRegimeMidFill(categoryIndex, out RgbColor seventhMidFill))
            {
                return seventhMidFill;
            }

            if (categoryIndex < 18)
            {
                return ShadeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeSeventhRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveSeventhRegimeFixedFill(categoryIndex, out RgbColor seventhFixedFill))
            {
                return seventhFixedFill;
            }

            if (TryResolveSeventhRegimeReplayFill(categoryIndex, out RgbColor seventhReplayFill))
            {
                return seventhReplayFill;
            }

            if (TryResolveSeventhRegimePalestFill(categoryIndex, out RgbColor seventhPalestFill))
            {
                return seventhPalestFill;
            }
        }

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
