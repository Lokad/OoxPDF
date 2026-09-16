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

    // Twenty-first-regime gate: one-hundred-twenty-six-plus points keep dark rows, re-tint early singles with new fixed fills, and replay seventh/seventeenth/eighteenth/ninth/fifteenth/eleventh/eighth/thirteenth/twelfth/ninth-replay rows past thirty-four.

    private static bool UseTwentyFirstVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimePointThreshold;
    }

    // Twenty-sixth-regime gate: one-hundred-fifty-six-plus points keep earlier rows, and replay seventeenth/eighteenth/eleventh/fifth/sixteenth/twelfth/overflow/sixth/fifteenth rows past seventy-six.

    private static bool UseTwentySixthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentySixthRegimePointThreshold;
    }

    // Twenty-fifth-regime gate: one-hundred-fifty-plus points keep earlier rows, and replay nineteenth/seventeenth/twentythird/seventh/fifteenth/sixth/ninth/overflow/fourth/eighth/thirteenth/eleventh/twelfth/sixteenth/twentysecond rows past forty-six.

    private static bool UseTwentyFifthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentyFifthRegimePointThreshold;
    }

    // Twenty-fourth-regime gate: one-hundred-forty-four-plus points keep earlier rows, and replay fifteenth/nineteenth/seventeenth/twentyfirst/eighteenth/eleventh/fifth/twelfth/thirteenth/fourteenth/twentysecond rows past twenty-three.

    private static bool UseTwentyFourthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentyFourthRegimePointThreshold;
    }

    // Twenty-third-regime gate: one-hundred-thirty-eight-plus points keep earlier rows, and replay seventh/nineteenth/seventeenth/eighteenth/fourteenth/third/ninth/eleventh/fifth/overflow/eighth/sixth/fifteenth/tenth/twentysecond rows past twenty-nine.

    private static bool UseTwentyThirdVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimePointThreshold;
    }

    // Twenty-second-regime gate: one-hundred-thirty-two-plus points keep earlier rows, and replay nineteenth/seventeenth/eighteenth/eleventh/tenth/sixth/overflow/fourth/fifth/twelfth/fourteenth rows past forty.

    private static bool UseTwentySecondVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimePointThreshold;
    }

    // Twentieth-regime gate: one-hundred-twenty-plus points keep dark rows, re-tint early singles with new fixed fills, and replay seventh/seventeenth/eighteenth/ninth/fifteenth/eleventh/fourth/overflow/twelfth rows past twenty-nine.

    private static bool UseTwentiethVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimePointThreshold;
    }

    // Nineteenth-regime gate: one-hundred-fourteen-plus points keep dark rows, re-tint early singles with new fixed fills, and replay fifteenth/seventh/seventeenth/fifth/eleventh/fourth/overflow/fourteenth/twelfth/eighteenth/sixteenth rows past eighteen.

    private static bool UseNineteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimePointThreshold;
    }

    // Eighteenth-regime gate: one-hundred-eight-plus points keep dark/seventh rows, re-tint early singles with new fixed fills, replay raw accents at fifty-five-plus, and replay seventh/eighth/twelfth/overflow/sixth/eleventh/sixteenth rows past sixty.

    private static bool UseEighteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimePointThreshold;
    }

    // Seventeenth-regime gate: one-hundred-two-plus points keep dark/seventh/fourth/seventh rows, re-tint mid rows with new fixed fills, and replay tenth/sixth/ninth/eleventh/thirteenth/twelfth/overflow/sixteenth rows past sixty.

    private static bool UseSeventeenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimePointThreshold;
    }

    // Sixteenth-regime gate: ninety-six-plus points take the dark/seventh/fourth/seventh/0.85/0.88/0.93/0.96/raw/seventh/fixed/seventh/ninth/seventh/overflow/tenth/overflow/tail rows.
    private static bool UseSixteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimePointThreshold;
    }

    // Fifteenth-regime gate: ninety-plus points take the dark/fixed/fixed/0.80/0.85/0.90/0.94/0.98/thirteenth/fixed/twelfth/eighth/eleventh/fixed/fixed rows.
    private static bool UseFifteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimePointThreshold;
    }

    // Fourteenth-regime gate: eighty-four-plus points take the dark/sixth/0.77/0.82/0.87/0.91/0.96/raw/fixed/sixth/eleventh/sixth/fixed/overflow/tail rows.
    private static bool UseFourteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimePointThreshold;
    }

    // Thirteenth-regime gate: seventy-eight-plus points take the dark/fixed/0.78/0.83/0.88/0.93/0.98/fixed/eighth/seventh/fixed/eighth/fixed rows.
    private static bool UseThirteenthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimePointThreshold;
    }

    // Twelfth-regime gate: seventy-two-plus points take the dark/0.73/0.79/0.85/0.90/0.95/raw/fifth/fixed/fifth/fixed/overflow/tail rows.
    private static bool UseTwelfthVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimePointThreshold;
    }

    // Eleventh-regime gate: sixty-six-plus points take the dark/second/0.80/0.87/0.92/0.97/fixed rows.
    private static bool UseEleventhVaryColorsRegime(int valuePointCount)
    {
        return valuePointCount >= PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimePointThreshold;
    }

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

    private static bool TryResolveSixteenthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeFourthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 13)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 15)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 16)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 17)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeSeventhMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 22)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeSeventh25Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeFixed61Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeSeventh31Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeSeventh37Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeOverflow25Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot26Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot28Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeTenth59Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeOverflow43Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot43Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot44Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot45Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot46Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSixteenthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot95Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeSeventhMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 22)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeFifthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeShadeFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeThirteenthSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeSeventhLightFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeRawFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeFixed55Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeTenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeSixth25Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot25Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeThirteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot43Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot44Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot51Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot46Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeEleventhLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveSeventeenthRegimeSixteenthTailReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot95Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEarlySinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 16)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 17)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot18Fill;
            return true;
        }

        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot22Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeShadeSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot35Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeRawDriftFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeRawFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeTwelfthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot74Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeTailSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot107Fill;
            return true;
        }

        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeEarlySinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeMidSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 13)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot14Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot15Fill;
            return true;
        }

        if (categoryIndex == 15)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot16Fill;
            return true;
        }

        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot24Fill;
            return true;
        }

        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeShadeRowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot36Fill;
            return true;
        }

        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeLateSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot48Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeFifteenthEarlyFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 18)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 22)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot17Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeSeventhSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot7Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeSeventeenthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot47Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeSeventeenthRawFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeSeventeenthFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeFifthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeEleventhEarlyFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeFourthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeFifthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeEleventhMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeOverflowSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeFourteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeOverflowMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot31Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot33Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot34Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeTwelfthLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot72Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeOverflowLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot49Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot50Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot51Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot52Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeEighteenthTailReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot107Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeSixteenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveNineteenthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static bool TryResolveTwentiethRegimeEarlySinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 16)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 17)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot18Fill;
            return true;
        }

        if (categoryIndex == 21)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot22Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeMidSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 32)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 33)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot35Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeTailSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot120Fill;
            return true;
        }

        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeMidSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 138)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 139)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 140)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 141)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot136Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 144)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 145)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 146)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 147)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot136Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 150)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 151)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 152)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 153)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot136Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeLateSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 154)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySixthRegimeSlot155Fill;
            return true;
        }

        if (categoryIndex == 156)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySixthRegimeSlot157Fill;
            return true;
        }

        if (categoryIndex == 157)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySixthRegimeSlot158Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeSeventeenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeEighteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot60Fill;
            return true;
        }

        if (categoryIndex == 148)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot107Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeFifthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot24Fill;
            return true;
        }

        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeSixteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot66Fill;
            return true;
        }

        if (categoryIndex == 149)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 130)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 131)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot66Fill;
            return true;
        }

        if (categoryIndex == 142)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        if (categoryIndex == 143)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot72Fill;
            return true;
        }

        if (categoryIndex == 150)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 151)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 152)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 153)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot76Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot13Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot14Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot15Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        if (categoryIndex == 138)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot31Fill;
            return true;
        }

        if (categoryIndex == 139)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 140)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot33Fill;
            return true;
        }

        if (categoryIndex == 141)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot34Fill;
            return true;
        }

        if (categoryIndex == 144)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot43Fill;
            return true;
        }

        if (categoryIndex == 145)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot44Fill;
            return true;
        }

        if (categoryIndex == 146)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot51Fill;
            return true;
        }

        if (categoryIndex == 147)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot46Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySixthRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 132)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot79Fill;
            return true;
        }

        if (categoryIndex == 133)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot80Fill;
            return true;
        }

        if (categoryIndex == 134)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot81Fill;
            return true;
        }

        if (categoryIndex == 135)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot82Fill;
            return true;
        }

        if (categoryIndex == 136)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot83Fill;
            return true;
        }

        if (categoryIndex == 137)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot84Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeNineteenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeSeventeenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot48Fill;
            return true;
        }

        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeTwentyThirdFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyThirdRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot30Fill;
            return true;
        }

        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot28Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 130)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 131)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot7Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot8Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot9Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot10Fill;
            return true;
        }

        if (categoryIndex == 133)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 138)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot43Fill;
            return true;
        }

        if (categoryIndex == 139)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot44Fill;
            return true;
        }

        if (categoryIndex == 140)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot45Fill;
            return true;
        }

        if (categoryIndex == 141)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot46Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeFourthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeThirteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 132)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 134)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 135)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 137)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        if (categoryIndex == 144)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot67Fill;
            return true;
        }

        if (categoryIndex == 145)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot68Fill;
            return true;
        }

        if (categoryIndex == 146)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot69Fill;
            return true;
        }

        if (categoryIndex == 147)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot70Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeTwelfthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 136)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeSixteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 142)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot95Fill;
            return true;
        }

        if (categoryIndex == 143)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFifthRegimeTwentySecondSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 148)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot131Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 24)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 132)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot85Fill;
            return true;
        }

        if (categoryIndex == 133)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot86Fill;
            return true;
        }

        if (categoryIndex == 134)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot87Fill;
            return true;
        }

        if (categoryIndex == 135)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot88Fill;
            return true;
        }

        if (categoryIndex == 136)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot89Fill;
            return true;
        }

        if (categoryIndex == 137)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot90Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeNineteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot24Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeSeventeenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeTwentyFirstFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot53Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeEighteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot60Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot74Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot48Fill;
            return true;
        }

        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot60Fill;
            return true;
        }

        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 130)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 131)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        if (categoryIndex == 138)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot67Fill;
            return true;
        }

        if (categoryIndex == 139)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot68Fill;
            return true;
        }

        if (categoryIndex == 140)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot69Fill;
            return true;
        }

        if (categoryIndex == 141)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot70Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeFifthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeThirteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeFourteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFourthRegimeTwentySecondSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 142)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot131Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeSeventhSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 30)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot7Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeNineteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot30Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeSeventeenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot36Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeEighteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot53Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeFourteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeThirdFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot14Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot15Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot16Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirdRegimeSlot18Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeFifthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot14Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot15Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot19Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot20Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot21Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot22Fill;
            return true;
        }

        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot26Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeSixthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot85Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot86Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot87Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot88Fill;
            return true;
        }

        if (categoryIndex == 130)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot89Fill;
            return true;
        }

        if (categoryIndex == 131)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot90Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeTenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 132)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 133)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 134)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 135)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot64Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyThirdRegimeTwentySecondSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 136)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot131Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeNineteenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeSeventeenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeEighteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeTenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot60Fill;
            return true;
        }

        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot64Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot30Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot7Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot8Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot9Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot10Fill;
            return true;
        }

        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot13Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot14Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot15Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot25Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot26Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot27Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot28Fill;
            return true;
        }

        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot37Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot38Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot39Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot40Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeFourthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeFifthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeFourteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot83Fill;
            return true;
        }

        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot84Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeLateSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 130)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot131Fill;
            return true;
        }

        if (categoryIndex == 131)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot132Fill;
            return true;
        }

        if (categoryIndex == 135)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentySecondRegimeSlot136Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentySecondRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 132)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 133)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 134)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeEarlySinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot53Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeLateSinglesFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentyFirstRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeEighteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 124)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot107Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 19)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot14Fill;
            return true;
        }

        if (categoryIndex == 20)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot15Fill;
            return true;
        }

        if (categoryIndex == 23)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot18Fill;
            return true;
        }

        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot79Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot80Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot81Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot82Fill;
            return true;
        }

        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot83Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot84Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeFourteenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot52Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeNineteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot42Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot48Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNineteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeNinthReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 120)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot55Fill;
            return true;
        }

        if (categoryIndex == 121)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot56Fill;
            return true;
        }

        if (categoryIndex == 122)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot57Fill;
            return true;
        }

        if (categoryIndex == 123)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot58Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeSeventeenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 26)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 28)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 34)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 25)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 27)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot40Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeSixteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeThirteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot66Fill;
            return true;
        }

        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 119)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeTwentiethFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 125)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwentiethRegimeSlot120Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentyFirstRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 126)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot97Fill;
            return true;
        }

        if (categoryIndex == 127)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot98Fill;
            return true;
        }

        if (categoryIndex == 128)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot99Fill;
            return true;
        }

        if (categoryIndex == 129)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot100Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeSeventhSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 29)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeSeventeenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 31)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 35)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeSeventeenthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot36Fill;
            return true;
        }

        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeSeventeenthLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventeenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeEighteenthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeEighteenthFixedFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeNinthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeFifteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeNinthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeNinthLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeNinthOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 106)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeOverflowFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 108)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot31Fill;
            return true;
        }

        if (categoryIndex == 109)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 110)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot33Fill;
            return true;
        }

        if (categoryIndex == 111)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot34Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 112)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        if (categoryIndex == 113)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot72Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeNinthReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 114)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot55Fill;
            return true;
        }

        if (categoryIndex == 115)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot56Fill;
            return true;
        }

        if (categoryIndex == 116)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot57Fill;
            return true;
        }

        if (categoryIndex == 117)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsNinthRegimeReplaySlot58Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwentiethRegimeEighteenthTailReplayFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 118)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighteenthRegimeSlot107Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeTwelfthEarlyFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEighthMidFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeOverflowEarlyFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot13Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot14Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot15Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot16Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEighthLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 94)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 95)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEleventhLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 96)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 97)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 98)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 99)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeTwelfth71Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 100)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeEleventhSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 101)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeOverflowLateFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 102)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot49Fill;
            return true;
        }

        if (categoryIndex == 103)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot50Fill;
            return true;
        }

        if (categoryIndex == 104)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot51Fill;
            return true;
        }

        if (categoryIndex == 105)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot52Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEighteenthRegimeSixteenthSingleFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 107)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixteenthRegimeSlot96Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static RgbColor ShadeFifteenthRegimeFourthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveFifteenthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeSecondFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeThirdFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 12)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot13Fill;
            return true;
        }

        if (categoryIndex == 13)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot14Fill;
            return true;
        }

        if (categoryIndex == 14)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot15Fill;
            return true;
        }

        if (categoryIndex == 15)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot16Fill;
            return true;
        }

        if (categoryIndex == 16)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot17Fill;
            return true;
        }

        if (categoryIndex == 17)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot18Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeThirteenthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeFixed55Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeTwelfthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeEleventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeFixed79Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot79Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot80Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot81Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot82Fill;
            return true;
        }

        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot83Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot84Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFifteenthRegimeFixed85Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot85Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot86Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot87Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot88Fill;
            return true;
        }

        if (categoryIndex == 88)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot89Fill;
            return true;
        }

        if (categoryIndex == 89)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot90Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static bool TryResolveFifteenthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 90)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot91Fill;
            return true;
        }

        if (categoryIndex == 91)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot92Fill;
            return true;
        }

        if (categoryIndex == 92)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot93Fill;
            return true;
        }

        if (categoryIndex == 93)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifteenthRegimeSlot94Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeFourteenthRegimeThirdRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveFourteenthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeSixthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeFixed49Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeSixth25Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeEleventh49Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeSixth31Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSixthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeFixed73Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeOverflow37Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot37Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot38Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot39Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot40Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveFourteenthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 82)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot83Fill;
            return true;
        }

        if (categoryIndex == 83)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot84Fill;
            return true;
        }

        if (categoryIndex == 84)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot85Fill;
            return true;
        }

        if (categoryIndex == 85)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot86Fill;
            return true;
        }

        if (categoryIndex == 86)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot87Fill;
            return true;
        }

        if (categoryIndex == 87)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFourteenthRegimeSlot88Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeThirteenthRegimeThirdRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeThirteenthRegimeFourthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeThirteenthRegimeSixthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeThirteenthRegimeSeventhRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSeventhRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSeventhRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSeventhRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveThirteenthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeSecondFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeFixed43Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeEighthFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeSeventhFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot31Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot32Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot33Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot34Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot35Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsSeventhRegimeSlot36Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeFixed61Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimePalestFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEighthRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveThirteenthRegimeFixed73Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot76Fill;
            return true;
        }

        if (categoryIndex == 76)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot77Fill;
            return true;
        }

        if (categoryIndex == 77)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot78Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static bool TryResolveThirteenthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 78)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot79Fill;
            return true;
        }

        if (categoryIndex == 79)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot80Fill;
            return true;
        }

        if (categoryIndex == 80)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot81Fill;
            return true;
        }

        if (categoryIndex == 81)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsThirteenthRegimeSlot82Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeTwelfthRegimeSecondRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSecondRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeTwelfthRegimeThirdRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveTwelfthRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeFifth19Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot19Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot20Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot21Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot22Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot23Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot24Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeFixed49Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeFifth25Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot25Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot26Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot27Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot28Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot29Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsFifthRegimeSlot30Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeFixed61Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeOverflow31Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot31Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot32Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot33Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsOverflowSlot34Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveTwelfthRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 70)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot71Fill;
            return true;
        }

        if (categoryIndex == 71)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot72Fill;
            return true;
        }

        if (categoryIndex == 72)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot73Fill;
            return true;
        }

        if (categoryIndex == 73)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot74Fill;
            return true;
        }

        if (categoryIndex == 74)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot75Fill;
            return true;
        }

        if (categoryIndex == 75)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTwelfthRegimeSlot76Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static RgbColor ShadeEleventhRegimeThirdRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeThirdRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeEleventhRegimeFourthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFourthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeEleventhRegimeFifthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeFifthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static RgbColor ShadeEleventhRegimeSixthRowSingleSeriesVaryColorsFill(RgbColor color)
    {
        return new RgbColor(
            (byte)System.Math.Round(color.Red * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Green * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero),
            (byte)System.Math.Round(color.Blue * PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSixthRowShadeFactor, System.MidpointRounding.AwayFromZero));
    }

    private static bool TryResolveEleventhRegimeDarkFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 0)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot1Fill;
            return true;
        }

        if (categoryIndex == 1)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot2Fill;
            return true;
        }

        if (categoryIndex == 2)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot3Fill;
            return true;
        }

        if (categoryIndex == 3)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot4Fill;
            return true;
        }

        if (categoryIndex == 4)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot5Fill;
            return true;
        }

        if (categoryIndex == 5)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot6Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeSecondFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 6)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot7Fill;
            return true;
        }

        if (categoryIndex == 7)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot8Fill;
            return true;
        }

        if (categoryIndex == 8)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot9Fill;
            return true;
        }

        if (categoryIndex == 9)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot10Fill;
            return true;
        }

        if (categoryIndex == 10)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot11Fill;
            return true;
        }

        if (categoryIndex == 11)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot12Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeFixed37Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 36)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot37Fill;
            return true;
        }

        if (categoryIndex == 37)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot38Fill;
            return true;
        }

        if (categoryIndex == 38)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot39Fill;
            return true;
        }

        if (categoryIndex == 39)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot40Fill;
            return true;
        }

        if (categoryIndex == 40)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot41Fill;
            return true;
        }

        if (categoryIndex == 41)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot42Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeFixed43Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 42)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot43Fill;
            return true;
        }

        if (categoryIndex == 43)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot44Fill;
            return true;
        }

        if (categoryIndex == 44)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot45Fill;
            return true;
        }

        if (categoryIndex == 45)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot46Fill;
            return true;
        }

        if (categoryIndex == 46)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot47Fill;
            return true;
        }

        if (categoryIndex == 47)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot48Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeFixed49Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 48)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot49Fill;
            return true;
        }

        if (categoryIndex == 49)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot50Fill;
            return true;
        }

        if (categoryIndex == 50)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot51Fill;
            return true;
        }

        if (categoryIndex == 51)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot52Fill;
            return true;
        }

        if (categoryIndex == 52)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot53Fill;
            return true;
        }

        if (categoryIndex == 53)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot54Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeFixed55Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 54)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot55Fill;
            return true;
        }

        if (categoryIndex == 55)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot56Fill;
            return true;
        }

        if (categoryIndex == 56)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot57Fill;
            return true;
        }

        if (categoryIndex == 57)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot58Fill;
            return true;
        }

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot60Fill;
            return true;
        }

        fill = default;
        return false;
    }

    private static bool TryResolveEleventhRegimeFixed61Fill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot64Fill;
            return true;
        }

        if (categoryIndex == 64)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot65Fill;
            return true;
        }

        if (categoryIndex == 65)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot66Fill;
            return true;
        }

        fill = default;
        return false;
    }
    private static bool TryResolveEleventhRegimeTailFill(int categoryIndex, out RgbColor fill)
    {
        if (categoryIndex == 66)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot67Fill;
            return true;
        }

        if (categoryIndex == 67)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot68Fill;
            return true;
        }

        if (categoryIndex == 68)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot69Fill;
            return true;
        }

        if (categoryIndex == 69)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsEleventhRegimeSlot70Fill;
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

        if (categoryIndex == 58)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot59Fill;
            return true;
        }

        if (categoryIndex == 59)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot60Fill;
            return true;
        }

        if (categoryIndex == 60)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot61Fill;
            return true;
        }

        if (categoryIndex == 61)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot62Fill;
            return true;
        }

        if (categoryIndex == 62)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot63Fill;
            return true;
        }

        if (categoryIndex == 63)
        {
            fill = PptxChartMetricRules.SingleSeriesVaryColorsTenthRegimeSlot64Fill;
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
        if (UseTwentySixthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentySixthRegimeSeventeenthSingleFill(categoryIndex, out RgbColor twentySixthSeventeenthSingleFill))
            {
                return twentySixthSeventeenthSingleFill;
            }

            if (TryResolveTwentySixthRegimeEighteenthFill(categoryIndex, out RgbColor twentySixthEighteenthFill))
            {
                return twentySixthEighteenthFill;
            }

            if (TryResolveTwentySixthRegimeEleventhFill(categoryIndex, out RgbColor twentySixthEleventhFill))
            {
                return twentySixthEleventhFill;
            }

            if (TryResolveTwentySixthRegimeFifthFill(categoryIndex, out RgbColor twentySixthFifthFill))
            {
                return twentySixthFifthFill;
            }

            if (TryResolveTwentySixthRegimeSixteenthFill(categoryIndex, out RgbColor twentySixthSixteenthFill))
            {
                return twentySixthSixteenthFill;
            }

            if (TryResolveTwentySixthRegimeTwelfthFill(categoryIndex, out RgbColor twentySixthTwelfthFill))
            {
                return twentySixthTwelfthFill;
            }

            if (TryResolveTwentySixthRegimeOverflowFill(categoryIndex, out RgbColor twentySixthOverflowFill))
            {
                return twentySixthOverflowFill;
            }

            if (TryResolveTwentySixthRegimeSixthFill(categoryIndex, out RgbColor twentySixthSixthFill))
            {
                return twentySixthSixthFill;
            }

            if (TryResolveTwentySixthRegimeFifteenthFill(categoryIndex, out RgbColor twentySixthFifteenthFill))
            {
                return twentySixthFifteenthFill;
            }

            if (TryResolveTwentySixthRegimeLateSingleFill(categoryIndex, out RgbColor twentySixthLateSingleFill))
            {
                return twentySixthLateSingleFill;
            }

        }
        if (UseTwentyFifthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentyFifthRegimeNineteenthSingleFill(categoryIndex, out RgbColor twentyFifthNineteenthSingleFill))
            {
                return twentyFifthNineteenthSingleFill;
            }

            if (TryResolveTwentyFifthRegimeSeventeenthFill(categoryIndex, out RgbColor twentyFifthSeventeenthFill))
            {
                return twentyFifthSeventeenthFill;
            }

            if (TryResolveTwentyFifthRegimeTwentyThirdFill(categoryIndex, out RgbColor twentyFifthTwentyThirdFill))
            {
                return twentyFifthTwentyThirdFill;
            }

            if (TryResolveTwentyFifthRegimeSeventhFill(categoryIndex, out RgbColor twentyFifthSeventhFill))
            {
                return twentyFifthSeventhFill;
            }

            if (TryResolveTwentyFifthRegimeFifteenthFill(categoryIndex, out RgbColor twentyFifthFifteenthFill))
            {
                return twentyFifthFifteenthFill;
            }

            if (TryResolveTwentyFifthRegimeSixthFill(categoryIndex, out RgbColor twentyFifthSixthFill))
            {
                return twentyFifthSixthFill;
            }

            if (TryResolveTwentyFifthRegimeNinthFill(categoryIndex, out RgbColor twentyFifthNinthFill))
            {
                return twentyFifthNinthFill;
            }

            if (TryResolveTwentyFifthRegimeOverflowFill(categoryIndex, out RgbColor twentyFifthOverflowFill))
            {
                return twentyFifthOverflowFill;
            }

            if (TryResolveTwentyFifthRegimeFourthFill(categoryIndex, out RgbColor twentyFifthFourthFill))
            {
                return twentyFifthFourthFill;
            }

            if (TryResolveTwentyFifthRegimeEighthFill(categoryIndex, out RgbColor twentyFifthEighthFill))
            {
                return twentyFifthEighthFill;
            }

            if (TryResolveTwentyFifthRegimeThirteenthFill(categoryIndex, out RgbColor twentyFifthThirteenthFill))
            {
                return twentyFifthThirteenthFill;
            }

            if (TryResolveTwentyFifthRegimeEleventhFill(categoryIndex, out RgbColor twentyFifthEleventhFill))
            {
                return twentyFifthEleventhFill;
            }

            if (TryResolveTwentyFifthRegimeTwelfthSingleFill(categoryIndex, out RgbColor twentyFifthTwelfthSingleFill))
            {
                return twentyFifthTwelfthSingleFill;
            }

            if (TryResolveTwentyFifthRegimeSixteenthFill(categoryIndex, out RgbColor twentyFifthSixteenthFill))
            {
                return twentyFifthSixteenthFill;
            }

            if (TryResolveTwentyFifthRegimeTwentySecondSingleFill(categoryIndex, out RgbColor twentyFifthTwentySecondSingleFill))
            {
                return twentyFifthTwentySecondSingleFill;
            }

            if (TryResolveTwentyFifthRegimeTailFill(categoryIndex, out RgbColor twentyFifthTailFill))
            {
                return twentyFifthTailFill;
            }

        }
        if (UseTwentyFourthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentyFourthRegimeFifteenthFill(categoryIndex, out RgbColor twentyFourthFifteenthFill))
            {
                return twentyFourthFifteenthFill;
            }

            if (TryResolveTwentyFourthRegimeNineteenthFill(categoryIndex, out RgbColor twentyFourthNineteenthFill))
            {
                return twentyFourthNineteenthFill;
            }

            if (TryResolveTwentyFourthRegimeSeventeenthFill(categoryIndex, out RgbColor twentyFourthSeventeenthFill))
            {
                return twentyFourthSeventeenthFill;
            }

            if (TryResolveTwentyFourthRegimeTwentyFirstFill(categoryIndex, out RgbColor twentyFourthTwentyFirstFill))
            {
                return twentyFourthTwentyFirstFill;
            }

            if (TryResolveTwentyFourthRegimeEighteenthFill(categoryIndex, out RgbColor twentyFourthEighteenthFill))
            {
                return twentyFourthEighteenthFill;
            }

            if (TryResolveTwentyFourthRegimeEleventhFill(categoryIndex, out RgbColor twentyFourthEleventhFill))
            {
                return twentyFourthEleventhFill;
            }

            if (TryResolveTwentyFourthRegimeFifthFill(categoryIndex, out RgbColor twentyFourthFifthFill))
            {
                return twentyFourthFifthFill;
            }

            if (TryResolveTwentyFourthRegimeTwelfthFill(categoryIndex, out RgbColor twentyFourthTwelfthFill))
            {
                return twentyFourthTwelfthFill;
            }

            if (TryResolveTwentyFourthRegimeThirteenthFill(categoryIndex, out RgbColor twentyFourthThirteenthFill))
            {
                return twentyFourthThirteenthFill;
            }

            if (TryResolveTwentyFourthRegimeFourteenthFill(categoryIndex, out RgbColor twentyFourthFourteenthFill))
            {
                return twentyFourthFourteenthFill;
            }

            if (TryResolveTwentyFourthRegimeTwentySecondSingleFill(categoryIndex, out RgbColor twentyFourthTwentySecondSingleFill))
            {
                return twentyFourthTwentySecondSingleFill;
            }

            if (TryResolveTwentyFourthRegimeTailFill(categoryIndex, out RgbColor twentyFourthTailFill))
            {
                return twentyFourthTailFill;
            }
        }
        if (UseTwentyThirdVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentyThirdRegimeSeventhSingleFill(categoryIndex, out RgbColor twentyThirdSeventhSingleFill))
            {
                return twentyThirdSeventhSingleFill;
            }

            if (TryResolveTwentyThirdRegimeNineteenthFill(categoryIndex, out RgbColor twentyThirdNineteenthFill))
            {
                return twentyThirdNineteenthFill;
            }

            if (TryResolveTwentyThirdRegimeSeventeenthFill(categoryIndex, out RgbColor twentyThirdSeventeenthFill))
            {
                return twentyThirdSeventeenthFill;
            }

            if (TryResolveTwentyThirdRegimeEighteenthFill(categoryIndex, out RgbColor twentyThirdEighteenthFill))
            {
                return twentyThirdEighteenthFill;
            }

            if (TryResolveTwentyThirdRegimeFourteenthFill(categoryIndex, out RgbColor twentyThirdFourteenthFill))
            {
                return twentyThirdFourteenthFill;
            }

            if (TryResolveTwentyThirdRegimeThirdFill(categoryIndex, out RgbColor twentyThirdThirdFill))
            {
                return twentyThirdThirdFill;
            }

            if (TryResolveTwentyThirdRegimeNinthFill(categoryIndex, out RgbColor twentyThirdNinthFill))
            {
                return twentyThirdNinthFill;
            }

            if (TryResolveTwentyThirdRegimeEleventhFill(categoryIndex, out RgbColor twentyThirdEleventhFill))
            {
                return twentyThirdEleventhFill;
            }

            if (TryResolveTwentyThirdRegimeFifthFill(categoryIndex, out RgbColor twentyThirdFifthFill))
            {
                return twentyThirdFifthFill;
            }

            if (TryResolveTwentyThirdRegimeOverflowFill(categoryIndex, out RgbColor twentyThirdOverflowFill))
            {
                return twentyThirdOverflowFill;
            }

            if (TryResolveTwentyThirdRegimeEighthFill(categoryIndex, out RgbColor twentyThirdEighthFill))
            {
                return twentyThirdEighthFill;
            }

            if (TryResolveTwentyThirdRegimeSixthSingleFill(categoryIndex, out RgbColor twentyThirdSixthSingleFill))
            {
                return twentyThirdSixthSingleFill;
            }

            if (TryResolveTwentyThirdRegimeFifteenthFill(categoryIndex, out RgbColor twentyThirdFifteenthFill))
            {
                return twentyThirdFifteenthFill;
            }

            if (TryResolveTwentyThirdRegimeTenthFill(categoryIndex, out RgbColor twentyThirdTenthFill))
            {
                return twentyThirdTenthFill;
            }

            if (TryResolveTwentyThirdRegimeTwentySecondSingleFill(categoryIndex, out RgbColor twentyThirdTwentySecondSingleFill))
            {
                return twentyThirdTwentySecondSingleFill;
            }

            if (TryResolveTwentyThirdRegimeMidSinglesFill(categoryIndex, out RgbColor twentyThirdMidSinglesFill))
            {
                return twentyThirdMidSinglesFill;
            }

            if (TryResolveTwentyThirdRegimeTailFill(categoryIndex, out RgbColor twentyThirdTailFill))
            {
                return twentyThirdTailFill;
            }

        }
        if (UseTwentySecondVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentySecondRegimeNineteenthSingleFill(categoryIndex, out RgbColor twentySecondNineteenthSingleFill))
            {
                return twentySecondNineteenthSingleFill;
            }

            if (TryResolveTwentySecondRegimeSeventeenthFill(categoryIndex, out RgbColor twentySecondSeventeenthFill))
            {
                return twentySecondSeventeenthFill;
            }

            if (TryResolveTwentySecondRegimeEighteenthFill(categoryIndex, out RgbColor twentySecondEighteenthFill))
            {
                return twentySecondEighteenthFill;
            }

            if (TryResolveTwentySecondRegimeEleventhFill(categoryIndex, out RgbColor twentySecondEleventhFill))
            {
                return twentySecondEleventhFill;
            }

            if (TryResolveTwentySecondRegimeTenthFill(categoryIndex, out RgbColor twentySecondTenthFill))
            {
                return twentySecondTenthFill;
            }

            if (TryResolveTwentySecondRegimeSixthFill(categoryIndex, out RgbColor twentySecondSixthFill))
            {
                return twentySecondSixthFill;
            }

            if (TryResolveTwentySecondRegimeOverflowFill(categoryIndex, out RgbColor twentySecondOverflowFill))
            {
                return twentySecondOverflowFill;
            }

            if (TryResolveTwentySecondRegimeFourthFill(categoryIndex, out RgbColor twentySecondFourthFill))
            {
                return twentySecondFourthFill;
            }

            if (TryResolveTwentySecondRegimeFifthFill(categoryIndex, out RgbColor twentySecondFifthFill))
            {
                return twentySecondFifthFill;
            }

            if (TryResolveTwentySecondRegimeTwelfthFill(categoryIndex, out RgbColor twentySecondTwelfthFill))
            {
                return twentySecondTwelfthFill;
            }

            if (TryResolveTwentySecondRegimeFourteenthFill(categoryIndex, out RgbColor twentySecondFourteenthFill))
            {
                return twentySecondFourteenthFill;
            }

            if (TryResolveTwentySecondRegimeLateSingleFill(categoryIndex, out RgbColor twentySecondLateSingleFill))
            {
                return twentySecondLateSingleFill;
            }

            if (TryResolveTwentySecondRegimeTailFill(categoryIndex, out RgbColor twentySecondTailFill))
            {
                return twentySecondTailFill;
            }

        }
        if (UseTwentyFirstVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentyFirstRegimeEarlySinglesFill(categoryIndex, out RgbColor twentyFirstEarlySinglesFill))
            {
                return twentyFirstEarlySinglesFill;
            }

            if (TryResolveTwentyFirstRegimeLateSinglesFill(categoryIndex, out RgbColor twentyFirstLateSinglesFill))
            {
                return twentyFirstLateSinglesFill;
            }

            if (TryResolveTwentyFirstRegimeEighteenthFill(categoryIndex, out RgbColor twentyFirstEighteenthFill))
            {
                return twentyFirstEighteenthFill;
            }

            if (TryResolveTwentyFirstRegimeEighthFill(categoryIndex, out RgbColor twentyFirstEighthFill))
            {
                return twentyFirstEighthFill;
            }

            if (TryResolveTwentyFirstRegimeFifteenthFill(categoryIndex, out RgbColor twentyFirstFifteenthFill))
            {
                return twentyFirstFifteenthFill;
            }

            if (TryResolveTwentyFirstRegimeFourteenthSingleFill(categoryIndex, out RgbColor twentyFirstFourteenthSingleFill))
            {
                return twentyFirstFourteenthSingleFill;
            }

            if (TryResolveTwentyFirstRegimeNineteenthFill(categoryIndex, out RgbColor twentyFirstNineteenthFill))
            {
                return twentyFirstNineteenthFill;
            }

            if (TryResolveTwentyFirstRegimeNinthReplayFill(categoryIndex, out RgbColor twentyFirstNinthReplayFill))
            {
                return twentyFirstNinthReplayFill;
            }

            if (TryResolveTwentyFirstRegimeSeventeenthFill(categoryIndex, out RgbColor twentyFirstSeventeenthFill))
            {
                return twentyFirstSeventeenthFill;
            }

            if (TryResolveTwentyFirstRegimeSeventhFill(categoryIndex, out RgbColor twentyFirstSeventhFill))
            {
                return twentyFirstSeventhFill;
            }

            if (TryResolveTwentyFirstRegimeSixteenthFill(categoryIndex, out RgbColor twentyFirstSixteenthFill))
            {
                return twentyFirstSixteenthFill;
            }

            if (TryResolveTwentyFirstRegimeThirteenthFill(categoryIndex, out RgbColor twentyFirstThirteenthFill))
            {
                return twentyFirstThirteenthFill;
            }

            if (TryResolveTwentyFirstRegimeTwelfthFill(categoryIndex, out RgbColor twentyFirstTwelfthFill))
            {
                return twentyFirstTwelfthFill;
            }

            if (TryResolveTwentyFirstRegimeTwentiethFill(categoryIndex, out RgbColor twentyFirstTwentiethFill))
            {
                return twentyFirstTwentiethFill;
            }

            if (TryResolveTwentyFirstRegimeTailFill(categoryIndex, out RgbColor twentyFirstTailFill))
            {
                return twentyFirstTailFill;
            }

        }
        if (UseTwentiethVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwentiethRegimeEarlySinglesFill(categoryIndex, out RgbColor twentiethEarlySinglesFill))
            {
                return twentiethEarlySinglesFill;
            }

            if (TryResolveTwentiethRegimeMidSinglesFill(categoryIndex, out RgbColor twentiethMidSinglesFill))
            {
                return twentiethMidSinglesFill;
            }

            if (TryResolveTwentiethRegimeTailSingleFill(categoryIndex, out RgbColor twentiethTailSingleFill))
            {
                return twentiethTailSingleFill;
            }

            if (TryResolveTwentiethRegimeSeventhSingleFill(categoryIndex, out RgbColor twentiethSeventhSingleFill))
            {
                return twentiethSeventhSingleFill;
            }

            if (TryResolveTwentiethRegimeSeventeenthSingleFill(categoryIndex, out RgbColor twentiethSeventeenthSingleFill))
            {
                return twentiethSeventeenthSingleFill;
            }

            if (TryResolveTwentiethRegimeSeventeenthMidFill(categoryIndex, out RgbColor twentiethSeventeenthMidFill))
            {
                return twentiethSeventeenthMidFill;
            }

            if (TryResolveTwentiethRegimeSeventeenthLateFill(categoryIndex, out RgbColor twentiethSeventeenthLateFill))
            {
                return twentiethSeventeenthLateFill;
            }

            if (TryResolveTwentiethRegimeEighteenthMidFill(categoryIndex, out RgbColor twentiethEighteenthMidFill))
            {
                return twentiethEighteenthMidFill;
            }

            if (TryResolveTwentiethRegimeEighteenthFixedFill(categoryIndex, out RgbColor twentiethEighteenthFixedFill))
            {
                return twentiethEighteenthFixedFill;
            }

            if (TryResolveTwentiethRegimeNinthFill(categoryIndex, out RgbColor twentiethNinthFill))
            {
                return twentiethNinthFill;
            }

            if (TryResolveTwentiethRegimeFifteenthFill(categoryIndex, out RgbColor twentiethFifteenthFill))
            {
                return twentiethFifteenthFill;
            }

            if (TryResolveTwentiethRegimeNinthMidFill(categoryIndex, out RgbColor twentiethNinthMidFill))
            {
                return twentiethNinthMidFill;
            }

            if (TryResolveTwentiethRegimeEleventhFill(categoryIndex, out RgbColor twentiethEleventhFill))
            {
                return twentiethEleventhFill;
            }

            if (TryResolveTwentiethRegimeNinthLateFill(categoryIndex, out RgbColor twentiethNinthLateFill))
            {
                return twentiethNinthLateFill;
            }

            if (TryResolveTwentiethRegimeSeventhFill(categoryIndex, out RgbColor twentiethSeventhFill))
            {
                return twentiethSeventhFill;
            }

            if (TryResolveTwentiethRegimeNinthOverflowFill(categoryIndex, out RgbColor twentiethNinthOverflowFill))
            {
                return twentiethNinthOverflowFill;
            }

            if (TryResolveTwentiethRegimeOverflowFill(categoryIndex, out RgbColor twentiethOverflowFill))
            {
                return twentiethOverflowFill;
            }

            if (TryResolveTwentiethRegimeTwelfthFill(categoryIndex, out RgbColor twentiethTwelfthFill))
            {
                return twentiethTwelfthFill;
            }

            if (TryResolveTwentiethRegimeNinthReplayFill(categoryIndex, out RgbColor twentiethNinthReplayFill))
            {
                return twentiethNinthReplayFill;
            }

            if (TryResolveTwentiethRegimeEighteenthTailReplayFill(categoryIndex, out RgbColor twentiethEighteenthTailReplayFill))
            {
                return twentiethEighteenthTailReplayFill;
            }

        }
        if (UseNineteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveNineteenthRegimeEarlySinglesFill(categoryIndex, out RgbColor nineteenthEarlySinglesFill))
            {
                return nineteenthEarlySinglesFill;
            }

            if (TryResolveNineteenthRegimeMidSinglesFill(categoryIndex, out RgbColor nineteenthMidSinglesFill))
            {
                return nineteenthMidSinglesFill;
            }

            if (TryResolveNineteenthRegimeShadeRowFill(categoryIndex, out RgbColor nineteenthShadeRowFill))
            {
                return nineteenthShadeRowFill;
            }

            if (TryResolveNineteenthRegimeLateSinglesFill(categoryIndex, out RgbColor nineteenthLateSinglesFill))
            {
                return nineteenthLateSinglesFill;
            }

            if (TryResolveNineteenthRegimeFifteenthEarlyFill(categoryIndex, out RgbColor nineteenthFifteenthEarlyFill))
            {
                return nineteenthFifteenthEarlyFill;
            }

            if (TryResolveNineteenthRegimeSeventhSingleFill(categoryIndex, out RgbColor nineteenthSeventhSingleFill))
            {
                return nineteenthSeventhSingleFill;
            }

            if (TryResolveNineteenthRegimeSeventeenthMidFill(categoryIndex, out RgbColor nineteenthSeventeenthMidFill))
            {
                return nineteenthSeventeenthMidFill;
            }

            if (TryResolveNineteenthRegimeSeventeenthRawFill(categoryIndex, out RgbColor nineteenthSeventeenthRawFill))
            {
                return nineteenthSeventeenthRawFill;
            }

            if (TryResolveNineteenthRegimeSeventeenthFixedFill(categoryIndex, out RgbColor nineteenthSeventeenthFixedFill))
            {
                return nineteenthSeventeenthFixedFill;
            }

            if (TryResolveNineteenthRegimeFifthFill(categoryIndex, out RgbColor nineteenthFifthFill))
            {
                return nineteenthFifthFill;
            }

            if (TryResolveNineteenthRegimeEleventhEarlyFill(categoryIndex, out RgbColor nineteenthEleventhEarlyFill))
            {
                return nineteenthEleventhEarlyFill;
            }

            if (TryResolveNineteenthRegimeFourthFill(categoryIndex, out RgbColor nineteenthFourthFill))
            {
                return nineteenthFourthFill;
            }

            if (TryResolveNineteenthRegimeFifthMidFill(categoryIndex, out RgbColor nineteenthFifthMidFill))
            {
                return nineteenthFifthMidFill;
            }

            if (TryResolveNineteenthRegimeEleventhMidFill(categoryIndex, out RgbColor nineteenthEleventhMidFill))
            {
                return nineteenthEleventhMidFill;
            }

            if (TryResolveNineteenthRegimeOverflowSingleFill(categoryIndex, out RgbColor nineteenthOverflowSingleFill))
            {
                return nineteenthOverflowSingleFill;
            }

            if (TryResolveNineteenthRegimeFourteenthFill(categoryIndex, out RgbColor nineteenthFourteenthFill))
            {
                return nineteenthFourteenthFill;
            }

            if (TryResolveNineteenthRegimeOverflowMidFill(categoryIndex, out RgbColor nineteenthOverflowMidFill))
            {
                return nineteenthOverflowMidFill;
            }

            if (TryResolveNineteenthRegimeTwelfthLateFill(categoryIndex, out RgbColor nineteenthTwelfthLateFill))
            {
                return nineteenthTwelfthLateFill;
            }

            if (TryResolveNineteenthRegimeOverflowLateFill(categoryIndex, out RgbColor nineteenthOverflowLateFill))
            {
                return nineteenthOverflowLateFill;
            }

            if (TryResolveNineteenthRegimeEighteenthTailReplayFill(categoryIndex, out RgbColor nineteenthEighteenthTailReplayFill))
            {
                return nineteenthEighteenthTailReplayFill;
            }

            if (TryResolveNineteenthRegimeSixteenthSingleFill(categoryIndex, out RgbColor nineteenthSixteenthSingleFill))
            {
                return nineteenthSixteenthSingleFill;
            }

            if (TryResolveNineteenthRegimeTailFill(categoryIndex, out RgbColor nineteenthTailFill))
            {
                return nineteenthTailFill;
            }

        }
        if (UseEighteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveEighteenthRegimeEarlySinglesFill(categoryIndex, out RgbColor eighteenthEarlySinglesFill))
            {
                return eighteenthEarlySinglesFill;
            }

            if (TryResolveEighteenthRegimeShadeSinglesFill(categoryIndex, out RgbColor eighteenthShadeSinglesFill))
            {
                return eighteenthShadeSinglesFill;
            }

            if (TryResolveEighteenthRegimeRawDriftFill(categoryIndex, out RgbColor eighteenthRawDriftFill))
            {
                return eighteenthRawDriftFill;
            }

            if (TryResolveEighteenthRegimeRawFill(categoryIndex, out RgbColor eighteenthRawFill))
            {
                return eighteenthRawFill;
            }

            if (TryResolveEighteenthRegimeTwelfthSingleFill(categoryIndex, out RgbColor eighteenthTwelfthSingleFill))
            {
                return eighteenthTwelfthSingleFill;
            }

            if (TryResolveEighteenthRegimeTailSingleFill(categoryIndex, out RgbColor eighteenthTailSingleFill))
            {
                return eighteenthTailSingleFill;
            }

            if (TryResolveEighteenthRegimeSeventhFill(categoryIndex, out RgbColor eighteenthSeventhFill))
            {
                return eighteenthSeventhFill;
            }

            if (TryResolveEighteenthRegimeEighthFill(categoryIndex, out RgbColor eighteenthEighthFill))
            {
                return eighteenthEighthFill;
            }

            if (TryResolveEighteenthRegimeTwelfthEarlyFill(categoryIndex, out RgbColor eighteenthTwelfthEarlyFill))
            {
                return eighteenthTwelfthEarlyFill;
            }

            if (TryResolveEighteenthRegimeEighthMidFill(categoryIndex, out RgbColor eighteenthEighthMidFill))
            {
                return eighteenthEighthMidFill;
            }

            if (TryResolveEighteenthRegimeOverflowEarlyFill(categoryIndex, out RgbColor eighteenthOverflowEarlyFill))
            {
                return eighteenthOverflowEarlyFill;
            }

            if (TryResolveEighteenthRegimeSixthFill(categoryIndex, out RgbColor eighteenthSixthFill))
            {
                return eighteenthSixthFill;
            }

            if (TryResolveEighteenthRegimeEighthLateFill(categoryIndex, out RgbColor eighteenthEighthLateFill))
            {
                return eighteenthEighthLateFill;
            }

            if (TryResolveEighteenthRegimeEleventhLateFill(categoryIndex, out RgbColor eighteenthEleventhLateFill))
            {
                return eighteenthEleventhLateFill;
            }

            if (TryResolveEighteenthRegimeTwelfth71Fill(categoryIndex, out RgbColor eighteenthTwelfth71Fill))
            {
                return eighteenthTwelfth71Fill;
            }

            if (TryResolveEighteenthRegimeEleventhSingleFill(categoryIndex, out RgbColor eighteenthEleventhSingleFill))
            {
                return eighteenthEleventhSingleFill;
            }

            if (TryResolveEighteenthRegimeOverflowLateFill(categoryIndex, out RgbColor eighteenthOverflowLateFill))
            {
                return eighteenthOverflowLateFill;
            }

            if (TryResolveEighteenthRegimeSixteenthSingleFill(categoryIndex, out RgbColor eighteenthSixteenthSingleFill))
            {
                return eighteenthSixteenthSingleFill;
            }

        }
        if (UseSeventeenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveSeventeenthRegimeSeventhMidFill(categoryIndex, out RgbColor seventeenthSeventhMidFill))
            {
                return seventeenthSeventhMidFill;
            }

            if (TryResolveSeventeenthRegimeFifthMidFill(categoryIndex, out RgbColor seventeenthFifthMidFill))
            {
                return seventeenthFifthMidFill;
            }

            if (TryResolveSeventeenthRegimeShadeFill(categoryIndex, out RgbColor seventeenthShadeFill))
            {
                return seventeenthShadeFill;
            }

            if (TryResolveSeventeenthRegimeThirteenthSixthFill(categoryIndex, out RgbColor seventeenthThirteenthSixthFill))
            {
                return seventeenthThirteenthSixthFill;
            }

            if (TryResolveSeventeenthRegimeSeventhLightFill(categoryIndex, out RgbColor seventeenthSeventhLightFill))
            {
                return seventeenthSeventhLightFill;
            }

            if (TryResolveSeventeenthRegimeRawFill(categoryIndex, out RgbColor seventeenthRawFill))
            {
                return seventeenthRawFill;
            }

            if (TryResolveSeventeenthRegimeFixed55Fill(categoryIndex, out RgbColor seventeenthFixed55Fill))
            {
                return seventeenthFixed55Fill;
            }

            if (TryResolveSeventeenthRegimeTenthFill(categoryIndex, out RgbColor seventeenthTenthFill))
            {
                return seventeenthTenthFill;
            }

            if (TryResolveSeventeenthRegimeSixth25Fill(categoryIndex, out RgbColor seventeenthSixth25Fill))
            {
                return seventeenthSixth25Fill;
            }

            if (TryResolveSeventeenthRegimeNinthFill(categoryIndex, out RgbColor seventeenthNinthFill))
            {
                return seventeenthNinthFill;
            }

            if (TryResolveSeventeenthRegimeEleventhFill(categoryIndex, out RgbColor seventeenthEleventhFill))
            {
                return seventeenthEleventhFill;
            }

            if (TryResolveSeventeenthRegimeThirteenthFill(categoryIndex, out RgbColor seventeenthThirteenthFill))
            {
                return seventeenthThirteenthFill;
            }

            if (TryResolveSeventeenthRegimeTwelfthFill(categoryIndex, out RgbColor seventeenthTwelfthFill))
            {
                return seventeenthTwelfthFill;
            }

            if (TryResolveSeventeenthRegimeOverflowFill(categoryIndex, out RgbColor seventeenthOverflowFill))
            {
                return seventeenthOverflowFill;
            }

            if (TryResolveSeventeenthRegimeEleventhLateFill(categoryIndex, out RgbColor seventeenthEleventhLateFill))
            {
                return seventeenthEleventhLateFill;
            }

            if (TryResolveSeventeenthRegimeSixteenthTailReplayFill(categoryIndex, out RgbColor seventeenthSixteenthTailReplayFill))
            {
                return seventeenthSixteenthTailReplayFill;
            }

        }
        if (UseSixteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveSixteenthRegimeDarkFill(categoryIndex, out RgbColor sixteenthDarkFill))
            {
                return sixteenthDarkFill;
            }

            if (TryResolveSixteenthRegimeSeventhFill(categoryIndex, out RgbColor sixteenthSeventhFill))
            {
                return sixteenthSeventhFill;
            }

            if (TryResolveSixteenthRegimeFourthFill(categoryIndex, out RgbColor sixteenthFourthFill))
            {
                return sixteenthFourthFill;
            }

            if (TryResolveSixteenthRegimeSeventhMidFill(categoryIndex, out RgbColor sixteenthSeventhMidFill))
            {
                return sixteenthSeventhMidFill;
            }

            if (categoryIndex < 30)
            {
                return ShadeFifthRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 42)
            {
                return ShadeThirteenthRegimeSixthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 48)
            {
                return ShadeSeventhRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 54)
            {
                return paletteColor;
            }

            if (TryResolveSixteenthRegimeSeventh25Fill(categoryIndex, out RgbColor sixteenthSeventh25Fill))
            {
                return sixteenthSeventh25Fill;
            }

            if (TryResolveSixteenthRegimeFixed61Fill(categoryIndex, out RgbColor sixteenthFixed61Fill))
            {
                return sixteenthFixed61Fill;
            }

            if (TryResolveSixteenthRegimeSeventh31Fill(categoryIndex, out RgbColor sixteenthSeventh31Fill))
            {
                return sixteenthSeventh31Fill;
            }

            if (TryResolveSixteenthRegimeNinthFill(categoryIndex, out RgbColor sixteenthNinthFill))
            {
                return sixteenthNinthFill;
            }

            if (TryResolveSixteenthRegimeSeventh37Fill(categoryIndex, out RgbColor sixteenthSeventh37Fill))
            {
                return sixteenthSeventh37Fill;
            }

            if (TryResolveSixteenthRegimeOverflow25Fill(categoryIndex, out RgbColor sixteenthOverflow25Fill))
            {
                return sixteenthOverflow25Fill;
            }

            if (TryResolveSixteenthRegimeTenth59Fill(categoryIndex, out RgbColor sixteenthTenth59Fill))
            {
                return sixteenthTenth59Fill;
            }

            if (TryResolveSixteenthRegimeOverflow43Fill(categoryIndex, out RgbColor sixteenthOverflow43Fill))
            {
                return sixteenthOverflow43Fill;
            }

            if (TryResolveSixteenthRegimeTailFill(categoryIndex, out RgbColor sixteenthTailFill))
            {
                return sixteenthTailFill;
            }
        }
        if (UseFifteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveFifteenthRegimeDarkFill(categoryIndex, out RgbColor fifteenthDarkFill))
            {
                return fifteenthDarkFill;
            }

            if (TryResolveFifteenthRegimeSecondFill(categoryIndex, out RgbColor fifteenthSecondFill))
            {
                return fifteenthSecondFill;
            }

            if (TryResolveFifteenthRegimeThirdFill(categoryIndex, out RgbColor fifteenthThirdFill))
            {
                return fifteenthThirdFill;
            }

            if (categoryIndex < 24)
            {
                return ShadeFifteenthRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeFifthRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeNinthRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 42)
            {
                return ShadeTenthRegimeFifthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 48)
            {
                return ShadeThirteenthRegimeSeventhRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveFifteenthRegimeThirteenthFill(categoryIndex, out RgbColor fifteenthThirteenthFill))
            {
                return fifteenthThirteenthFill;
            }

            if (TryResolveFifteenthRegimeFixed55Fill(categoryIndex, out RgbColor fifteenthFixed55Fill))
            {
                return fifteenthFixed55Fill;
            }

            if (TryResolveFifteenthRegimeTwelfthFill(categoryIndex, out RgbColor fifteenthTwelfthFill))
            {
                return fifteenthTwelfthFill;
            }

            if (TryResolveFifteenthRegimeEighthFill(categoryIndex, out RgbColor fifteenthEighthFill))
            {
                return fifteenthEighthFill;
            }

            if (TryResolveFifteenthRegimeEleventhFill(categoryIndex, out RgbColor fifteenthEleventhFill))
            {
                return fifteenthEleventhFill;
            }

            if (TryResolveFifteenthRegimeFixed79Fill(categoryIndex, out RgbColor fifteenthFixed79Fill))
            {
                return fifteenthFixed79Fill;
            }

            if (TryResolveFifteenthRegimeFixed85Fill(categoryIndex, out RgbColor fifteenthFixed85Fill))
            {
                return fifteenthFixed85Fill;
            }

            if (TryResolveFifteenthRegimeTailFill(categoryIndex, out RgbColor fifteenthTailFill))
            {
                return fifteenthTailFill;
            }
        }
        if (UseFourteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveFourteenthRegimeDarkFill(categoryIndex, out RgbColor fourteenthDarkFill))
            {
                return fourteenthDarkFill;
            }

            if (TryResolveFourteenthRegimeSixthFill(categoryIndex, out RgbColor fourteenthSixthFill))
            {
                return fourteenthSixthFill;
            }

            if (categoryIndex < 18)
            {
                return ShadeFourteenthRegimeThirdRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeSecondRegimeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeEleventhRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeSixthRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 42)
            {
                return ShadeSeventhRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 48)
            {
                return paletteColor;
            }

            if (TryResolveFourteenthRegimeFixed49Fill(categoryIndex, out RgbColor fourteenthFixed49Fill))
            {
                return fourteenthFixed49Fill;
            }

            if (TryResolveFourteenthRegimeSixth25Fill(categoryIndex, out RgbColor fourteenthSixth25Fill))
            {
                return fourteenthSixth25Fill;
            }

            if (TryResolveFourteenthRegimeEleventh49Fill(categoryIndex, out RgbColor fourteenthEleventh49Fill))
            {
                return fourteenthEleventh49Fill;
            }

            if (TryResolveFourteenthRegimeSixth31Fill(categoryIndex, out RgbColor fourteenthSixth31Fill))
            {
                return fourteenthSixth31Fill;
            }

            if (TryResolveFourteenthRegimeFixed73Fill(categoryIndex, out RgbColor fourteenthFixed73Fill))
            {
                return fourteenthFixed73Fill;
            }

            if (TryResolveFourteenthRegimeOverflow37Fill(categoryIndex, out RgbColor fourteenthOverflow37Fill))
            {
                return fourteenthOverflow37Fill;
            }

            if (TryResolveFourteenthRegimeTailFill(categoryIndex, out RgbColor fourteenthTailFill))
            {
                return fourteenthTailFill;
            }
        }
        if (UseThirteenthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveThirteenthRegimeDarkFill(categoryIndex, out RgbColor thirteenthDarkFill))
            {
                return thirteenthDarkFill;
            }

            if (TryResolveThirteenthRegimeSecondFill(categoryIndex, out RgbColor thirteenthSecondFill))
            {
                return thirteenthSecondFill;
            }

            if (categoryIndex < 18)
            {
                return ShadeThirteenthRegimeThirdRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeThirteenthRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeThirteenthRegimeSixthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 42)
            {
                return ShadeThirteenthRegimeSeventhRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveThirteenthRegimeFixed43Fill(categoryIndex, out RgbColor thirteenthFixed43Fill))
            {
                return thirteenthFixed43Fill;
            }

            if (TryResolveThirteenthRegimeEighthFill(categoryIndex, out RgbColor thirteenthEighthFill))
            {
                return thirteenthEighthFill;
            }

            if (TryResolveThirteenthRegimeSeventhFill(categoryIndex, out RgbColor thirteenthSeventhFill))
            {
                return thirteenthSeventhFill;
            }

            if (TryResolveThirteenthRegimeFixed61Fill(categoryIndex, out RgbColor thirteenthFixed61Fill))
            {
                return thirteenthFixed61Fill;
            }

            if (TryResolveThirteenthRegimePalestFill(categoryIndex, out RgbColor thirteenthPalestFill))
            {
                return thirteenthPalestFill;
            }

            if (TryResolveThirteenthRegimeFixed73Fill(categoryIndex, out RgbColor thirteenthFixed73Fill))
            {
                return thirteenthFixed73Fill;
            }

            if (TryResolveThirteenthRegimeTailFill(categoryIndex, out RgbColor thirteenthTailFill))
            {
                return thirteenthTailFill;
            }
        }
        if (UseTwelfthVaryColorsRegime(valuePointCount))
        {
            if (TryResolveTwelfthRegimeDarkFill(categoryIndex, out RgbColor twelfthDarkFill))
            {
                return twelfthDarkFill;
            }

            if (categoryIndex < 12)
            {
                return ShadeTwelfthRegimeSecondRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 18)
            {
                return ShadeTwelfthRegimeThirdRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeFifthRegimeMidSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeNinthRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeFifthRegimeLightSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 42)
            {
                return paletteColor;
            }

            if (TryResolveTwelfthRegimeFifth19Fill(categoryIndex, out RgbColor twelfthFifth19Fill))
            {
                return twelfthFifth19Fill;
            }

            if (TryResolveTwelfthRegimeFixed49Fill(categoryIndex, out RgbColor twelfthFixed49Fill))
            {
                return twelfthFixed49Fill;
            }

            if (TryResolveTwelfthRegimeFifth25Fill(categoryIndex, out RgbColor twelfthFifth25Fill))
            {
                return twelfthFifth25Fill;
            }

            if (TryResolveTwelfthRegimeFixed61Fill(categoryIndex, out RgbColor twelfthFixed61Fill))
            {
                return twelfthFixed61Fill;
            }

            if (TryResolveTwelfthRegimeOverflow31Fill(categoryIndex, out RgbColor twelfthOverflow31Fill))
            {
                return twelfthOverflow31Fill;
            }

            if (TryResolveTwelfthRegimeTailFill(categoryIndex, out RgbColor twelfthTailFill))
            {
                return twelfthTailFill;
            }
        }
        if (UseEleventhVaryColorsRegime(valuePointCount))
        {
            if (TryResolveEleventhRegimeDarkFill(categoryIndex, out RgbColor eleventhDarkFill))
            {
                return eleventhDarkFill;
            }

            if (TryResolveEleventhRegimeSecondFill(categoryIndex, out RgbColor eleventhSecondFill))
            {
                return eleventhSecondFill;
            }

            if (categoryIndex < 18)
            {
                return ShadeEleventhRegimeThirdRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 24)
            {
                return ShadeEleventhRegimeFourthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 30)
            {
                return ShadeEleventhRegimeFifthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (categoryIndex < 36)
            {
                return ShadeEleventhRegimeSixthRowSingleSeriesVaryColorsFill(paletteColor);
            }

            if (TryResolveEleventhRegimeFixed37Fill(categoryIndex, out RgbColor eleventhFixed37Fill))
            {
                return eleventhFixed37Fill;
            }

            if (TryResolveEleventhRegimeFixed43Fill(categoryIndex, out RgbColor eleventhFixed43Fill))
            {
                return eleventhFixed43Fill;
            }

            if (TryResolveEleventhRegimeFixed49Fill(categoryIndex, out RgbColor eleventhFixed49Fill))
            {
                return eleventhFixed49Fill;
            }

            if (TryResolveEleventhRegimeFixed55Fill(categoryIndex, out RgbColor eleventhFixed55Fill))
            {
                return eleventhFixed55Fill;
            }

            if (TryResolveEleventhRegimeFixed61Fill(categoryIndex, out RgbColor eleventhFixed61Fill))
            {
                return eleventhFixed61Fill;
            }

            if (TryResolveEleventhRegimeTailFill(categoryIndex, out RgbColor eleventhTailFill))
            {
                return eleventhTailFill;
            }
        }
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
