namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Fraction-curve varyColors engine v3 (linear-light recipe).
    // Validated range: 96-plus points, single-series vertical bar/column varyColors fills.
    // accent = categoryIndex % 6, level = categoryIndex / 6, cap = valuePointCount / 6,
    // fraction = level / cap. Base shade k per turn band; at/below fraction 0.5 lerp
    // the linear shade toward raw, above fraction 0.5 tint toward white with weight 1.28.
    // Bases resolve live from the chart palette/theme, so custom themes work.
    // End-to-end vs Office: worst 4 (see PLAN.md).
    private const int FractionCurveVaryColorsPointThreshold = 96;

    private const double FractionCurveTintWeight = 1.28d;

    // Linear-light base shade factor per turn band (Office-measured, accent-uniform).
    private static double FractionCurveShadeK(int valuePointCount)
    {
        if (valuePointCount < 114)
        {
            return 0.3686d;
        }

        if (valuePointCount < 132)
        {
            return 0.3592d;
        }

        if (valuePointCount < 162)
        {
            return 0.3511d;
        }

        if (valuePointCount < 204)
        {
            return 0.3414d;
        }

        return 0.3298d;
    }

    private static bool TryResolveFractionCurveVaryColorsFill(
        IReadOnlyList<RgbColor>? chartPalette,
        PptxTheme? theme,
        PptxColorMap colorMap,
        int categoryIndex,
        int valuePointCount,
        out RgbColor fill)
    {
        fill = default;
        if (valuePointCount < FractionCurveVaryColorsPointThreshold)
        {
            return false;
        }

        int accent = categoryIndex % 6;
        int level = categoryIndex / 6;
        int cap = valuePointCount / 6;
        if (cap <= 0)
        {
            return false;
        }

        RgbColor accentBase = ChartPalette(chartPalette, theme, colorMap, accent);
        double shadeK = FractionCurveShadeK(valuePointCount);
        double fraction = (double)level / cap;
        double accentRed = SrgbToLinear(accentBase.Red);
        double accentGreen = SrgbToLinear(accentBase.Green);
        double accentBlue = SrgbToLinear(accentBase.Blue);
        double red, green, blue;
        if (fraction <= 0.5d)
        {
            double t = fraction / 0.5d;
            red = accentRed * shadeK + (accentRed - accentRed * shadeK) * t;
            green = accentGreen * shadeK + (accentGreen - accentGreen * shadeK) * t;
            blue = accentBlue * shadeK + (accentBlue - accentBlue * shadeK) * t;
        }
        else
        {
            double w = (fraction - 0.5d) * FractionCurveTintWeight;
            red = accentRed + (1d - accentRed) * w;
            green = accentGreen + (1d - accentGreen) * w;
            blue = accentBlue + (1d - accentBlue) * w;
        }

        fill = new RgbColor(LinearToSrgbByte(red), LinearToSrgbByte(green), LinearToSrgbByte(blue));
        return true;
    }

    private static double SrgbToLinear(byte channel)
    {
        double c = channel / 255d;
        return c <= 0.04045d ? c / 12.92d : System.Math.Pow((c + 0.055d) / 1.055d, 2.4d);
    }

    private static byte LinearToSrgbByte(double linear)
    {
        double clamped = System.Math.Clamp(linear, 0d, 1d);
        double c = clamped <= 0.0031308d ? clamped * 12.92d : 1.055d * System.Math.Pow(clamped, 1d / 2.4d) - 0.055d;
        return (byte)System.Math.Clamp((int)System.Math.Round(c * 255d, System.MidpointRounding.AwayFromZero), 0, 255);
    }
}
