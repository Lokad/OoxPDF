namespace Lokad.OoxPdf.Pptx;

internal static class LinearLightColor
{
    // Single owner for the Office-calibrated linear-light color math (sRGB to linear,
    // float, round-half-away bytes). Shared by the chart fraction-curve engine and the
    // table band fills; previously byte-identical private copies lived in both files.
    public static double ToLinear(byte channel)
    {
        double c = channel / 255d;
        return c <= 0.04045d ? c / 12.92d : System.Math.Pow((c + 0.055d) / 1.055d, 2.4d);
    }

    public static byte ToSrgbByte(double linear)
    {
        double clamped = System.Math.Clamp(linear, 0d, 1d);
        double c = clamped <= 0.0031308d ? clamped * 12.92d : 1.055d * System.Math.Pow(clamped, 1d / 2.4d) - 0.055d;
        return (byte)System.Math.Clamp((int)System.Math.Round(c * 255d, System.MidpointRounding.AwayFromZero), 0, 255);
    }

    // Linear-light tint toward white (table bands; fraction-curve upper-half form).
    public static RgbColor TintTowardWhite(RgbColor color, double weight)
    {
        return new RgbColor(
            ToSrgbByte(ToLinear(color.Red) + (1d - ToLinear(color.Red)) * weight),
            ToSrgbByte(ToLinear(color.Green) + (1d - ToLinear(color.Green)) * weight),
            ToSrgbByte(ToLinear(color.Blue) + (1d - ToLinear(color.Blue)) * weight));
    }

    // Linear-light shade toward black (table dark bands; fraction-curve base form).
    public static RgbColor ShadeTowardBlack(RgbColor color, double weight)
    {
        return new RgbColor(
            ToSrgbByte(ToLinear(color.Red) * weight),
            ToSrgbByte(ToLinear(color.Green) * weight),
            ToSrgbByte(ToLinear(color.Blue) * weight));
    }
}
