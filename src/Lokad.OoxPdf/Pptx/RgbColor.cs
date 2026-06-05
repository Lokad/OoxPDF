using System.Globalization;

namespace Lokad.OoxPdf.Pptx;

internal readonly record struct RgbColor(byte Red, byte Green, byte Blue)
{
    public static bool TryParse(string? hex, out RgbColor color)
    {
        if (hex is null || hex.Length != 6)
        {
            color = default;
            return false;
        }

        color = new RgbColor(
            byte.Parse(hex.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(hex.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(hex.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        return true;
    }
}
