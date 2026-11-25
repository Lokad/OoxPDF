using System.Globalization;

namespace Lokad.OoxPdf.Ooxml;

internal static class OoxUnits
{
    public const double EmusPerInch = 914400d;
    public const double TwipsPerPoint = 20d;
    public const double PointsPerInch = 72d;

    public static double EmuToPoints(long emu)
    {
        return emu / EmusPerInch * PointsPerInch;
    }

    public static double TwipsToPoints(long twips)
    {
        return twips / TwipsPerPoint;
    }

    // A bare number is already in the field's units. A recognized suffix must
    // carry a valid physical measure before it can be converted to those units.
    public static long? ReadPhysicalMeasure(ReadOnlySpan<char> measure, decimal unitsPerPoint)
    {
        if (measure.Length < 3)
            return null;

        decimal? pointsPerUnit = measure[^2..] switch
        {
            "pt" => 1m,
            "in" => 72m,
            "cm" => 72m / 2.54m,
            "mm" => 72m / 25.4m,
            "pc" or "pi" => 12m,
            _ => null
        };
        return pointsPerUnit is { } factor
            ? ReadScaledMeasure(measure[..^2], factor * unitsPerPoint) : null;
    }

    public static long ReadScaledMeasure(ReadOnlySpan<char> number, decimal factor)
    {
        if (!decimal.TryParse(number, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out decimal value))
            throw new InvalidDataException("Malformed OOXML measure.");

        try
        {
            return checked((long)decimal.Round(value * factor, 0, MidpointRounding.AwayFromZero));
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("OOXML measure is outside the supported range.", ex);
        }
    }
}
