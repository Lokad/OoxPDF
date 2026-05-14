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

    public static double HalfPointsToPoints(int halfPoints)
    {
        return halfPoints / 2d;
    }
}
