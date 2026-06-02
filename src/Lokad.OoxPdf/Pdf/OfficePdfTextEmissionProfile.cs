namespace Lokad.OoxPdf.Pdf;

internal static class OfficePdfTextEmissionProfile
{
    private const double OfficeExportFontGridDpi = 600d;
    private const double PointsPerInch = 72d;
    private const double WordNumberedListTextStateCharacterSpacingEm = 0.004d;

    public static double FontSize(double layoutFontSize)
    {
        double deviceUnits = Math.Round(
            layoutFontSize * OfficeExportFontGridDpi / PointsPerInch,
            MidpointRounding.AwayFromZero);

        return deviceUnits * PointsPerInch / OfficeExportFontGridDpi;
    }

    public static double WordNumberedListTextStateCharacterSpacing(double layoutFontSize)
    {
        return layoutFontSize * WordNumberedListTextStateCharacterSpacingEm;
    }
}
