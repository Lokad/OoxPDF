using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxTextEmissionPlan(
    double PdfFontSize,
    double PdfCharacterSpacing,
    double PositioningCharacterSpacing,
    bool CompensatePdfCharacterSpacing);

internal static class DocxTextEmissionPlanner
{
    public static DocxTextEmissionPlan Create(
        DocxTextRun style,
        double layoutFontSize,
        double pdfCharacterSpacing,
        bool compensatePdfCharacterSpacing)
    {
        double positioningCharacterSpacing = compensatePdfCharacterSpacing
            ? style.CharacterSpacingPoints - pdfCharacterSpacing
            : style.CharacterSpacingPoints;
        return new(
            OfficePdfTextEmissionProfile.FontSize(layoutFontSize),
            pdfCharacterSpacing,
            positioningCharacterSpacing,
            compensatePdfCharacterSpacing);
    }

    public static DocxTextEmissionPlan CreateTerminalLineSpace(DocxTextRun style, double layoutFontSize)
    {
        return Create(style, layoutFontSize, pdfCharacterSpacing: 0d, compensatePdfCharacterSpacing: true);
    }
}
