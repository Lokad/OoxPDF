using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxTextEmissionPlan(
    double PdfFontSize,
    double PdfCharacterSpacing,
    double PositioningCharacterSpacing,
    bool CompensatePdfCharacterSpacing);

internal readonly record struct DocxTextEmissionPart(string Text, double X, double Width);

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

    public static IReadOnlyList<DocxTextEmissionPart> SplitOfficeTextOperationParts(
        DocxTextSegmentLayout segment,
        double fontSize,
        IDocxTextMeasurer? textMeasurer)
    {
        if (textMeasurer is null ||
            segment.Text.Length == 0 ||
            !segment.Text.Any(IsOfficeTextOperationBoundaryPunctuation))
        {
            return [new DocxTextEmissionPart(segment.Text, segment.X, segment.Width)];
        }

        var parts = new List<DocxTextEmissionPart>();
        int partStart = 0;
        for (int i = 0; i < segment.Text.Length; i++)
        {
            if (!IsOfficeTextOperationBoundaryPunctuation(segment.Text[i]))
            {
                continue;
            }

            AddTextEmissionPart(segment, fontSize, textMeasurer, partStart, i - partStart, parts);
            AddTextEmissionPart(segment, fontSize, textMeasurer, i, 1, parts);
            partStart = i + 1;
        }

        AddTextEmissionPart(segment, fontSize, textMeasurer, partStart, segment.Text.Length - partStart, parts);
        return parts.Count == 0 ? [new DocxTextEmissionPart(segment.Text, segment.X, segment.Width)] : parts;
    }

    private static void AddTextEmissionPart(
        DocxTextSegmentLayout segment,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        int start,
        int length,
        List<DocxTextEmissionPart> parts)
    {
        if (length <= 0)
        {
            return;
        }

        string prefix = start == 0 ? string.Empty : segment.Text[..start];
        string text = segment.Text.Substring(start, length);
        double x = segment.X + textMeasurer.MeasureText(segment.StyleRun, prefix, fontSize);
        double width = start + length == segment.Text.Length
            ? Math.Max(0d, segment.X + segment.Width - x)
            : textMeasurer.MeasureText(segment.StyleRun, text, fontSize);
        parts.Add(new DocxTextEmissionPart(text, x, width));
    }

    private static bool IsOfficeTextOperationBoundaryPunctuation(char value)
    {
        return CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.DashPunctuation;
    }
}
