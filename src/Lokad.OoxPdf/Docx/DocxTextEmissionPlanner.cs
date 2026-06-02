using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxTextEmissionPlan(
    double PdfFontSize,
    double PdfCharacterSpacing,
    double PositioningCharacterSpacing,
    bool CompensatePdfCharacterSpacing);

internal readonly record struct DocxTextEmissionPart(string Text, double X, double Width);

internal readonly record struct DocxTextEmissionCharacterProfile(
    int DigitCount,
    int LetterCount,
    int WhitespaceCount,
    int PunctuationCount,
    int SymbolCount,
    int OtherCount);

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

    public static DocxTextEmissionCharacterProfile ClassifyText(string text)
    {
        int digitCount = 0;
        int letterCount = 0;
        int whitespaceCount = 0;
        int punctuationCount = 0;
        int symbolCount = 0;
        int otherCount = 0;

        foreach (Rune rune in text.EnumerateRunes())
        {
            if (Rune.IsWhiteSpace(rune))
            {
                whitespaceCount++;
                continue;
            }

            switch (Rune.GetUnicodeCategory(rune))
            {
                case UnicodeCategory.DecimalDigitNumber:
                    digitCount++;
                    break;
                case UnicodeCategory.UppercaseLetter:
                case UnicodeCategory.LowercaseLetter:
                case UnicodeCategory.TitlecaseLetter:
                case UnicodeCategory.ModifierLetter:
                case UnicodeCategory.OtherLetter:
                    letterCount++;
                    break;
                case UnicodeCategory.ConnectorPunctuation:
                case UnicodeCategory.DashPunctuation:
                case UnicodeCategory.OpenPunctuation:
                case UnicodeCategory.ClosePunctuation:
                case UnicodeCategory.InitialQuotePunctuation:
                case UnicodeCategory.FinalQuotePunctuation:
                case UnicodeCategory.OtherPunctuation:
                    punctuationCount++;
                    break;
                case UnicodeCategory.MathSymbol:
                case UnicodeCategory.CurrencySymbol:
                case UnicodeCategory.ModifierSymbol:
                case UnicodeCategory.OtherSymbol:
                    symbolCount++;
                    break;
                default:
                    otherCount++;
                    break;
            }
        }

        return new(digitCount, letterCount, whitespaceCount, punctuationCount, symbolCount, otherCount);
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
