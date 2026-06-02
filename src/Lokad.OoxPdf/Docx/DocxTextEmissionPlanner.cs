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

internal readonly record struct DocxTextEmissionAdvanceProfile(
    int GlyphCount,
    int GlyphGapCount,
    double NaturalPdfWidth,
    double LayoutWidth,
    double LayoutToNaturalResidual,
    double? UniformResidualPerGap);

internal readonly record struct DocxTextEmissionGlyphAdvanceSignature(
    int GlyphCount,
    int GlyphPairCount,
    int AdvanceUnits,
    int KerningUnits,
    int PairAdvanceUnits,
    int PairAdvanceMinUnits,
    int PairAdvanceMaxUnits,
    string PairHash,
    string Hash);

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

    public static DocxTextEmissionAdvanceProfile MeasureAdvanceProfile(
        string text,
        PdfEmbeddedFont embedded,
        double layoutWidth,
        DocxTextEmissionPlan plan)
    {
        int glyphCount = CountMappedGlyphs(text, embedded);
        int glyphGapCount = Math.Max(0, glyphCount - 1);
        double naturalPdfWidth = embedded.MeasureTextPoints(text, plan.PdfFontSize, kerningEnabled: true);
        double residual = layoutWidth - naturalPdfWidth;
        double? residualPerGap = glyphGapCount == 0 ? null : residual / glyphGapCount;
        return new(glyphCount, glyphGapCount, naturalPdfWidth, layoutWidth, residual, residualPerGap);
    }

    public static DocxTextEmissionGlyphAdvanceSignature CreateGlyphAdvanceSignature(string text, PdfEmbeddedFont embedded)
    {
        const ulong fnvOffset = 14695981039346656037UL;
        const ulong fnvPrime = 1099511628211UL;

        ulong hash = fnvOffset;
        int glyphCount = 0;
        int glyphPairCount = 0;
        int advanceUnits = 0;
        int kerningUnits = 0;
        int pairAdvanceUnits = 0;
        int pairAdvanceMinUnits = 0;
        int pairAdvanceMaxUnits = 0;
        ulong pairHash = fnvOffset;
        ushort previousGlyph = 0;
        ushort previousAdvance = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = embedded.Font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            ushort advance = embedded.Font.GetAdvanceWidth(glyph);
            short kerning = previousGlyph == 0 ? (short)0 : embedded.Font.GetKerning(previousGlyph, glyph);
            if (previousGlyph != 0)
            {
                glyphPairCount++;
                kerningUnits += kerning;
                int pairAdvance = previousAdvance + advance + kerning;
                pairAdvanceUnits += pairAdvance;
                pairAdvanceMinUnits = glyphPairCount == 1 ? pairAdvance : Math.Min(pairAdvanceMinUnits, pairAdvance);
                pairAdvanceMaxUnits = glyphPairCount == 1 ? pairAdvance : Math.Max(pairAdvanceMaxUnits, pairAdvance);
                pairHash = AppendHash(pairHash, previousGlyph, fnvPrime);
                pairHash = AppendHash(pairHash, glyph, fnvPrime);
                pairHash = AppendHash(pairHash, unchecked((ushort)pairAdvance), fnvPrime);
            }

            advanceUnits += advance;
            hash = AppendHash(hash, glyph, fnvPrime);
            hash = AppendHash(hash, advance, fnvPrime);
            hash = AppendHash(hash, unchecked((ushort)kerning), fnvPrime);
            previousGlyph = glyph;
            previousAdvance = advance;
            glyphCount++;
        }

        return new(
            glyphCount,
            glyphPairCount,
            advanceUnits,
            kerningUnits,
            pairAdvanceUnits,
            pairAdvanceMinUnits,
            pairAdvanceMaxUnits,
            pairHash.ToString("X16", CultureInfo.InvariantCulture),
            hash.ToString("X16", CultureInfo.InvariantCulture));
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

    private static int CountMappedGlyphs(string text, PdfEmbeddedFont embedded)
    {
        int count = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (embedded.Font.MapCodePoint(rune.Value) != 0)
            {
                count++;
            }
        }

        return count;
    }

    private static ulong AppendHash(ulong hash, ushort value, ulong fnvPrime)
    {
        hash ^= (byte)(value & 0xFF);
        hash *= fnvPrime;
        hash ^= (byte)(value >> 8);
        hash *= fnvPrime;
        return hash;
    }
}
