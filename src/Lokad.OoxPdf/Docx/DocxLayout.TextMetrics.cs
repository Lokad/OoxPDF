using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

// Text measurement contracts and shared metrics. Split from the layout file:
// measurer interfaces, spacing/line/vertical-align metrics, and the
// embedded-font measurer used for layout and static stories.

internal interface IDocxTextMeasurer
{
    double MeasureText(DocxTextRun? run, string text, double fontSize);
}

internal interface IDocxLineMetricsProvider
{
    double MeasureSingleLineHeight(DocxTextRun? run, double fontSize);
}

internal interface IDocxStaticTextMetricsProvider
{
    double MeasureWindowsAscender(DocxTextRun? run, double fontSize);

    double MeasureWindowsDescender(DocxTextRun? run, double fontSize);
}

internal static class DocxTextSpacing
{
    public static double AddCharacterSpacing(double measuredWidth, DocxTextRun? run, string text)
    {
        return measuredWidth + CountCharacterSpacingGaps() * (run?.EffectiveProperties.CharacterSpacingPoints ?? 0d);

        int CountCharacterSpacingGaps()
        {
            int count = 0;
            foreach (Rune _ in text.EnumerateRunes())
            {
                count++;
            }

            return Math.Max(0, count - 1);
        }
    }

    public static double BoundarySpacing(DocxTextRun? left, string leftText, string rightText)
    {
        return leftText.Length == 0 ||
            rightText.Length == 0 ||
            leftText[^1] == '\t' ||
            rightText[0] == '\t'
            ? 0d
            : left?.EffectiveProperties.CharacterSpacingPoints ?? 0d;
    }

}

internal static class DocxLineMetrics
{
    private const double WordSingleLineMinimumEm = 1.15d;
    private const double WordAutoLineBaselineOffsetEm = 0.94d;
    private const double WordExactLineTextBottomInsetEm = 0.299d;

    public static double MeasureOpenTypeSingleLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize * WordSingleLineMinimumEm;
        }

        double units = font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;

        return Math.Max(fontSize * WordSingleLineMinimumEm, units * fontSize / font.UnitsPerEm);
    }

    public static double MeasureWindowsAscender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? fontSize
            : font.Os2.WindowsAscender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureWindowsDescender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Os2.WindowsDescender * fontSize / font.UnitsPerEm;
    }

    public static double ResolveBodyBaselineOffset(double fontSize, double lineHeight, bool hasExplicitLineSpacing)
    {
        return hasExplicitLineSpacing
            ? Math.Max(0d, lineHeight - fontSize * WordExactLineTextBottomInsetEm)
            : fontSize * WordAutoLineBaselineOffsetEm;
    }

    public static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        DocxParagraph? firstTextParagraph = paragraphs.FirstOrDefault(paragraph => paragraph.Runs.Count != 0);
        // Word places the in-cell first baseline with the body rule (shading probes 2026-09-06: in-cell offsets match winAscent like body text); the full-em inset sat 0.05em too deep.
        return firstTextParagraph is null ? 0d : firstTextParagraph.Runs.Max(run => run.EffectiveProperties.FontSize) * WordAutoLineBaselineOffsetEm;
    }
}

internal static class DocxVerticalAlignMetrics
{
    private const double SuperscriptSubscriptScale = 2d / 3d;
    private const double HalfPointGrid = 2d;
    private const double SubscriptBaselineShiftEm = -0.06d;

    public static double ResolveScriptBaseSize(DocxTextRun run, double nominalFontSize)
    {
        return run.ScriptBaseFontSize ?? nominalFontSize;
    }

    public static double ResolveFontSize(double nominalFontSize, DocxTextRun run)
    {
        if (!IsSuperscript(run) && !IsSubscript(run))
        {
            return nominalFontSize;
        }

        double baseSize = ResolveScriptBaseSize(run, nominalFontSize);
        return Math.Max(0.5d, Math.Floor(baseSize * SuperscriptSubscriptScale * HalfPointGrid) / HalfPointGrid);
    }

    public static double ResolveBaselineOffset(double nominalFontSize, double layoutFontSize, DocxTextRun run)
    {
        if (IsSuperscript(run))
        {
            double baseSize = ResolveScriptBaseSize(run, nominalFontSize);
            return Math.Max(0d, baseSize - ResolveFontSize(baseSize, run));
        }

        return IsSubscript(run) ? nominalFontSize * SubscriptBaselineShiftEm : 0d;
    }

    private static bool IsSuperscript(DocxTextRun run)
    {
        return run.EffectiveProperties.VerticalAlignmentValue?.Equals("superscript", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsSubscript(DocxTextRun run)
    {
        return run.EffectiveProperties.VerticalAlignmentValue?.Equals("subscript", StringComparison.OrdinalIgnoreCase) == true;
    }
}

internal sealed class DocxEmbeddedTextMeasurer(PdfEmbeddedFont embedded) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        return DocxTextSpacing.AddCharacterSpacing(embedded.MeasureTextPoints(text, fontSize), run, text);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureOpenTypeSingleLineHeight(embedded.Font, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsAscender(embedded.Font, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsDescender(embedded.Font, fontSize);
    }
}
