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
    private const double WordExactLineFirstBaselineRatio = 0.8d;

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
        // RV06 pagination probe (edge-page-ex48-body, Word 16.0): Office drops the
        // first baseline of exact-spaced body text to 0.8 x the exact line height,
        // the same ratio as in-cell text; the 0.299em bottom inset sat 6pt too deep.
        return hasExplicitLineSpacing
            ? Math.Max(0d, lineHeight * WordExactLineFirstBaselineRatio)
            : fontSize * WordAutoLineBaselineOffsetEm;
    }

    public static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        DocxParagraph? firstTextParagraph = paragraphs.FirstOrDefault(paragraph => paragraph.Runs.Count != 0);
        if (firstTextParagraph is null)
        {
            return 0d;
        }

        // RV06 pagination probe (edge-page-ex24/ex36/ex48, Word 16.0): Office drops
        // the first baseline of exact-spaced in-cell text to 0.8 x the exact line
        // height below the content top (19.2/28.8/38.4), independent of the font
        // ascent the auto rule applies. Font-size interaction past 12pt stays a
        // probe follow-up; atLeast keeps the auto rule for lack of Office evidence.
        DocxEffectiveParagraphProperties effective = firstTextParagraph.EffectiveProperties;
        if (effective.LineSpacingPoints is { } exactLineHeight &&
            !string.Equals(effective.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))
        {
            return exactLineHeight * WordExactLineFirstBaselineRatio;
        }

        // Word places the in-cell first baseline with the body rule (shading probes 2026-09-06: in-cell offsets match winAscent like body text); the full-em inset sat 0.05em too deep.
        return firstTextParagraph.Runs.Max(run => run.EffectiveProperties.FontSize) * WordAutoLineBaselineOffsetEm;
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
        return run.EffectiveProperties.VerticalAlignment == DocxRunVerticalAlignment.Superscript;
    }

    private static bool IsSubscript(DocxTextRun run)
    {
        return run.EffectiveProperties.VerticalAlignment == DocxRunVerticalAlignment.Subscript;
    }
}

// RV01: deterministic fallback measurer for text with no usable embeddable
// face. Average advances (wide scripts count double, combining marks zero)
// plus Helvetica-fraction line metrics; character spacing follows the shared
// helper so layout and fallback emission agree exactly.
internal sealed class DocxFallbackTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        double units = 0d;
        foreach (Rune rune in text.EnumerateRunes())
        {
            units += PdfFallbackFont.MeasureAdvanceEm(rune);
        }

        return DocxTextSpacing.AddCharacterSpacing(units * fontSize / PdfFallbackFont.UnitsPerEm, run, text);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        return fontSize * PdfFallbackFont.SingleLineHeightEm;
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        return fontSize * PdfFallbackFont.AscentEm;
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        return fontSize * PdfFallbackFont.DescentEm;
    }
}

// RV01: routes runs without a usable embedded resource to the fallback
// measurer while the remaining runs keep their font measurement, so mixed
// documents measure each run with the face that emits it.
internal sealed class MissingFontRoutingMeasurer(IDocxTextMeasurer? inner, DocxFallbackTextMeasurer fallback, IReadOnlyDictionary<DocxTextRun, PdfFallbackFontResource> fallbackFaces) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        if (run is not null && fallbackFaces.ContainsKey(run))
        {
            return fallback.MeasureText(run, text, fontSize);
        }

        return inner is not null
            ? inner.MeasureText(run, text, fontSize)
            : fallback.MeasureText(run, text, fontSize);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        if (run is not null && fallbackFaces.ContainsKey(run))
        {
            return fallback.MeasureSingleLineHeight(run, fontSize);
        }

        return inner is IDocxLineMetricsProvider provider
            ? provider.MeasureSingleLineHeight(run, fontSize)
            : fallback.MeasureSingleLineHeight(run, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        if (run is not null && fallbackFaces.ContainsKey(run))
        {
            return fallback.MeasureWindowsAscender(run, fontSize);
        }

        return inner is IDocxStaticTextMetricsProvider provider
            ? provider.MeasureWindowsAscender(run, fontSize)
            : fallback.MeasureWindowsAscender(run, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        if (run is not null && fallbackFaces.ContainsKey(run))
        {
            return fallback.MeasureWindowsDescender(run, fontSize);
        }

        return inner is IDocxStaticTextMetricsProvider provider
            ? provider.MeasureWindowsDescender(run, fontSize)
            : fallback.MeasureWindowsDescender(run, fontSize);
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
