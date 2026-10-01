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

    // RV06 separator probes (Word 16.0): footnote and endnote separator rules follow
    // OS/2 strikeout geometry, with the rule top at the strikeout position and the rule
    // thickness at the strikeout size. Providers without font metrics keep the default,
    // so layout falls back to the legacy separator constants.
    bool TryGetStrikeoutRuleMetrics(DocxTextRun? run, out double positionEm, out double thicknessEm)
    {
        positionEm = 0d;
        thicknessEm = 0d;
        return false;
    }

    // RV06 separator-bottom probes (Word 16.0): the Office footnote gap above the
    // body equals one single-spaced line box minus the first-baseline inset, so it
    // derives from the mark font. Providers without font metrics keep the default,
    // so layout falls back to the legacy gap constant.
    bool TryGetSingleLineEm(DocxTextRun? run, out double singleLineEm)
    {
        singleLineEm = 0d;
        return false;
    }
}

internal interface IDocxLineMetricsProvider
{
    double MeasureSingleLineHeight(DocxTextRun? run, double fontSize);

    // Horizontal-header ascender for auto first-baseline insets (Word 16.0
    // body-grid probes). Providers without hhea metrics return zero so the
    // legacy inset applies and test doubles stay byte-identical.
    double MeasureHheaAscender(DocxTextRun? run, double fontSize)
    {
        return 0d;
    }

    // Horizontal-header line height for auto line boxes (Word 16.0 body-grid
    // probes). Providers without hhea metrics keep the single-line height, so
    // test doubles and fallbacks stay byte-identical.
    double MeasureHheaLineHeight(DocxTextRun? run, double fontSize)
    {
        return MeasureSingleLineHeight(run, fontSize);
    }
}

internal interface IDocxHheaLineGapProvider
{
    double MeasureHheaLineGap(DocxTextRun? run, double fontSize);
}

internal interface IDocxHheaDescenderProvider
{
    double MeasureHheaDescender(DocxTextRun? run, double fontSize);
}

internal interface IDocxStaticTextMetricsProvider
{
    double MeasureWindowsAscender(DocxTextRun? run, double fontSize);

    double MeasureWindowsDescender(DocxTextRun? run, double fontSize);
}

// RV06 Aptos probes (Word 16.0, uniform Aptos first baseline 0.61 high with the
// Windows-ascender supplement standing while Segoe uniform keeps it): only Aptos sets
// OS/2 USE_TYPO_METRICS among probed families, so runs whose resolved face requests
// typographic metrics skip the Windows-ascender content-gap supplement. Opt-in per
// provider; measurers without the interface keep legacy behavior byte-identically.
internal interface IDocxTypographicMetricsProvider
{
    bool UseTypographicMetrics(DocxTextRun? run);

    // Unfloored typographic line box for auto line boxes (Word 16.0 Abadi probes): runs whose resolved face requests typographic metrics box the typo box alone. Providers without typo metrics report zero so legacy maxima apply bit-identically.
    double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize)
    {
        return 0d;
    }
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
    internal const double WordAutoLineBaselineOffsetEm = 0.94d;
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

    public static double MeasureTypographicLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize;
        }

        double units = font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;
        return units * fontSize / font.UnitsPerEm;
    }

    public static double MeasureHheaLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize;
        }

        double units = font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap;
        return units * fontSize / font.UnitsPerEm;
    }

    public static double MeasureWindowsAscender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? fontSize
            : font.Os2.WindowsAscender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureHheaLineGap(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Hhea.HorizontalLineGap * fontSize / font.UnitsPerEm;
    }

    public static double MeasureHheaDescender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : -font.Hhea.HorizontalDescender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureHheaAscender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Hhea.HorizontalAscender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureWindowsDescender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Os2.WindowsDescender * fontSize / font.UnitsPerEm;
    }

    public static double ResolveBodyBaselineOffset(double fontSize, double lineHeight, bool hasExplicitLineSpacing, double? hheaAscenderPoints = null, double? bodyTierAMaxPoints = null)
    {
        // RV06 pagination probe (edge-page-ex48-body, Word 16.0): Office drops the
        // first baseline of exact-spaced body text to 0.8 x the exact line height,
        // the same ratio as in-cell text; the 0.299em bottom inset sat 6pt too deep.
        // Word 16.0 body-grid probes (Tahoma/Verdana/Segoe UI first baselines drift
        // with hheaAsc while Times/Aptos hold the 0.94em rule): auto insets take
        // max(hheaAscender, 0.94em). Opt-in per call site; null keeps legacy.
        return hasExplicitLineSpacing
            ? Math.Max(0d, lineHeight * WordExactLineFirstBaselineRatio)
            : Math.Max(Math.Max(fontSize * WordAutoLineBaselineOffsetEm, hheaAscenderPoints ?? 0d), bodyTierAMaxPoints ?? 0d);
    }

    // RV06 cell first-baseline box rule (Word 16.0, Harlow/Arial/Calibri size matrix at 11 and 12pt plus Palatino/Harlow mixed orders plus Tahoma/BookAntiqua/Bauhaus uniforms, genuine embeds throughout): Office pins the first-line box top at the cell content top, so the first baseline sits one full line box above the box bottom. Per non-typographic run that is max(Windows box, hhea box plus hhea gap) minus max descender, maximized across runs with no 0.94em floor. Fit table, candidate minus Office: Harlow plus 0.00 and plus 0.09, Arial minus 0.15 and minus 0.05, Calibri plus 0.02 and plus 0.01, Palatino minus 0.03 both mixed orders order-blind at Palatino level, Tahoma plus 0.03 and minus 0.02, Book Antiqua minus 0.14, Bauhaus minus 0.05 with its 673-unit gap lifting the hhea box. Gap-inclusion alone is falsified by Calibri (minus 0.65); ascender-only max is falsified by Arial (minus 1.35) and Bauhaus (minus 4.00); the floor alone is falsified by Harlow (plus 0.70) and Palatino (minus 1.34). Mixed-size cells keep the legacy floor slope (Office 12->15 transition carries exactly 0.94em at 2.82), so the gate admits size-uniform cells only. Full-metric quartet required, so provider-less doubles and the embedded single-font measurer keep legacy behavior byte-identically.
    internal static bool TableCellFontSizesUniform(IReadOnlyList<DocxParagraph> cellParagraphs)
    {
        double? size = null;
        foreach (DocxParagraph paragraph in cellParagraphs)
        {
            foreach (DocxTextRun run in paragraph.Runs)
            {
                double runSize = run.EffectiveProperties.FontSize;
                if (size is null)
                {
                    size = runSize;
                }
                else if (runSize != size.Value)
                {
                    return false;
                }
            }
        }

        return true;
    }

    internal static double? ResolveTableCellFirstBaselinePoints(IReadOnlyList<DocxParagraph> cellParagraphs, DocxParagraph paragraph, double maxSize, IDocxTextMeasurer? measurer)
    {
        if (!TableCellFontSizesUniform(cellParagraphs) ||
            measurer is not IDocxStaticTextMetricsProvider cellStatic ||
            measurer is not IDocxLineMetricsProvider cellLines ||
            measurer is not IDocxHheaDescenderProvider cellDesc ||
            measurer is not IDocxHheaLineGapProvider cellGap ||
            paragraph.Runs.Count == 0)
        {
            return null;
        }

        double? firstBaseline = null;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            if (measurer is IDocxTypographicMetricsProvider cellTypographic && cellTypographic.UseTypographicMetrics(run))
            {
                continue;
            }

            double windowsBox = cellStatic.MeasureWindowsAscender(run, maxSize) + cellStatic.MeasureWindowsDescender(run, maxSize);
            double hheaBox = cellLines.MeasureHheaAscender(run, maxSize) + cellDesc.MeasureHheaDescender(run, maxSize) + cellGap.MeasureHheaLineGap(run, maxSize);
            double descender = Math.Max(cellStatic.MeasureWindowsDescender(run, maxSize), cellDesc.MeasureHheaDescender(run, maxSize));
            double runFirst = Math.Max(windowsBox, hheaBox) - descender;
            firstBaseline = Math.Max(firstBaseline ?? 0d, runFirst);
        }

        return firstBaseline;
    }

    // Cell-visible first-baseline offset: the table-cell inset cancels out of the first-line transition, so the visible baseline rides this term at every cell BodyOffset site (transition, midline and estimate plans, single-line path). Size-uniform cells with full-metric measurers take the unfloored box rule; exact spacing, mixed-size cells, flagged-only content and partial doubles keep the legacy body offset byte-identically.
    internal static double ResolveTableCellBodyBaselineOffset(double fontSize, double lineHeight, bool hasExplicitLineSpacing, IReadOnlyList<DocxParagraph> cellParagraphs, DocxParagraph paragraph, IDocxTextMeasurer? measurer)
    {
        if (!hasExplicitLineSpacing &&
            ResolveTableCellFirstBaselinePoints(cellParagraphs, paragraph, fontSize, measurer) is { } cellFirst)
        {
            return cellFirst;
        }

        return ResolveBodyBaselineOffset(fontSize, lineHeight, hasExplicitLineSpacing, ResolveHheaAscenderPoints(paragraph, fontSize, measurer));
    }

    // Related-story content (footnotes/endnotes) keeps order-blind min-hhea selection:
    // Office footnote first baselines stay invariant across same-size mixed runs in every
    // order (edge-fnmix edge-fnmix3: 85.46 everywhere), which neither max-hhea nor widest-run
    // tie-breaking reproduces. Mixed-size takes stay a separate probe.
    internal static double? ResolveHheaAscenderPoints(DocxParagraph paragraph, double fontSize, IDocxLineMetricsProvider? provider, bool selectMaxHhea = true)
    {
        if (provider is null || paragraph.Runs.Count == 0)
        {
            return null;
        }

        // RV05 bodymix probes (Word COM references edge-bodymix-cal/tah plus the
        // mixed-size edge-bodymix10-tah probe): Office sizes first-baseline insets through
        // the max-hhea run at its own size, not the widest run by size (a direct Tahoma
        // 12pt run shifts a Calibri body by 0.73pt, while a Tahoma 10pt run leaves it
        // unchanged at own-size 10.005 below the floor). Uniform documents resolve
        // identically; only mixed-font lines change.
        if (!selectMaxHhea)
        {
            // Related-story content keeps order-blind legacy behavior: Office footnote
            // first baselines stay invariant across same-size mixed runs in every order
            // (edge-fnmix3 tahfirst/calfirst: 85.46 both), which neither max-hhea nor
            // widest-run tie-breaking reproduces. Mixed-size takes stay a separate probe.
            double? minAscender = null;
            foreach (DocxTextRun run in paragraph.Runs)
            {
                double ascender = provider.MeasureHheaAscender(run, run.EffectiveProperties.FontSize);
                if (minAscender is null || ascender < minAscender.Value)
                {
                    minAscender = ascender;
                }
            }

            return minAscender;
        }

        double? maxAscender = null;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            double ascender = provider.MeasureHheaAscender(run, run.EffectiveProperties.FontSize);
            if (maxAscender is null || ascender > maxAscender.Value)
            {
                maxAscender = ascender;
            }
        }

        return maxAscender;
    }

    internal static double? ResolveHheaAscenderPoints(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? measurer, bool selectMaxHhea = true)
    {
        return ResolveHheaAscenderPoints(paragraph, fontSize, measurer as IDocxLineMetricsProvider, selectMaxHhea);
    }

    // RV06 big-gap body probes (Word COM references edge-endsepgrid-fort/palb/bauh/sitka with Calibri
    // bodies): Office body first-baseline insets follow the max tier-A (hhea box minus Windows
    // descender) across runs at own sizes, not max-hhea (Forte/Palatino/Bauhaus bodies sit 0.9-3.5
    // deeper while Magneto/Constantia and all small-gap families hold the validated levels).
    // Full-metric measurers only; the caller keeps the legacy floor so all other paths stay identical.
    internal static double? ResolveBodyTierAMaxPoints(DocxParagraph paragraph, IDocxTextMeasurer? measurer)
    {
        if (measurer is not IDocxLineMetricsProvider lineMetrics ||
            measurer is not IDocxHheaLineGapProvider gapMetrics ||
            measurer is not IDocxStaticTextMetricsProvider staticMetrics ||
            measurer is not IDocxHheaDescenderProvider ||
            paragraph.Runs.Count == 0)
        {
            return null;
        }

        double? maxTier = null;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            double size = run.EffectiveProperties.FontSize;
            // Gap-less faces keep the legacy max-hhea path: synthetic and zero-gap families resolve
            // identically there, so only faces carrying hhea line gap take the tier-A branch.
            if (gapMetrics.MeasureHheaLineGap(run, size) <= 0d)
            {
                continue;
            }
            double tier = lineMetrics.MeasureHheaLineHeight(run, size) - staticMetrics.MeasureWindowsDescender(run, size);
            // Sub-em tiers keep the legacy path: Calibri/Times/Magneto and the synthetic faces resolve
            // below a full em, so only full-em tiers take the branch.
            if (!(size > 0d) || tier < size)
            {
                continue;
            }
            if (maxTier is null || tier > maxTier.Value)
            {
                maxTier = tier;
            }
        }

        return maxTier;
    }

    public static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs, IDocxTextMeasurer? measurer = null)
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
        // Word 16.0 cell-inset probes (Calibri/Tahoma single cells 710.50 vs 710.02 at
        // 10pt): the rule is max(hheaAscender, 0.94em) (Tahoma gap 0.48 matches hheaAsc
        // 1.0005 over Calibri 0.9521). Opt-in measurer; null keeps legacy.
        double maxSize = firstTextParagraph.Runs.Max(run => run.EffectiveProperties.FontSize);
        double? hheaAscender = null;
        if (measurer is IDocxLineMetricsProvider provider)
        {
            DocxTextRun? widest = null;
            double widestSize = -1d;
            foreach (DocxTextRun run in firstTextParagraph.Runs)
            {
                double size = run.EffectiveProperties.FontSize;
                if (size > widestSize)
                {
                    widestSize = size;
                    widest = run;
                }
            }

            if (widest is not null)
            {
                hheaAscender = provider.MeasureHheaAscender(widest, maxSize);
            }
        }

        // The visible first baseline rides ResolveTableCellBodyBaselineOffset at the transition; the inset joins it there for non-top vertical alignment coherence while cancelling out of top-aligned first lines.
        if (ResolveTableCellFirstBaselinePoints(paragraphs, firstTextParagraph, maxSize, measurer) is { } cellFirstInset)
        {
            return cellFirstInset;
        }

        return Math.Max(maxSize * WordAutoLineBaselineOffsetEm, hheaAscender ?? 0d);
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
