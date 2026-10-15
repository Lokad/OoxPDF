using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathRadical? GetRadical(DocxEquation equation) => GetSupportedRadical(GetSingleSupportedChild(equation));
    private static DocxMathRadical? GetSupportedRadical(DocxMathNode? node) =>
        node is DocxMathRadical { HasUnsupportedProperties: false, DegreeHidden: true, Degree.Children.Count: 0 } radical &&
        radical.Degree.HasSupportedProperties && (radical.ControlText is null ||
            DocxMathRowLayout.HasSupportedTextFormatting(radical.ControlText, allowBold: false)) ? radical : null;

    private static bool TryMeasureRadical(DocxEquation equation, DocxMathRadical radical,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<DocxMathRadical, PdfFontResource?>? resolveRadicalFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathRadicalSource"; return false; }
        PdfFontResource? resource = resolveRadicalFont is null ? resolveFont(sources[0].Run) : resolveRadicalFont(radical);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathRadicalConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = radical.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "MixedMathRadicalMetrics"; return false; }
        if (!DocxMathRowLayout.TryMeasure(radical.Radicand, resolveFont, token, out var child, out fallback,
            exportFontSize: true, operatorSpacingFontSize: size) || child is null || child.Glyphs.Count == 0)
        { fallback ??= "EmptyMathRadical"; return false; }
        double scale = size / font.UnitsPerEm;
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), glyphScale = glyphSize / font.UnitsPerEm;
        double gap = constants.GetValue(equation.Display ? OpenTypeMathConstant.RadicalDisplayStyleVerticalGap : OpenTypeMathConstant.RadicalVerticalGap) * scale;
        double rule = constants.GetValue(OpenTypeMathConstant.RadicalRuleThickness) * scale;
        double extra = constants.GetValue(OpenTypeMathConstant.RadicalExtraAscender) * scale;
        double needed = child.AboveBaseline + child.BelowBaseline + gap + rule;
        if (gap < 0d || rule <= 0d || extra < 0d || !double.IsFinite(needed) ||
            !font.TryGetMathVerticalVariant(font.MapCodePoint(0x221A), needed / glyphScale, token, out var selected) ||
            !font.TryGetGlyphBounds(selected.Glyph, out var bounds) || !resource.Font.TryGetEncodedCid(selected.Glyph, out ushort cid))
        { fallback = "UnsupportedMathRadicalVariant"; return false; }
        gap += Math.Max(0d, selected.Advance * glyphScale - needed) / 2d;
        double baseline = child.AboveBaseline + gap + rule - bounds.YMax * glyphScale;
        double childX = font.GetAdvanceWidth(selected.Glyph) * glyphScale;
        double advance = childX + child.AdvanceWidth;
        double width = Math.Max(advance, Math.Max(childX + child.Width, bounds.XMax * glyphScale));
        double above = Math.Max(child.AboveBaseline, baseline + bounds.YMax * glyphScale + extra);
        double below = Math.Max(child.BelowBaseline, -baseline - bounds.YMin * glyphScale);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(baseline) || width <= 0d || above < 0d || below < 0d)
        { fallback = "InvalidMathRadicalExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(child.Glyphs.Count + 1);
        var color = child.Glyphs[0].Color;
        if (radical.ControlText is { } control)
        { color = default; Lokad.OoxPdf.Pptx.RgbColor.TryParse(control.ColorHex, out color); }
        OoxConversionBudget.Current?.ChargeSceneNodes(2);
        glyphs.Add(new(0x221A, 0x221A, resource, cid.ToString("X4", CultureInfo.InvariantCulture), glyphSize, 0d, baseline,
            childX, color) { OriginalGlyph = selected.Glyph });
        foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + childX }); }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, ParentMetrics = new(font, size),
            Rules = [new(childX, child.AboveBaseline + gap, child.AdvanceWidth, rule, color)] };
        fallback = null;
        return true;
    }
}
