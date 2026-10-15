using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathRadical? GetRadical(DocxEquation equation) => GetSupportedRadical(GetSingleSupportedChild(equation));
    private static DocxMathRadical? GetSupportedRadical(DocxMathNode? node) =>
        node is DocxMathRadical { HasUnsupportedProperties: false } radical &&
        (radical.DegreeHidden == true ? radical.Degree.Children.Count == 0 && radical.Degree.HasSupportedProperties :
            radical.Degree.Children.Count > 0 && DocxMathArgumentMetrics.TryGetScriptLevel(radical.Degree, out _, startingLevel: 2)) &&
        (radical.ControlText is null ||
            DocxMathRowLayout.HasSupportedTextFormatting(radical.ControlText, allowBold: false)) ? radical : null;

    private static bool TryMeasureRadical(DocxEquation equation, DocxMathRadical radical,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathRadicalSource"; return false; }
        PdfFontResource? resource = resolveControlFont is null ? resolveFont(sources[0].Run) : resolveControlFont(radical);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathRadicalConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = radical.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "MixedMathRadicalMetrics"; return false; }
        DocxEquation radicandEquation = equation with { Content = radical.Radicand, ParentFontSize = size, Cramped = true };
        // Discovery admits existing bounded fraction/script boxes; their own
        // literal-role checks prevent arbitrary recursive radicands.
        bool measuredChild = TryMeasureBoxArgument(radicandEquation, resolveFont, font, size, token, out var child, out fallback);
        if (!measuredChild || child is null || child.Glyphs.Count == 0)
        { fallback ??= "EmptyMathRadical"; return false; }
        double scale = size / font.UnitsPerEm;
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), glyphScale = glyphSize / font.UnitsPerEm;
        double gap = constants.GetValue(equation.Display ? OpenTypeMathConstant.RadicalDisplayStyleVerticalGap : OpenTypeMathConstant.RadicalVerticalGap) * scale;
        double rule = constants.GetValue(OpenTypeMathConstant.RadicalRuleThickness) * scale;
        double extra = constants.GetValue(OpenTypeMathConstant.RadicalExtraAscender) * scale;
        double needed = child.AboveBaseline + (child.EnclosureBelowBaseline ?? child.BelowBaseline) + gap + rule;
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
        DocxMathBox? degree = null;
        double degreeX = 0d, degreeY = 0d, rootX = 0d;
        if (radical.DegreeHidden != true)
        {
            if (!DocxMathArgumentMetrics.TryGetScriptLevel(radical.Degree, out int level, startingLevel: 2) ||
                !DocxMathRowLayout.TryMeasure(radical.Degree, resolveFont, token, out degree, out fallback, level,
                    exportFontSize: true, allowArgumentSize: true, scriptMetricsFont: font) ||
                degree is null || degree.Glyphs.Count == 0)
            { fallback ??= "UnsupportedMathRadicalDegree"; return false; }
            degreeX = Math.Max(0d, constants.GetValue(OpenTypeMathConstant.RadicalKernBeforeDegree) * scale);
            rootX = degreeX + degree.AdvanceWidth + Math.Max(-degree.AdvanceWidth,
                constants.GetValue(OpenTypeMathConstant.RadicalKernAfterDegree) * scale);
            degreeY = baseline + bounds.YMin * glyphScale +
                constants.RadicalDegreeBottomRaisePercent / 100d * (bounds.YMax - bounds.YMin) * glyphScale + degree.BelowBaseline;
            childX += rootX;
            advance += rootX;
            width = Math.Max(width + rootX, degreeX + degree.Width);
            above = Math.Max(above, degreeY + degree.AboveBaseline);
            below = Math.Max(below, degree.BelowBaseline - degreeY);
        }
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(baseline) || width <= 0d || above < 0d || below < 0d)
        { fallback = "InvalidMathRadicalExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(child.Glyphs.Count + 1 + (degree?.Glyphs.Count ?? 0));
        var color = child.Glyphs[0].Color;
        if (radical.ControlText is { } control)
        { color = default; Lokad.OoxPdf.Pptx.RgbColor.TryParse(control.ColorHex, out color); }
        OoxConversionBudget.Current?.ChargeSceneNodes(2);
        glyphs.Add(new(0x221A, 0x221A, resource, cid.ToString("X4", CultureInfo.InvariantCulture), glyphSize, rootX, baseline,
            childX - rootX, color) { OriginalGlyph = selected.Glyph });
        foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + childX }); }
        if (degree is not null)
            foreach (var glyph in degree.Glyphs)
            { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + degreeX, BaselineY = glyph.BaselineY + degreeY }); }
        var rules = new List<DocxMathRulePlacement>(child.Rules.Count + 1)
            { new(childX, child.AboveBaseline + gap, child.AdvanceWidth, rule, color) };
        foreach (var childRule in child.Rules)
        { token.ThrowIfCancellationRequested(); rules.Add(childRule with { X = childRule.X + childX }); }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, ParentMetrics = new(font, size),
            Rules = rules };
        fallback = null;
        return true;
    }
}
