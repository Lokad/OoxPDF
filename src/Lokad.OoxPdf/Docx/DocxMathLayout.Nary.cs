using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // Explicit stacked display limits are independent of document-level limit
    // defaults. Side limits and integral sizing retain their own fallback.
    private static DocxMathNary? GetNary(DocxEquation equation) =>
        equation.Display && GetSingleSupportedChild(equation) is DocxMathNary
        { HasUnsupportedProperties: false, CharacterValue: "\u2211" or "\u220F", LimitLocationValue: "undOvr",
            GrowValue: not false, SubscriptHiddenValue: not true, SuperscriptHiddenValue: not true } nary &&
        (nary.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(nary.ControlText, allowBold: false)) ? nary : null;

    private static bool TryMeasureNary(DocxEquation equation, DocxMathNary nary,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathNarySource"; return false; }
        PdfFontResource? resource = resolveControlFont is null ? resolveFont(sources[0].Run) : resolveControlFont(nary);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathNaryConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = nary.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathNarySize"; return false; }
        ReadOnlySpan<OpenTypeMathConstant> required = [OpenTypeMathConstant.UpperLimitGapMin,
            OpenTypeMathConstant.UpperLimitBaselineRiseMin, OpenTypeMathConstant.LowerLimitGapMin,
            OpenTypeMathConstant.LowerLimitBaselineDropMin];
        foreach (var field in required)
            if (constants.GetValue(field) < 0) { fallback = "InvalidMathNaryConstants"; return false; }
        if (!DocxMathRowLayout.TryMeasure(nary.Body, resolveFont, token, out var body, out fallback,
                exportFontSize: true, operatorSpacingFontSize: size) || body is null || body.Glyphs.Count == 0 ||
            !MeasureLimit(nary.Subscript, out var lower, out fallback) || lower is null || lower.Glyphs.Count == 0 ||
            !MeasureLimit(nary.Superscript, out var upper, out fallback) || upper is null || upper.Glyphs.Count == 0)
        { fallback ??= "UnsupportedMathNaryArgument"; return false; }
        int point = char.ConvertToUtf32(nary.CharacterValue!, 0);
        ushort canonical = font.MapCodePoint(point);
        if (!font.TryGetMathVerticalVariant(canonical, constants.DisplayOperatorMinHeight, token, out var variant) ||
            !font.TryGetGlyphBounds(variant.Glyph, out var bounds) || !resource.Font.TryGetEncodedCid(variant.Glyph, out ushort cid))
        { fallback = "UnsupportedMathNaryVariant"; return false; }
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), scale = glyphSize / font.UnitsPerEm;
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * size / font.UnitsPerEm;
        double baseline = Value(OpenTypeMathConstant.AxisHeight) - (bounds.YMax + bounds.YMin) / 2d * scale;
        double operatorAdvance = font.GetAdvanceWidth(variant.Glyph) * scale;
        double groupAdvance = Math.Max(operatorAdvance, Math.Max(lower.AdvanceWidth, upper.AdvanceWidth));
        double operatorX = (groupAdvance - operatorAdvance) / 2d;
        double upperY = baseline + bounds.YMax * scale + Math.Max(Value(OpenTypeMathConstant.UpperLimitBaselineRiseMin),
            Value(OpenTypeMathConstant.UpperLimitGapMin) + upper.BelowBaseline);
        double lowerY = baseline + bounds.YMin * scale - Math.Max(Value(OpenTypeMathConstant.LowerLimitBaselineDropMin),
            Value(OpenTypeMathConstant.LowerLimitGapMin) + lower.AboveBaseline);
        // Three math units separate the large-operator group from its body.
        double bodyX = groupAdvance + size / 6d;
        double advance = bodyX + body.AdvanceWidth;
        double width = Math.Max(advance, operatorX + bounds.XMax * scale);
        double above = Math.Max(0d, baseline + bounds.YMax * scale), below = Math.Max(0d, -baseline - bounds.YMin * scale);
        RgbColor color = body.Glyphs[0].Color;
        if (nary.ControlText is { } control) { color = default; RgbColor.TryParse(control.ColorHex, out color); }
        OoxConversionBudget.Current?.ChargeSceneNodes(1);
        var glyphs = new List<DocxMathGlyphPlacement> { new(point, point, resource, cid.ToString("X4", CultureInfo.InvariantCulture),
            glyphSize, operatorX, baseline, operatorAdvance, color) { OriginalGlyph = variant.Glyph } };
        var rules = new List<DocxMathRulePlacement>();
        Append(body, bodyX, 0d);
        Append(upper, (groupAdvance - upper.AdvanceWidth) / 2d, upperY);
        Append(lower, (groupAdvance - lower.AdvanceWidth) / 2d, lowerY);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) || width <= 0d)
        { fallback = "InvalidMathNaryExtent"; return false; }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = new(font, size) };
        fallback = null;
        return true;

        bool MeasureLimit(DocxMathRow argument, out DocxMathBox? child, out string? reason)
        {
            child = null; reason = "UnsupportedMathNaryLimit";
            return DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level) &&
                DocxMathRowLayout.TryMeasure(argument, resolveFont, token, out child, out reason, scriptLevel: level,
                    exportFontSize: true, allowArgumentSize: true, scriptMetricsFont: font);
        }

        void Append(DocxMathBox child, double x, double y)
        {
            foreach (var glyph in child.Glyphs)
            { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + x, BaselineY = glyph.BaselineY + y }); }
            foreach (var rule in child.Rules)
            { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + x, Y = rule.Y + y }); }
            width = Math.Max(width, x + child.Width);
            above = Math.Max(above, y + child.AboveBaseline);
            below = Math.Max(below, child.BelowBaseline - y);
        }
    }
}
