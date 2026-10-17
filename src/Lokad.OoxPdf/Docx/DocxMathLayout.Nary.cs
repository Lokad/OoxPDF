using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // Node placement overrides the document owner. Inline equations place
    // limits beside the operator, as Office does for undOvr.
    // Integral sizing retains its independently measured fallback.
    private static DocxMathNary? GetNary(DocxEquation equation, CancellationToken token) =>
        GetSupportedNary(GetSingleSupportedChild(equation), equation, token);

    private static DocxMathNary? GetSupportedNary(DocxMathNode? node, DocxEquation equation, CancellationToken token) =>
        node is DocxMathNary
        { HasUnsupportedProperties: false, CharacterValue: "\u2211" or "\u220F" or "\u222B", LimitLocationValue: null or "undOvr" or "subSup" } nary &&
        (!equation.Display || nary.CharacterValue != "\u222B") &&
        NaryLimitLocation(equation, nary) is "undOvr" or "subSup" &&
        (nary.SubscriptHiddenValue != true || IsEmptyLiteralArgument(nary.Subscript, token)) &&
        (nary.SuperscriptHiddenValue != true || IsEmptyLiteralArgument(nary.Superscript, token)) &&
        (nary.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(nary.ControlText, allowBold: false)) ? nary : null;

    private static string? NaryLimitLocation(DocxEquation equation, DocxMathNary nary) =>
        nary.LimitLocationValue ?? (nary.CharacterValue == "\u222B"
            ? equation.LimitSettings.UnsupportedIntegral ? null : equation.LimitSettings.IntegralValue ?? "subSup"
            : equation.LimitSettings.UnsupportedNary ? null : equation.LimitSettings.NaryValue ?? "undOvr");

    // Word reassigns stored nonempty hidden arguments into visible roles.
    // Only a truly empty supported literal argument may be omitted here.
    private static bool IsEmptyLiteralArgument(DocxMathRow row, CancellationToken token) =>
        DocxMathRowLayout.TryGetRenderCodePoints(row, token, out var points, out _, allowArgumentSize: true) && points.Count == 0;

    private static bool TryMeasureNary(DocxEquation equation, DocxMathNary nary,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources))
        { fallback = "UnsupportedMathNarySource"; return false; }
        PdfFontResource? resource = ResolveControlResource(nary, sources, resolveFont, resolveControlFont);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathNaryConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = ResolveControlSize(nary, equation, sources);
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathNarySize"; return false; }
        ReadOnlySpan<OpenTypeMathConstant> required = [OpenTypeMathConstant.UpperLimitGapMin,
            OpenTypeMathConstant.UpperLimitBaselineRiseMin, OpenTypeMathConstant.LowerLimitGapMin,
            OpenTypeMathConstant.LowerLimitBaselineDropMin];
        foreach (var field in required)
            if (constants.GetValue(field) < 0) { fallback = "InvalidMathNaryConstants"; return false; }
        if (!(equation.Display && NaryLimitLocation(equation, nary) == "undOvr") && !HasValidScriptConstants(constants, equation.Cramped))
        { fallback = "InvalidMathNaryScriptConstants"; return false; }
        if (!TryMeasureBoxArgument(equation with { Content = nary.Body }, resolveFont, font, size, token,
                out var body, out fallback) || body is null)
        { fallback ??= "UnsupportedMathNaryBody"; return false; }
        DocxMathBox? lower = null, upper = null;
        if (nary.SubscriptHiddenValue != true &&
            (!TryMeasureScriptArgument(nary.Subscript, resolveFont, font, size, token, out lower, out fallback) || lower is null) ||
            nary.SuperscriptHiddenValue != true &&
            (!TryMeasureScriptArgument(nary.Superscript, resolveFont, font, size, token, out upper, out fallback) || upper is null))
        { fallback ??= "UnsupportedMathNaryLimit"; return false; }
        int point = char.ConvertToUtf32(nary.CharacterValue!, 0);
        ushort canonical = font.MapCodePoint(point);
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), scale = glyphSize / font.UnitsPerEm;
        double target = equation.Display ? constants.DisplayOperatorMinHeight : 0d;
        // Unlike delimiters, an omitted n-ary grow property defaults to off.
        // Explicit growth fits the body's occupied height, without imposing
        // a delimiter's symmetric extension around the math axis.
        if (nary.GrowValue == true) { target = Math.Max(target, (body.AboveBaseline + body.BelowBaseline) / scale); }
        if (!double.IsFinite(target) || target < 0d ||
            !font.TryGetMathVerticalVariant(canonical, target, token, out var variant) ||
            !font.TryGetGlyphBounds(variant.Glyph, out var bounds) || !resource.Font.TryGetEncodedCid(variant.Glyph, out ushort cid))
        { fallback = "UnsupportedMathNaryVariant"; return false; }
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * size / font.UnitsPerEm;
        double baseline = Value(OpenTypeMathConstant.AxisHeight) - (bounds.YMax + bounds.YMin) / 2d * scale;
        double operatorAdvance = font.GetAdvanceWidth(variant.Glyph) * scale;
        bool stacked = equation.Display && NaryLimitLocation(equation, nary) == "undOvr";
        double groupAdvance, operatorX = 0d, upperX, lowerX, upperY, lowerY, bodyX;
        if (stacked)
        {
            groupAdvance = Math.Max(operatorAdvance, Math.Max((lower?.AdvanceWidth ?? 0d), (upper?.AdvanceWidth ?? 0d)));
            operatorX = (groupAdvance - operatorAdvance) / 2d;
            upperX = (groupAdvance - (upper?.AdvanceWidth ?? 0d)) / 2d;
            lowerX = (groupAdvance - (lower?.AdvanceWidth ?? 0d)) / 2d;
            upperY = baseline + bounds.YMax * scale + Math.Max(Value(OpenTypeMathConstant.UpperLimitBaselineRiseMin),
                Value(OpenTypeMathConstant.UpperLimitGapMin) + (upper?.BelowBaseline ?? 0d));
            lowerY = baseline + bounds.YMin * scale - Math.Max(Value(OpenTypeMathConstant.LowerLimitBaselineDropMin),
                Value(OpenTypeMathConstant.LowerLimitGapMin) + (lower?.AboveBaseline ?? 0d));
            bodyX = groupAdvance + size / 6d;
        }
        else
        {
            var basis = new DocxMathBox(operatorAdvance, Math.Max(0d, bounds.YMax * scale), Math.Max(0d, -bounds.YMin * scale), []);
            if (!font.TryIsMathExtendedShape(variant.Glyph, out bool extended) ||
                !font.TryGetMathItalicCorrection(variant.Glyph, out short correction))
            { fallback = "InvalidMathNaryScriptMetrics"; return false; }
            GetScriptShifts(constants, size / font.UnitsPerEm, basis, lower, upper, extended, equation.Cramped,
                out double up, out double down);
            upperY = baseline + up; lowerY = baseline - down;
            upperX = operatorAdvance;
            lowerX = operatorAdvance - correction * scale;
            groupAdvance = Math.Max(operatorAdvance, Math.Max(lower is null ? 0d : lowerX + lower.AdvanceWidth,
                upper is null ? 0d : upperX + upper.AdvanceWidth));
            // Word's inline large-operator group retains a trailing thin space;
            // explicit display side limits abut the body after the script box.
            bodyX = groupAdvance + (equation.Display ? 0d : size / 6d);
        }
        double advance = bodyX + body.AdvanceWidth;
        double width = Math.Max(advance, operatorX + bounds.XMax * scale);
        double above = Math.Max(0d, baseline + bounds.YMax * scale), below = Math.Max(0d, -baseline - bounds.YMin * scale);
        RgbColor color = body.Glyphs.Count == 0 ? default : body.Glyphs[0].Color;
        if (nary.ControlText is { } control) { color = default; RgbColor.TryParse(control.ColorHex, out color); }
        OoxConversionBudget.Current?.ChargeSceneNodes(1);
        var glyphs = new List<DocxMathGlyphPlacement> { new(point, point, resource, cid.ToString("X4", CultureInfo.InvariantCulture),
            glyphSize, operatorX, baseline, operatorAdvance, color) { OriginalGlyph = variant.Glyph, ControlOwner = nary } };
        var rules = new List<DocxMathRulePlacement>();
        Append(body, bodyX, 0d);
        if (upper is not null) { Append(upper, upperX, upperY); }
        if (lower is not null) { Append(lower, lowerX, lowerY); }
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) || width <= 0d ||
            lower is not null && lowerX < 0d || upper is not null && upperX < 0d)
        { fallback = "InvalidMathNaryExtent"; return false; }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = new(font, size),
            Edges = new(DocxMathRowLayout.AtomKind.LargeOperator, body.Edges?.Last ?? DocxMathRowLayout.AtomKind.Ordinary, false) };
        fallback = null;
        return true;

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
