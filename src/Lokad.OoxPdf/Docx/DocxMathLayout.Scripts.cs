using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathScripts? GetScripts(DocxEquation equation) =>
        GetSupportedScripts(GetSingleSupportedChild(equation));

    private static DocxMathScripts? GetSupportedScripts(DocxMathNode? node) =>
        node is DocxMathScripts { PropertiesXml.Count: 0 } scripts &&
        (scripts.Subscript is not null || scripts.Superscript is not null) ? scripts : null;

    private static bool TryMeasureScripts(DocxEquation equation, DocxMathScripts scripts,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token, out DocxMathBox? box, out string? fallback)
    {
        box = null; fallback = null;
        if (!TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> sources) || sources.Count == 0)
        { fallback = "UnsupportedMathScriptSource"; return false; }
        DocxMathRun owner = sources[0].Run;
        PdfFontResource? ownerResource = resolveFont(owner);
        if (ownerResource is null || !ownerResource.Font.Font.TryGetMathConstants(out OpenTypeMathConstants constants))
        { fallback = "MissingMathScriptConstants"; return false; }
        OpenTypeFont font = ownerResource.Font.Font;
        if (!HasValidScriptConstants(constants, equation.Cramped))
        { fallback = "InvalidMathScriptConstants"; return false; }
        double size = owner.Text.FontSize;
        foreach (DocxMathRunSource source in sources)
        {
            token.ThrowIfCancellationRequested();
            if (source.Run.Text.FontSize != size || !source.Run.NormalText && !ReferenceEquals(resolveFont(source.Run)?.Font.Font, font))
            { fallback = "MixedMathScriptMetrics"; return false; }
        }
        DocxEquation baseEquation = equation with { Content = scripts.Base };
        bool fractionBase = GetFraction(baseEquation) is not null;
        // Discovery has already required literal fraction children. This one
        // composition step cannot recurse through arbitrary authored nesting.
        bool measuredBase = fractionBase
            ? TryMeasure(baseEquation, resolveFont, token, out DocxMathBox? basis, out fallback)
            : DocxMathRowLayout.TryMeasure(scripts.Base, resolveFont, token, out basis, out fallback, exportFontSize: true);
        if (!measuredBase ||
            basis is null || basis.Glyphs.Count == 0) { fallback ??= "EmptyMathScriptBase"; return false; }
        if (basis.Glyphs.Any(glyph => !ReferenceEquals(glyph.Font.Font.Font, font)))
            { fallback = "MixedMathScriptBaseMetrics"; return false; }
        DocxMathBox? subscript = null, superscript = null;
        if (scripts.Subscript is { } sub && (!TryMeasureScriptArgument(sub, resolveFont, font, size, token, out subscript, out fallback) ||
            subscript is null)) { fallback ??= "EmptyMathSubscript"; return false; }
        if (scripts.Superscript is { } sup && (!TryMeasureScriptArgument(sup, resolveFont, font, size, token, out superscript, out fallback) ||
            superscript is null)) { fallback ??= "EmptyMathSuperscript"; return false; }
        double scale = size / font.UnitsPerEm;
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * scale;
        bool boxBase = fractionBase || basis.Glyphs.Count != 1;
        if (!boxBase)
        {
            ushort glyph = basis.Glyphs[0].OriginalGlyph;
            if (!font.TryIsMathExtendedShape(glyph, out boxBase))
            { fallback = "InvalidMathExtendedShapeCoverage"; return false; }
        }
        GetScriptShifts(constants, size / font.UnitsPerEm, basis, subscript, superscript,
            boxBase, equation.Cramped, out double up, out double down);
        DocxMathGlyphPlacement last = basis.Glyphs[^1];
        ushort lastGlyph = last.OriginalGlyph;
        short correction = 0;
        if (!fractionBase && !font.TryGetMathItalicCorrection(lastGlyph, out correction))
        { fallback = "InvalidMathScriptItalicCorrection"; return false; }
        double subX = basis.AdvanceWidth - correction * last.FontSize / font.UnitsPerEm;
        double supX = basis.AdvanceWidth;
        if (basis.Glyphs.Count == 1)
        {
            if (subscript?.Glyphs.Count == 1)
            {
                if (!TryScriptKern(last, subscript.Glyphs[0], -down, superscript: false, token, out double kern))
                { fallback = "InvalidMathScriptKern"; return false; }
                subX += kern;
            }
            if (superscript?.Glyphs.Count == 1)
            {
                if (!TryScriptKern(last, superscript.Glyphs[0], up, superscript: true, token, out double kern))
                { fallback = "InvalidMathScriptKern"; return false; }
                supX += kern;
            }
        }
        double advanceWidth = Math.Max(basis.AdvanceWidth, Math.Max(subscript is null ? 0d : subX + subscript.AdvanceWidth,
            superscript is null ? 0d : supX + superscript.AdvanceWidth)) + Value(OpenTypeMathConstant.SpaceAfterScript);
        double width = Math.Max(advanceWidth, Math.Max(basis.Width, Math.Max(subscript is null ? 0d : subX + subscript.Width,
            superscript is null ? 0d : supX + superscript.Width)) + Value(OpenTypeMathConstant.SpaceAfterScript));
        double above = Math.Max(basis.AboveBaseline, Math.Max(subscript is null ? 0d : subscript.AboveBaseline - down,
            superscript is null ? 0d : superscript.AboveBaseline + up));
        double below = Math.Max(basis.BelowBaseline, Math.Max(subscript is null ? 0d : subscript.BelowBaseline + down,
            superscript is null ? 0d : superscript.BelowBaseline - up));
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(up) || !double.IsFinite(down) || width < 0d || above < 0d || below < 0d || up < 0d || down < 0d || subX < 0d)
        { fallback = "InvalidMathScriptExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(basis.Glyphs.Count + (subscript?.Glyphs.Count ?? 0) + (superscript?.Glyphs.Count ?? 0));
        var rules = new List<DocxMathRulePlacement>(basis.Rules.Count);
        Append(basis, 0d, 0d);
        if (subscript is not null) { Append(subscript, subX, -down); }
        if (superscript is not null) { Append(superscript, supX, up); }
        double? enclosureBelow = null;
        if (subscript is not null)
        {
            double descent = below;
            foreach (var glyph in subscript.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                OpenTypeFont argumentFont = glyph.Font.Font.Font;
                descent = Math.Max(descent, down - glyph.BaselineY +
                    Math.Max(0d, -argumentFont.Os2.TypographicDescender * glyph.FontSize / argumentFont.UnitsPerEm));
            }
            enclosureBelow = descent;
        }
        box = new DocxMathBox(width, above, below, glyphs)
        { LogicalAdvance = advanceWidth == width ? null : advanceWidth, Rules = rules, ParentMetrics = basis.ParentMetrics,
            EnclosureBelowBaseline = enclosureBelow };
        fallback = null;
        return true;

        void Append(DocxMathBox child, double x, double baseline)
        {
            foreach (DocxMathGlyphPlacement glyph in child.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                glyphs.Add(glyph with { X = glyph.X + x, BaselineY = glyph.BaselineY + baseline });
            }
            foreach (DocxMathRulePlacement rule in child.Rules)
            {
                token.ThrowIfCancellationRequested();
                rules.Add(rule with { X = rule.X + x, Y = rule.Y + baseline });
            }
        }
    }

    private static bool TryMeasureScriptArgument(DocxMathRow argument, Func<DocxMathRun, PdfFontResource?> resolveFont,
        OpenTypeFont metricFont, double size, CancellationToken token, out DocxMathBox? child, out string? fallback)
    {
        child = null; fallback = "UnsupportedMathScriptArgument";
        return DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level) &&
            TryMeasureLiteralArgument(argument, resolveFont, metricFont, size, level, token, out child, out fallback, out _,
                exportFontSize: true, allowArgumentSize: true);
    }

    private static bool HasValidScriptConstants(OpenTypeMathConstants constants, bool cramped)
    {
        ReadOnlySpan<OpenTypeMathConstant> required = [OpenTypeMathConstant.SubscriptShiftDown, OpenTypeMathConstant.SubscriptTopMax,
            OpenTypeMathConstant.SubscriptBaselineDropMin,
            cramped ? OpenTypeMathConstant.SuperscriptShiftUpCramped : OpenTypeMathConstant.SuperscriptShiftUp,
            OpenTypeMathConstant.SuperscriptBottomMin,
            OpenTypeMathConstant.SuperscriptBaselineDropMax, OpenTypeMathConstant.SubSuperscriptGapMin,
            OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript, OpenTypeMathConstant.SpaceAfterScript];
        foreach (OpenTypeMathConstant field in required)
            if (constants.GetValue(field) < 0) { return false; }
        return true;
    }

    private static void GetScriptShifts(OpenTypeMathConstants constants, double scale, DocxMathBox basis,
        DocxMathBox? subscript, DocxMathBox? superscript, bool boxBase, bool cramped, out double up, out double down)
    {
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * scale;
        up = down = 0d;
        if (superscript is not null)
            up = Math.Max(Value(cramped ? OpenTypeMathConstant.SuperscriptShiftUpCramped : OpenTypeMathConstant.SuperscriptShiftUp), Math.Max(
                boxBase ? basis.AboveBaseline - Value(OpenTypeMathConstant.SuperscriptBaselineDropMax) : 0d,
                superscript.BelowBaseline + Value(OpenTypeMathConstant.SuperscriptBottomMin)));
        if (subscript is not null)
            down = Math.Max(Value(OpenTypeMathConstant.SubscriptShiftDown), Math.Max(
                subscript.AboveBaseline - Value(OpenTypeMathConstant.SubscriptTopMax),
                boxBase ? basis.BelowBaseline + Value(OpenTypeMathConstant.SubscriptBaselineDropMin) : 0d));
        if (subscript is not null && superscript is not null)
        {
            double missing = Math.Max(0d, Value(OpenTypeMathConstant.SubSuperscriptGapMin) -
                (up + down - superscript.BelowBaseline - subscript.AboveBaseline));
            double raise = Math.Min(missing, Math.Max(0d,
                Value(OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript) - (up - superscript.BelowBaseline)));
            up += raise;
            down += missing - raise;
        }
    }

    private static bool TryScriptKern(DocxMathGlyphPlacement basis, DocxMathGlyphPlacement script,
        double scriptBaseline, bool superscript, CancellationToken token, out double result)
    {
        result = 0d;
        token.ThrowIfCancellationRequested();
        OpenTypeFont baseFont = basis.Font.Font.Font, scriptFont = script.Font.Font.Font;
        ushort baseGlyph = basis.OriginalGlyph;
        ushort scriptGlyph = script.OriginalGlyph;
        if (scriptGlyph == 0 || !baseFont.TryGetGlyphBounds(baseGlyph, out var baseBounds) ||
            !scriptFont.TryGetGlyphBounds(scriptGlyph, out var scriptBounds)) { return false; }
        double baseScale = basis.FontSize / baseFont.UnitsPerEm, scriptScale = script.FontSize / scriptFont.UnitsPerEm;
        MathKernCorner baseCorner = superscript ? MathKernCorner.TopRight : MathKernCorner.BottomRight;
        MathKernCorner scriptCorner = superscript ? MathKernCorner.BottomLeft : MathKernCorner.TopLeft;
        double scriptEdge = (superscript ? scriptBounds.YMin : scriptBounds.YMax) * scriptScale;
        double baseEdge = (superscript ? baseBounds.YMax : baseBounds.YMin) * baseScale;
        if (!baseFont.TryGetMathKern(baseGlyph, baseCorner, (scriptBaseline + scriptEdge) / baseScale, out short baseAtScript) ||
            !scriptFont.TryGetMathKern(scriptGlyph, scriptCorner, scriptEdge / scriptScale, out short scriptAtScript) ||
            !baseFont.TryGetMathKern(baseGlyph, baseCorner, baseEdge / baseScale, out short baseAtBase) ||
            !scriptFont.TryGetMathKern(scriptGlyph, scriptCorner, (baseEdge - scriptBaseline) / scriptScale, out short scriptAtBase)) { return false; }
        // Word lower/upper-pair comparisons favor the greater corner sum.
        // The OpenType overview describes the minimum; retain this deliberate
        // Office behavior with independent glyph/height contrasts and gates.
        result = Math.Max(baseAtScript * baseScale + scriptAtScript * scriptScale,
            baseAtBase * baseScale + scriptAtBase * scriptScale);
        return double.IsFinite(result);
    }
}
