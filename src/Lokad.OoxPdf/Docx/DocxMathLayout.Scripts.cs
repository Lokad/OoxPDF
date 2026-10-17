using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathScripts? GetScripts(DocxEquation equation) =>
        GetScriptArgument(equation.Content);

    private static DocxMathScripts? GetScriptArgument(DocxMathRow argument) =>
        argument.HasSupportedProperties && argument.Children.Count == 1
            ? GetSupportedScripts(argument.Children[0]) : null;

    private static DocxMathScripts? GetSupportedScripts(DocxMathNode? node) =>
        node is DocxMathScripts { PropertiesXml.Count: 0 } scripts &&
        (scripts.Subscript is not null || scripts.Superscript is not null) ? scripts : null;

    internal static bool IsEmptyScripts(DocxEquation equation, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        return GetScripts(equation) is { } scripts && IsEmptyArgument(scripts.Base, token) &&
            (scripts.Subscript is null || IsEmptyArgument(scripts.Subscript, token, allowArgumentSize: true)) &&
            (scripts.Superscript is null || IsEmptyArgument(scripts.Superscript, token, allowArgumentSize: true));
    }

    private static bool TryMeasureScripts(DocxEquation equation, DocxMathScripts scripts,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token, out DocxMathBox? box, out string? fallback,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics, int startingLevel = 0)
    {
        box = null; fallback = null;
        if (!TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> sources) ||
            sources.Count == 0 && !IsEmptyScripts(equation, token))
        { fallback = "UnsupportedMathScriptSource"; return false; }
        bool metricOnly = IsEmptyScripts(equation, token);
        var firstSource = sources.FirstOrDefault(source => source.Run.Text.Text.Length != 0);
        OpenTypeFont? font = metricOnly ? resolveControlMetrics?.Invoke(scripts) :
            firstSource is null ? null : resolveFont(firstSource.Run)?.Font.Font;
        if (font is null || !font.TryGetMathConstants(out OpenTypeMathConstants constants))
        { fallback = "MissingMathScriptConstants"; return false; }
        if (!HasValidScriptConstants(constants, equation.Cramped))
        { fallback = "InvalidMathScriptConstants"; return false; }
        double size = metricOnly ? ResolveControlSize(scripts, equation, sources) : firstSource!.Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathScriptSize"; return false; }
        foreach (DocxMathRunSource source in sources)
        {
            token.ThrowIfCancellationRequested();
            if (source.Run.Text.Text.Length == 0) continue;
            if (source.Run.Text.FontSize != size || !source.Run.NormalText && !ReferenceEquals(resolveFont(source.Run)?.Font.Font, font))
            { fallback = "MixedMathScriptMetrics"; return false; }
        }
        double styleSize = size;
        if (startingLevel != 0 && !DocxMathArgumentMetrics.TryGetNominalScriptFontSize(font, size, startingLevel, out styleSize))
        { fallback = "MissingMathScriptScale"; return false; }
        DocxEquation baseEquation = equation with { Content = scripts.Base };
        bool fractionBase = GetFraction(baseEquation) is not null;
        // Discovery has already required literal fraction children. This one
        // composition step cannot recurse through arbitrary authored nesting.
        bool measuredBase = fractionBase
            ? TryMeasure(baseEquation, resolveFont, token, out DocxMathBox? basis, out fallback)
            : TryMeasureLiteralArgument(scripts.Base, resolveFont, font, size, startingLevel, token, out basis, out fallback, out _, exportFontSize: true);
        if (!measuredBase || basis is null) { fallback ??= "EmptyMathScriptBase"; return false; }
        if (basis.Glyphs.Any(glyph => !ReferenceEquals(glyph.Font.Font.Font, font)))
            { fallback = "MixedMathScriptBaseMetrics"; return false; }
        DocxMathBox? subscript = null, superscript = null;
        int argumentLevel = Math.Min(2, startingLevel + 1);
        if (scripts.Subscript is { } sub && (!TryMeasureScriptArgument(sub, resolveFont, font, size, token, out subscript, out fallback, argumentLevel) ||
            subscript is null)) { fallback ??= "EmptyMathSubscript"; return false; }
        if (scripts.Superscript is { } sup && (!TryMeasureScriptArgument(sup, resolveFont, font, size, token, out superscript, out fallback, argumentLevel) ||
            superscript is null)) { fallback ??= "EmptyMathSuperscript"; return false; }
        double scale = styleSize / font.UnitsPerEm;
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * scale;
        bool boxBase = fractionBase || basis.Glyphs.Count > 1;
        if (!boxBase)
        {
            // An empty literal base is one metric-only placeholder atom,
            // with the same extended-shape classification as its font glyph.
            ushort glyph = basis.Glyphs.Count == 0 ? font.MapCodePoint(HiddenPlaceholderCodePoint) : basis.Glyphs[0].OriginalGlyph;
            if (!font.TryIsMathExtendedShape(glyph, out boxBase))
            { fallback = "InvalidMathExtendedShapeCoverage"; return false; }
        }
        GetScriptShifts(constants, scale, basis, subscript, superscript,
            boxBase, equation.Cramped, out double up, out double down);
        DocxMathGlyphPlacement? last = basis.Glyphs.Count == 0 ? null : basis.Glyphs[^1];
        short correction = 0;
        if (!fractionBase && last is not null && !font.TryGetMathItalicCorrection(last.OriginalGlyph, out correction))
        { fallback = "InvalidMathScriptItalicCorrection"; return false; }
        double subX = basis.AdvanceWidth - (last is null ? 0d : correction * last.FontSize / font.UnitsPerEm);
        double supX = basis.AdvanceWidth;
        if (basis.Glyphs.Count == 1 && last is not null)
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
        { LogicalAdvance = advanceWidth == width ? null : advanceWidth, Rules = rules, ParentMetrics = basis.ParentMetrics ?? (basis.Glyphs.Count == 0 ? new(font, styleSize) : null),
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
        OpenTypeFont metricFont, double size, CancellationToken token, out DocxMathBox? child, out string? fallback, int startingLevel = 1)
    {
        child = null; fallback = "UnsupportedMathScriptArgument";
        return DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level, startingLevel) &&
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
