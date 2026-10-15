using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

// Shared capability discovery and composition. Literal rows retain their
// existing layout; one display fraction adds two literal child boxes and a rule.
internal sealed record DocxMathRunSource(DocxMathRun Run, int ScriptLevel);

internal static partial class DocxMathLayout
{
    public static bool TryGetLineMetrics(DocxMathBox box, CancellationToken token, out double above, out double height,
        double lineSpacingFactor = 1d)
    {
        above = box.AboveBaseline;
        double below = box.BelowBaseline, leading = 0d, naturalLine = 0d;
        height = 0d;
        if (!double.IsFinite(lineSpacingFactor) || lineSpacingFactor < 1d) { return false; }
        foreach (DocxMathGlyphPlacement glyph in box.Glyphs)
        {
            token.ThrowIfCancellationRequested();
            var font = glyph.Font.Font.Font;
            above = Math.Max(above, DocxLineMetrics.MeasureTypographicBaselineInset(font, glyph.FontSize));
            below = Math.Max(below, -font.Os2.TypographicDescender * glyph.FontSize / font.UnitsPerEm);
            naturalLine = Math.Max(naturalLine, DocxLineMetrics.MeasureTypographicLineHeight(font, glyph.FontSize));
            if (font.TryGetMathConstants(out OpenTypeMathConstants constants))
                leading = Math.Max(leading, constants.GetValue(OpenTypeMathConstant.MathLeading) * glyph.FontSize / font.UnitsPerEm);
        }
        if (box.ParentMetrics is { } parent && parent.Font.TryGetMathConstants(out OpenTypeMathConstants parentConstants))
            leading = Math.Max(leading, parentConstants.GetValue(OpenTypeMathConstant.MathLeading) * parent.Size / parent.Font.UnitsPerEm);
        // MathLeading protects tall ink above the ordinary typographic box.
        // Apply it once to the composed box, rather than once per child/glyph.
        above = Math.Max(above, box.AboveBaseline + leading);
        // Automatic paragraph expansion adds ordinary font-line spacing after
        // the occupied math frame; it does not scale tall ink or its baseline.
        height = above + below + naturalLine * (lineSpacingFactor - 1d);
        return double.IsFinite(above) && double.IsFinite(height) && height > 0d;
    }

    public static bool TryGetRuns(DocxEquation equation, CancellationToken token, out IReadOnlyList<DocxMathRunSource> runs)
    {
        var rows = new List<(DocxMathRow Row, int Level, bool AllowArgumentSize)>();
        if (GetScripts(equation) is { } scripts)
        {
            if (GetFraction(equation with { Content = scripts.Base }) is { } basis)
            {
                int level = equation.Display ? 0 : 1;
                rows.Add((basis.Numerator, level, false));
                rows.Add((basis.Denominator, level, false));
            }
            else { rows.Add((scripts.Base, 0, false)); }
            foreach (DocxMathRow? argument in new[] { scripts.Subscript, scripts.Superscript })
            {
                if (argument is null) { continue; }
                if (!DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level)) { runs = []; return false; }
                rows.Add((argument, level, true));
            }
        }
        else if (GetFraction(equation) is { } fraction)
        { int level = equation.Display ? 0 : 1; rows.Add((fraction.Numerator, level, false)); rows.Add((fraction.Denominator, level, false)); }
        else if (GetFractionRow(equation, token) is { } compound)
        {
            rows.Add((compound.Before, 0, false));
            int level = equation.Display ? 0 : 1;
            rows.Add((compound.Fraction.Numerator, level, false));
            rows.Add((compound.Fraction.Denominator, level, false));
            rows.Add((compound.After, 0, false));
        }
        else { rows.Add((equation.Content, 0, false)); }
        var result = new List<DocxMathRunSource>();
        foreach ((DocxMathRow row, int level, bool allowArgumentSize) in rows)
        {
            if (!DocxMathRowLayout.TryGetRenderCodePoints(row, token, out _, out _, allowArgumentSize)) { runs = []; return false; }
            var pending = new Stack<DocxMathNode>();
            pending.Push(row);
            while (pending.TryPop(out DocxMathNode? node))
            {
                token.ThrowIfCancellationRequested();
                if (node is DocxMathRun run) { result.Add(new(run, level)); }
                else if (node is DocxMathRow sequence)
                    for (int index = sequence.Children.Count - 1; index >= 0; index--) { pending.Push(sequence.Children[index]); }
            }
        }
        runs = result;
        return true;
    }

    public static bool TryMeasure(DocxEquation equation, Func<DocxMathRun, PdfFontResource?> resolveFont,
        CancellationToken token, out DocxMathBox? box, out string? fallback, OpenTypeFont? parentMetricsFont = null)
    {
        if (GetScripts(equation) is { } scripts)
            return TryMeasureScripts(equation, scripts, resolveFont, token, out box, out fallback);
        if (GetFractionRow(equation, token) is { } compound)
            return TryMeasureFractionRow(equation, compound, resolveFont, token, out box, out fallback, parentMetricsFont);
        if (GetFraction(equation) is not { } fraction)
            return DocxMathRowLayout.TryMeasure(equation.Content, resolveFont, token, out box, out fallback);
        box = null;
        int scriptLevel = equation.Display ? 0 : 1;
        if (!TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> sources) || sources.Count == 0)
        { fallback = "UnsupportedMathFractionSource"; return false; }
        double childNominal = sources[0].Run.Text.FontSize;
        double parentSize = equation.ParentFontSize ?? childNominal;
        if (!double.IsFinite(parentSize) || parentSize <= 0d || sources.Any(source => source.Run.Text.FontSize != childNominal))
        { fallback = "MixedMathFractionMetrics"; return false; }
        OpenTypeFont? metricOwner = parentMetricsFont ?? resolveFont(sources[0].Run)?.Font.Font;
        if (metricOwner is null || !metricOwner.TryGetMathConstants(out OpenTypeMathConstants constants))
        { fallback = "MissingMathFractionConstants"; return false; }
        if (!DocxMathRowLayout.TryMeasure(fraction.Numerator, resolveFont, token, out DocxMathBox? numerator, out fallback, scriptLevel,
                operatorSpacingFontSize: parentSize, scriptMetricsFont: metricOwner) ||
            !DocxMathRowLayout.TryMeasure(fraction.Denominator, resolveFont, token, out DocxMathBox? denominator, out fallback, scriptLevel,
                operatorSpacingFontSize: parentSize, scriptMetricsFont: metricOwner) ||
            numerator is null || denominator is null || numerator.Glyphs.Count == 0 || denominator.Glyphs.Count == 0)
        { fallback ??= "EmptyMathFraction"; return false; }
        DocxMathGlyphPlacement first = numerator.Glyphs[0];
        double size = first.FontSize;
        if (numerator.Glyphs.Concat(denominator.Glyphs).Any(glyph =>
            glyph.FontSize != size) || sources.Any(source => !source.Run.NormalText &&
                !ReferenceEquals(resolveFont(source.Run)?.Font.Font, metricOwner)))
        { fallback = "MixedMathFractionMetrics"; return false; }
        double scale = parentSize / metricOwner.UnitsPerEm;
        bool stack = fraction.TypeValue == "noBar";
        double axis = 0d, thickness = 0d, numeratorShift, denominatorShift;
        if (stack)
        {
            numeratorShift = constants.GetValue(equation.Display ? OpenTypeMathConstant.StackTopDisplayStyleShiftUp : OpenTypeMathConstant.StackTopShiftUp) * scale;
            denominatorShift = constants.GetValue(equation.Display ? OpenTypeMathConstant.StackBottomDisplayStyleShiftDown : OpenTypeMathConstant.StackBottomShiftDown) * scale;
            double gap = constants.GetValue(equation.Display ? OpenTypeMathConstant.StackDisplayStyleGapMin : OpenTypeMathConstant.StackGapMin) * scale;
            if (numeratorShift < 0d || denominatorShift < 0d || gap < 0d)
            { fallback = "InvalidMathStackConstants"; return false; }
            double missing = Math.Max(0d, gap - (numeratorShift + denominatorShift - numerator.BelowBaseline - denominator.AboveBaseline));
            numeratorShift += missing / 2d;
            denominatorShift += missing / 2d;
        }
        else
        {
            axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * scale;
            thickness = constants.GetValue(OpenTypeMathConstant.FractionRuleThickness) * scale;
            numeratorShift = constants.GetValue(equation.Display ? OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp : OpenTypeMathConstant.FractionNumeratorShiftUp) * scale;
            denominatorShift = constants.GetValue(equation.Display ? OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown : OpenTypeMathConstant.FractionDenominatorShiftDown) * scale;
            double numeratorGap = constants.GetValue(equation.Display ? OpenTypeMathConstant.FractionNumDisplayStyleGapMin : OpenTypeMathConstant.FractionNumeratorGapMin) * scale;
            double denominatorGap = constants.GetValue(equation.Display ? OpenTypeMathConstant.FractionDenomDisplayStyleGapMin : OpenTypeMathConstant.FractionDenominatorGapMin) * scale;
            if (thickness <= 0d || numeratorShift < 0d || denominatorShift < 0d || numeratorGap < 0d || denominatorGap < 0d)
            { fallback = "InvalidMathFractionConstants"; return false; }
            numeratorShift = Math.Max(numeratorShift, axis + thickness / 2d + numeratorGap + numerator.BelowBaseline);
            denominatorShift = Math.Max(denominatorShift, denominator.AboveBaseline - axis + thickness / 2d + denominatorGap);
        }
        double width = Math.Max(numerator.Width, denominator.Width);
        double above = Math.Max(numeratorShift + numerator.AboveBaseline, axis + thickness / 2d);
        double below = Math.Max(denominatorShift + denominator.BelowBaseline, -axis + thickness / 2d);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(thickness) || width <= 0d || above < 0d || below < 0d)
        { fallback = "InvalidMathFractionExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(numerator.Glyphs.Count + denominator.Glyphs.Count);
        Append(numerator, numeratorShift);
        Append(denominator, -denominatorShift);
        if (!stack) { OoxConversionBudget.Current?.ChargeSceneNodes(1); }
        box = new DocxMathBox(width, above, below, glyphs)
        { Rules = stack ? [] : [new DocxMathRulePlacement(0d, axis - thickness / 2d, width, thickness, first.Color)],
            ParentMetrics = sources.Any(source => source.Run.NormalText) ? new(metricOwner, parentSize) : null };
        fallback = null;
        return true;

        void Append(DocxMathBox child, double baseline)
        {
            double x = (width - child.Width) / 2d;
            foreach (DocxMathGlyphPlacement glyph in child.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                glyphs.Add(glyph with { X = glyph.X + x, BaselineY = glyph.BaselineY + baseline });
            }
        }
    }

    private static DocxMathFraction? GetFraction(DocxEquation equation) => GetSupportedFraction(GetSingleSupportedChild(equation));

    private static DocxMathFraction? GetSupportedFraction(DocxMathNode? node) =>
        node is DocxMathFraction { HasUnsupportedProperties: false, TypeValue: null or "bar" or "noBar" } fraction
            ? fraction : null;

    public static bool NeedsParentFractionMetrics(DocxEquation equation, IReadOnlyList<DocxMathRunSource> runs,
        CancellationToken token = default) =>
        (GetFraction(equation) is not null || GetFractionRow(equation, token) is not null) &&
        runs.Any(source => source.Run.NormalText);

    private static DocxMathNode? GetSingleSupportedChild(DocxEquation equation) =>
        equation.Content.HasSupportedProperties && equation.Content.Children.Count == 1
            ? equation.Content.Children[0] : null;
}
