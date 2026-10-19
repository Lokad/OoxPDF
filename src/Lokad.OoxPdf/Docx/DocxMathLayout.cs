using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

// Shared discovery retains painted source runs and ordered control owners.
// Nested metric-only controls do not duplicate source runs or PDF resources.
internal sealed record DocxMathRunSource(DocxMathRun Run, int ScriptLevel)
{
    public IReadOnlyList<IDocxMathControl> ControlOwners { get; init; } = [];
    public IDocxMathControl? ControlOwner => ControlOwners.Count == 0 ? null : ControlOwners[0];
}

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
            if (box.ParentMetrics is null)
                above = Math.Max(above, DocxLineMetrics.MeasureTypographicBaselineInset(font, glyph.FontSize));
            if (box.ParentMetrics is null)
                below = Math.Max(below, -font.Os2.TypographicDescender * glyph.FontSize / font.UnitsPerEm);
            naturalLine = Math.Max(naturalLine, DocxLineMetrics.MeasureTypographicLineHeight(font, glyph.FontSize));
            if (box.ParentMetrics is null && font.TryGetMathConstants(out OpenTypeMathConstants constants))
                leading = Math.Max(leading, constants.GetValue(OpenTypeMathConstant.MathLeading) * glyph.FontSize / font.UnitsPerEm);
        }
        if (box.ParentMetrics is { } parent)
        {
            // An explicit control owns the ordinary baseline floor. Child ink
            // remains in the box extent; a larger child font must not add its
            // unrelated typographic line gap above that occupied frame.
            above = Math.Max(above, DocxLineMetrics.MeasureTypographicBaselineInset(parent.Font,
                OfficePdfTextEmissionProfile.FontSize(parent.Size)));
            if (parent.IncludeTypographicFrame)
            {
                double parentGlyphSize = OfficePdfTextEmissionProfile.FontSize(parent.Size);
                naturalLine = Math.Max(naturalLine, DocxLineMetrics.MeasureTypographicLineHeight(parent.Font, parentGlyphSize));
            }
            // The same control owns the ordinary descent floor. Larger child
            // ink remains in the box; an unused child-font descender does not.
            below = Math.Max(below, -parent.Font.Os2.TypographicDescender *
                OfficePdfTextEmissionProfile.FontSize(parent.Size) / parent.Font.UnitsPerEm);
            if (parent.Font.TryGetMathConstants(out OpenTypeMathConstants parentConstants))
                leading = Math.Max(0d, parentConstants.GetValue(OpenTypeMathConstant.MathLeading) * parent.Size / parent.Font.UnitsPerEm);
        }
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
        if (GetFunction(equation) is { } function) { return TryGetFunctionRuns(equation, function, token, out runs); }
        if (GetDelimiter(equation) is { } delimiter) { return TryGetDelimiterRuns(equation, delimiter, token, out runs); }
        var result = new List<DocxMathRunSource>();
        IDocxMathControl? controlOwner = null;
        if (GetMatrix(equation, token) is { } matrix)
        {
            foreach (var row in matrix.Rows)
            foreach (var cell in row.Cells)
            {
                token.ThrowIfCancellationRequested();
                if (IsEmptyMatrixCell(cell, token)) { continue; }
                if (!AddBoxArgumentRows(cell))
                { runs = []; return false; }
            }
            controlOwner = matrix;
        }
        else if (GetScripts(equation) is { } scripts)
        {
            if (!AddScriptRows(scripts)) { runs = []; return false; }
        }
        else if (GetFraction(equation) is { } fraction)
        {
            if (!AddFractionRows(fraction, allowScripts: true))
            { runs = []; return false; }
        }
        else if (GetNary(equation, token) is { } nary)
        {
            if (!AddBoxArgumentRows(nary.Body, allowMatrix: true)) { runs = []; return false; }
            foreach (var (argument, hidden) in new[] { (nary.Subscript, nary.SubscriptHiddenValue), (nary.Superscript, nary.SuperscriptHiddenValue) })
            {
                if (hidden == true) { continue; }
                if (!DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level)) { runs = []; return false; }
                if (!AddRow(argument, level, true)) { runs = []; return false; }
            }
            controlOwner = nary;
        }
        else if (GetRadical(equation) is { } radical)
        {
            if (!AddBoxArgumentRows(radical.Radicand, allowMatrix: true)) { runs = []; return false; }
            controlOwner = radical;
            if (radical.DegreeHidden != true)
            {
                if (!DocxMathArgumentMetrics.TryGetScriptLevel(radical.Degree, out int level, startingLevel: 2))
                { runs = []; return false; }
                if (!AddRow(radical.Degree, level, true)) { runs = []; return false; }
            }
        }
        else if (GetStructuredRow(equation, token) is { } compound)
        {
            var sources = new List<DocxMathRunSource>();
            foreach (StructuredPart part in compound.Parts)
            {
                token.ThrowIfCancellationRequested();
                if (!TryGetRuns(equation with { Content = part.Before }, token, out var before) ||
                    !TryGetRuns(equation with { Content = new([part.Structure]) }, token, out var child))
                { runs = []; return false; }
                sources.AddRange(before); sources.AddRange(child);
            }
            if (!TryGetRuns(equation with { Content = compound.After }, token, out var after)) { runs = []; return false; }
            sources.AddRange(after); runs = sources;
            return true;
        }
        else if (!AddRow(equation.Content, 0, false)) { runs = []; return false; }
        if (controlOwner is not null && result.Count > 0) { result[0] = result[0] with { ControlOwners = [controlOwner, .. result[0].ControlOwners] }; }
        runs = result;
        return true;

        bool AddBoxArgumentRows(DocxMathRow argument, bool allowMatrix = false)
        {
            var child = equation with { Content = argument };
            // Matrix cells share mixed fraction/script rows with outer boxes,
            // but cannot admit another grid. Append each child's sources now so
            // mixed and ordinary cells retain their complete authored order.
            if (allowMatrix && GetMatrix(child, token) is not null || HasMixedBoxArgument(child, token, allowMatrix))
            {
                if (!TryGetRuns(child, token, out var sources)) return false;
                result.AddRange(sources);
                return true;
            }
            return GetScripts(child) is { } childScripts ? AddScriptRows(childScripts) : AddArgumentRows(argument);
        }

        bool AddFractionRows(DocxMathFraction fraction, bool allowScripts = false)
        {
            int first = result.Count;
            bool skewed = fraction.TypeValue == "skw";
            int level = skewed || equation.Display ? 0 : 1;
            if (!AddArgument(fraction.Numerator) || !AddArgument(fraction.Denominator)) return false;
            if (skewed && result.Count > first)
                result[first] = result[first] with { ControlOwners = [fraction, .. result[first].ControlOwners] };
            return true;

            bool AddArgument(DocxMathRow argument) => allowScripts && !skewed && GetScriptArgument(argument) is { } child
                ? AddScriptRows(child, level, literalBase: true) : AddRow(argument, level, false);
        }

        bool AddScriptRows(DocxMathScripts script, int startingLevel = 0, bool literalBase = false)
        {
            if (!(literalBase ? AddRow(script.Base, startingLevel, false) : AddArgumentRows(script.Base))) return false;
            foreach (DocxMathRow? argument in new[] { script.Subscript, script.Superscript })
            {
                if (argument is null) { continue; }
                if (!DocxMathArgumentMetrics.TryGetScriptLevel(argument, out int level, Math.Min(2, startingLevel + 1))) { return false; }
                if (!AddRow(argument, level, true)) return false;
            }
            return true;
        }

        bool AddArgumentRows(DocxMathRow argument)
        {
            // These argument roles admit a literal row or one fraction/stack with
            // literal children. Discovery never recursively expands structures.
            if (GetFraction(equation with { Content = argument }) is { } fraction)
            {
                return AddFractionRows(fraction);
            }
            return AddRow(argument, 0, false);
        }

        bool AddRow(DocxMathRow row, int level, bool allowArgumentSize)
        {
            if (!DocxMathRowLayout.TryGetRenderCodePoints(row, token, out _, out _, allowArgumentSize)) return false;
            var pending = new Stack<DocxMathNode>();
            pending.Push(row);
            while (pending.TryPop(out DocxMathNode? node))
            {
                token.ThrowIfCancellationRequested();
                if (node is DocxMathRun run) result.Add(new(run, level));
                else if (node is DocxMathRow sequence)
                    for (int index = sequence.Children.Count - 1; index >= 0; index--) pending.Push(sequence.Children[index]);
            }
            return true;
        }
    }

    public static bool TryMeasure(DocxEquation equation, Func<DocxMathRun, PdfFontResource?> resolveFont,
        CancellationToken token, out DocxMathBox? box, out string? fallback, OpenTypeFont? parentMetricsFont = null,
        Func<IDocxMathControl, PdfFontResource?>? resolveControlFont = null,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics = null, double? rowControlSize = null)
    {
        if (GetMatrix(equation, token) is { } matrix)
            return TryMeasureMatrix(equation, matrix, resolveFont, token, out box, out fallback, resolveControlFont, resolveControlMetrics, rowControlSize);
        if (GetFunction(equation) is { } function)
            return TryMeasureFunction(equation, function, resolveFont, token, out box, out fallback, resolveControlFont, resolveControlMetrics, rowControlSize);
        if (GetScripts(equation) is { } scripts)
            return TryMeasureScripts(equation, scripts, resolveFont, token, out box, out fallback, resolveControlMetrics);
        if (GetNary(equation, token) is { } nary)
            return TryMeasureNary(equation, nary, resolveFont, token, out box, out fallback, resolveControlFont, resolveControlMetrics);
        if (GetDelimiter(equation) is { } delimiter)
            return TryMeasureDelimiter(equation, delimiter, resolveFont, token, out box, out fallback, resolveControlFont, resolveControlMetrics);
        if (GetRadical(equation) is { } radical)
            return TryMeasureRadical(equation, radical, resolveFont, token, out box, out fallback, resolveControlFont, resolveControlMetrics);
        if (GetStructuredRow(equation, token) is { } compound)
            return TryMeasureStructuredRow(equation, compound, resolveFont, token, out box, out fallback, parentMetricsFont, resolveControlFont, resolveControlMetrics, rowControlSize);
        if (GetFraction(equation) is not { } fraction)
            return DocxMathRowLayout.TryMeasure(equation.Content, resolveFont, token, out box, out fallback);
        if (fraction.TypeValue == "skw")
            return TryMeasureSkewedFraction(equation, fraction, resolveFont, token, out box, out fallback, resolveControlFont);
        if (IsEmptyFraction(equation))
            return TryMeasureEmptyFraction(equation, fraction, parentMetricsFont ?? resolveControlMetrics?.Invoke(fraction),
                token, out box, out fallback);
        box = null;
        int scriptLevel = equation.Display ? 0 : 1;
        if (!TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> sources))
        { fallback = "UnsupportedMathFractionSource"; return false; }
        var firstSource = sources.FirstOrDefault(source => source.Run.Text.Text.Length != 0);
        bool metricOnly = firstSource is null;
        if (metricOnly && (fraction.TypeValue == "noBar" || fraction.ControlText is null ||
            !DocxMathRowLayout.HasSupportedTextFormatting(fraction.ControlText, allowBold: false)))
        { fallback = "UnsupportedEmptyFractionControl"; return false; }
        double childNominal = firstSource?.Run.Text.FontSize ?? ResolveControlSize(fraction, equation, sources);
        double parentSize = metricOnly ? childNominal : equation.ParentFontSize ?? childNominal;
        if (!double.IsFinite(parentSize) || parentSize <= 0d || sources.Any(source =>
            source.Run.Text.Text.Length != 0 && source.Run.Text.FontSize != childNominal))
        { fallback = "MixedMathFractionMetrics"; return false; }
        OpenTypeFont? metricOwner = parentMetricsFont ?? (metricOnly ? resolveControlMetrics?.Invoke(fraction) :
            resolveFont(firstSource!.Run)?.Font.Font);
        if (metricOwner is null || !metricOwner.TryGetMathConstants(out OpenTypeMathConstants constants))
        { fallback = "MissingMathFractionConstants"; return false; }
        if (!TryMeasureFractionArgument(fraction.Numerator, denominator: false, out DocxMathBox? numerator, out fallback, out bool emptyNumerator) ||
            !TryMeasureFractionArgument(fraction.Denominator, denominator: true, out DocxMathBox? denominator, out fallback, out bool emptyDenominator) ||
            numerator is null || denominator is null)
        { fallback ??= "EmptyMathFraction"; return false; }
        bool hasEmptyArgument = emptyNumerator || emptyDenominator;
        if (hasEmptyArgument && (fraction.TypeValue == "noBar" || fraction.ControlText is null ||
            !DocxMathRowLayout.HasSupportedTextFormatting(fraction.ControlText, allowBold: false)))
        { fallback = "UnsupportedEmptyFractionControl"; return false; }
        DocxMathGlyphPlacement? first = numerator.Glyphs.FirstOrDefault() ?? denominator.Glyphs.FirstOrDefault();
        // Uniform source nominal size was checked above. Distinct emitted
        // sizes now belong to the supported argument/script roles.
        if (sources.Any(source => source.Run.Text.Text.Length != 0 && !source.Run.NormalText &&
                !ReferenceEquals(resolveFont(source.Run)?.Font.Font, metricOwner)))
        { fallback = "MixedMathFractionMetrics"; return false; }
        var ruleColor = first?.Color ?? default;
        if (hasEmptyArgument || metricOnly) RgbColor.TryParse(fraction.ControlText!.ColorHex, out ruleColor);
        // An authored control owns fraction geometry; argument runs retain
        // their own script sizes and spacing. An unowned control keeps the
        // existing parent/child metric inheritance.
        double controlSize = fraction.ControlFontSize ?? parentSize;
        if (!double.IsFinite(controlSize) || controlSize <= 0d)
        { fallback = "InvalidMathFractionControlSize"; return false; }
        bool ownsControlMetrics = fraction.ControlFontSize is not null && controlSize != parentSize;
        return TryComposeFraction(equation, fraction, metricOwner, constants, controlSize, numerator, denominator,
            ruleColor, ownsControlMetrics || hasEmptyArgument || metricOnly ||
                sources.Any(source => source.Run.Text.Text.Length != 0 && source.Run.NormalText),
            token, out box, out fallback);

        bool TryMeasureFractionArgument(DocxMathRow argument, bool denominator, out DocxMathBox? child,
            out string? reason, out bool empty)
        {
            var childEquation = equation with { Content = argument, Cramped = equation.Cramped || denominator };
            if (GetScripts(childEquation) is { } childScripts)
            {
                empty = false;
                return TryMeasureScripts(childEquation, childScripts, resolveFont, token, out child, out reason,
                    resolveControlMetrics, startingLevel: scriptLevel);
            }
            return TryMeasureLiteralArgument(argument, resolveFont, metricOwner, parentSize, scriptLevel, token,
                out child, out reason, out empty, operatorSpacingFontSize: parentSize);
        }
    }

    private static bool TryComposeFraction(DocxEquation equation, DocxMathFraction fraction, OpenTypeFont metricOwner,
        OpenTypeMathConstants constants, double parentSize, DocxMathBox numerator, DocxMathBox denominator,
        RgbColor ruleColor, bool useParentMetrics, CancellationToken token, out DocxMathBox? box, out string? fallback)
    {
        box = null;
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
        box = new DocxMathBox(width, above, below, glyphs)
        { Rules = stack ? [] : [new DocxMathRulePlacement(0d, axis - thickness / 2d, width, thickness, ruleColor)],
            EnclosureHeight = above + below,
            Edges = new(DocxMathRowLayout.AtomKind.Inner, DocxMathRowLayout.AtomKind.Inner, false),
            ParentMetrics = useParentMetrics ? new(metricOwner, parentSize) : null };
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
        node is DocxMathFraction { HasUnsupportedProperties: false, TypeValue: null or "bar" or "noBar" or "skw" } fraction
            ? fraction : null;

    public static bool NeedsParentFractionMetrics(DocxEquation equation, IReadOnlyList<DocxMathRunSource> runs,
        CancellationToken token = default) =>
        runs.Any(source => source.Run.Text.Text.Length != 0 && source.Run.NormalText) &&
            (GetFraction(equation) is not null || GetStructuredRow(equation, token)?.Parts.Any(part => part.Structure is DocxMathFraction) == true);

    private static DocxMathNode? GetSingleSupportedChild(DocxEquation equation) =>
        equation.Content.HasSupportedProperties && equation.Content.Children.Count == 1
            ? equation.Content.Children[0] : null;
}
