using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private sealed record StructuredPart(DocxMathRow Before, DocxMathNode Structure);
    private sealed record StructuredRow(IReadOnlyList<StructuredPart> Parts, DocxMathRow After);

    private static DocxMathNode? GetSupportedStructure(DocxMathNode node, DocxEquation equation, CancellationToken token) =>
        (DocxMathNode?)GetSupportedFraction(node) ?? (DocxMathNode?)GetSupportedScripts(node) ??
        (DocxMathNode?)GetSupportedRadical(node) ?? (DocxMathNode?)GetSupportedDelimiter(node) ?? GetSupportedNary(node, equation, token);

    private static StructuredRow? GetStructuredRow(DocxEquation equation, CancellationToken token)
    {
        var literals = new List<DocxMathNode>();
        var parts = new List<StructuredPart>();
        var pending = new Stack<DocxMathNode>(); pending.Push(equation.Content);
        while (pending.TryPop(out DocxMathNode? node))
        {
            token.ThrowIfCancellationRequested();
            if (node is DocxMathRow row && row.HasSupportedProperties)
            {
                for (int i = row.Children.Count - 1; i >= 0; i--) { pending.Push(row.Children[i]); }
            }
            else if (node is DocxMathRun) { literals.Add(node); }
            else if (GetSupportedStructure(node, equation, token) is { } supported)
            { parts.Add(new(new(literals), supported)); literals = []; }
            else { return null; }
        }
        // A lone structure is measured by its existing owner, so child
        // delegation below cannot recurse into this same sequence grammar.
        return parts.Count > 1 || parts.Count == 1 && parts[0].Before.Children.Count + literals.Count > 0
            ? new(parts, new(literals)) : null;
    }

    private static bool TryMeasureStructuredRow(DocxEquation equation, StructuredRow row,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, OpenTypeFont? parentMetricsFont,
        Func<IDocxMathControl, PdfFontResource?>? resolveControlFont,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics)
    {
        box = null;
        var glyphs = new List<DocxMathGlyphPlacement>();
        var rules = new List<DocxMathRulePlacement>();
        double advance = 0d, width = 0d, above = 0d, below = 0d;
        DocxMathMetricOwner? leadingOwner = null;
        bool afterStructure = false;
        var previousEnd = DocxMathRowLayout.AtomKind.Ordinary;
        foreach (StructuredPart part in row.Parts)
        {
            token.ThrowIfCancellationRequested();
            DocxEquation childEquation = equation with { Content = new([part.Structure]) };
            if (!TryGetRuns(childEquation, token, out var childRuns) || childRuns.Count == 0)
            { fallback = "UnsupportedMathStructuredRow"; return false; }
            double parentSize = equation.ParentFontSize ?? childRuns[0].Run.Text.FontSize;
            if (!TryMeasure(childEquation, resolveFont, token, out var structure, out fallback, parentMetricsFont, resolveControlFont, resolveControlMetrics) || structure is null)
            { fallback ??= "UnsupportedMathStructuredRow"; return false; }
            var first = structure.Edges?.First ?? DocxMathRowLayout.AtomKind.Ordinary;
            if (!DocxMathRowLayout.TryMeasure(part.Before, resolveFont, token, out var before, out fallback,
                    neighbors: new(Before: afterStructure ? previousEnd : null, After: first)) || before is null)
            { fallback ??= "UnsupportedMathStructuredRow"; return false; }
            AppendLiteral(before);
            // Equation metrics own the gap before each structured atom;
            // child glyph size continues to own its internal geometry.
            if (before.Edges is { } left)
                advance += DocxMathRowLayout.Glue(left.Last, first) * parentSize / 18d;
            else if (afterStructure)
                advance += DocxMathRowLayout.Glue(previousEnd, first) * parentSize / 18d;
            Append(structure, advance); advance += structure.AdvanceWidth;
            afterStructure = true;
            previousEnd = structure.Edges?.Last ?? DocxMathRowLayout.AtomKind.Ordinary;
        }
        if (!DocxMathRowLayout.TryMeasure(row.After, resolveFont, token, out var after, out fallback,
            neighbors: new(Before: previousEnd)) || after is null) { return false; }
        AppendLiteral(after);
        width = Math.Max(width, advance);
        if (!double.IsFinite(width) || width < 0d || !double.IsFinite(above) || !double.IsFinite(below))
        { fallback = "InvalidMathStructuredRowExtent"; return false; }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = leadingOwner };
        fallback = null;
        return true;

        void Append(DocxMathBox child, double x)
        {
            width = Math.Max(width, x + child.Width);
            above = Math.Max(above, child.AboveBaseline); below = Math.Max(below, child.BelowBaseline);
            if (child.ParentMetrics is { } owner && (leadingOwner is null || Leading(owner) > Leading(leadingOwner)))
                leadingOwner = owner;
            foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + x }); }
            foreach (var rule in child.Rules) { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + x }); }
        }

        void AppendLiteral(DocxMathBox literal)
        {
            if (afterStructure && literal.Edges is { FirstNormalText: false } right)
                advance += DocxMathRowLayout.Glue(previousEnd, right.First) * literal.Glyphs[0].FontSize / 18d;
            Append(literal, advance); advance += literal.AdvanceWidth;
        }

        static double Leading(DocxMathMetricOwner owner) => owner.Font.TryGetMathConstants(out var constants)
            ? constants.GetValue(OpenTypeMathConstant.MathLeading) * owner.Size / owner.Font.UnitsPerEm : 0d;
    }
}
