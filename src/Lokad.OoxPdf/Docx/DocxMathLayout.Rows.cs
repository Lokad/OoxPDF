using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private sealed record StructuredRow(DocxMathRow Before, DocxMathNode Structure, DocxMathRow After);

    private static DocxMathNode? GetSupportedStructure(DocxMathNode node) =>
        (DocxMathNode?)GetSupportedFraction(node) ?? GetSupportedScripts(node);

    private static StructuredRow? GetStructuredRow(DocxEquation equation, CancellationToken token)
    {
        var before = new List<DocxMathNode>();
        var after = new List<DocxMathNode>();
        DocxMathNode? structure = null;
        var pending = new Stack<DocxMathNode>(); pending.Push(equation.Content);
        while (pending.TryPop(out DocxMathNode? node))
        {
            token.ThrowIfCancellationRequested();
            if (node is DocxMathRow row && row.HasSupportedProperties)
            {
                for (int i = row.Children.Count - 1; i >= 0; i--) { pending.Push(row.Children[i]); }
            }
            else if (node is DocxMathRun) { (structure is null ? before : after).Add(node); }
            else if (structure is null && GetSupportedStructure(node) is { } supported)
            { structure = supported; }
            else { return null; }
        }
        return structure is not null && before.Count + after.Count > 0 ? new(new(before), structure, new(after)) : null;
    }

    private static bool TryMeasureStructuredRow(DocxEquation equation, StructuredRow row,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, OpenTypeFont? parentMetricsFont)
    {
        box = null;
        DocxEquation childEquation = equation with { Content = new([row.Structure]) };
        if (!TryGetRuns(childEquation, token, out var childRuns) || childRuns.Count == 0)
        { fallback = "UnsupportedMathStructuredRow"; return false; }
        // The surrounding equation owns glue before a structured atom;
        // its authored child glyph size continues to own the child's geometry.
        double parentSize = equation.ParentFontSize ?? childRuns[0].Run.Text.FontSize;
        if (!TryMeasure(childEquation, resolveFont, token, out var structure, out fallback, parentMetricsFont) || structure is null ||
            !DocxMathRowLayout.TryMeasure(row.Before, resolveFont, token, out var before, out fallback,
                neighbors: new(After: DocxMathRowLayout.AtomKind.Ordinary)) || before is null ||
            !DocxMathRowLayout.TryMeasure(row.After, resolveFont, token, out var after, out fallback,
                neighbors: new(Before: DocxMathRowLayout.AtomKind.Ordinary)) || after is null)
        { fallback ??= "UnsupportedMathStructuredRow"; return false; }
        double structureX = before.AdvanceWidth;
        if (before.Edges is { } left) { structureX += DocxMathRowLayout.Glue(left.Last, DocxMathRowLayout.AtomKind.Ordinary) * parentSize / 18d; }
        double afterX = structureX + structure.AdvanceWidth;
        if (after.Edges is { FirstNormalText: false } right)
            afterX += DocxMathRowLayout.Glue(DocxMathRowLayout.AtomKind.Ordinary, right.First) * after.Glyphs[0].FontSize / 18d;
        double advance = afterX + after.AdvanceWidth;
        double width = Math.Max(advance, Math.Max(before.Width, Math.Max(structureX + structure.Width, afterX + after.Width)));
        double above = Math.Max(before.AboveBaseline, Math.Max(structure.AboveBaseline, after.AboveBaseline));
        double below = Math.Max(before.BelowBaseline, Math.Max(structure.BelowBaseline, after.BelowBaseline));
        if (!double.IsFinite(width) || width < 0d || !double.IsFinite(above) || !double.IsFinite(below))
        { fallback = "InvalidMathStructuredRowExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(before.Glyphs.Count + structure.Glyphs.Count + after.Glyphs.Count);
        var rules = new List<DocxMathRulePlacement>(structure.Rules.Count);
        Append(before, 0d); Append(structure, structureX); Append(after, afterX);
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = structure.ParentMetrics };
        fallback = null;
        return true;

        void Append(DocxMathBox child, double x)
        {
            foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + x }); }
            foreach (var rule in child.Rules) { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + x }); }
        }
    }
}
