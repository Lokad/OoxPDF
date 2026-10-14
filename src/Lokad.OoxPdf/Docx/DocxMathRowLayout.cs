using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

// First component slice: literal rows. Scripts/fractions retain typed roles and
// explicit fallback until they have their own independently qualified boxes.
internal static class DocxMathRowLayout
{
    private enum AtomKind { Ordinary, Binary, Relation, Open, Close, Punctuation }
    private sealed record Atom(DocxMathRun Run, Rune Source, Rune Render, AtomKind Kind);

    public static bool TryMeasure(DocxMathRow row, Func<DocxMathRun, PdfFontResource?> resolveFont,
        CancellationToken token, out DocxMathBox? box, out string? fallback)
    {
        box = null;
        if (!TryGetAtoms(row, token, out List<Atom> atoms, out fallback)) { return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(atoms.Count);
        double x = 0d, above = 0d, below = 0d, previousCorrection = 0d;
        AtomKind previous = AtomKind.Open;
        for (int i = 0; i < atoms.Count; i++)
        {
            token.ThrowIfCancellationRequested();
            Atom atom = atoms[i];
            PdfFontResource? resource = resolveFont(atom.Run);
            if (resource is null) { fallback = "MissingMathFont"; return false; }
            var font = resource.Font.Font;
            double size = atom.Run.Text.FontSize;
            if (!double.IsFinite(size) || size <= 0d) { fallback = "InvalidMathFontSize"; return false; }
            ushort glyph = font.MapCodePoint(atom.Render.Value);
            if (glyph == 0) { fallback = "MissingMathGlyph"; return false; }
            if (!resource.Font.TryGetEncodedCid(glyph, out ushort cid)) { fallback = "MissingMathEncoding"; return false; }
            if (!font.TryGetMathItalicCorrection(glyph, out short italic)) { fallback = "InvalidMathGlyphMetrics"; return false; }
            double scale = size / font.UnitsPerEm;
            double advance = font.GetAdvanceWidth(glyph) * scale;
            if (font.TryGetGlyphBounds(glyph, out var bounds))
            { above = Math.Max(above, bounds.YMax * scale); below = Math.Max(below, -bounds.YMin * scale); }
            else if (!Rune.IsWhiteSpace(atom.Render)) { fallback = "MissingMathGlyphBounds"; return false; }
            AtomKind kind = atom.Kind;
            if (kind == AtomKind.Binary && (i == 0 || previous is AtomKind.Open or AtomKind.Binary or AtomKind.Relation or AtomKind.Punctuation ||
                i + 1 == atoms.Count || atoms[i + 1].Kind is AtomKind.Close or AtomKind.Relation or AtomKind.Punctuation))
            { kind = AtomKind.Ordinary; }
            if (i != 0)
            {
                // Italic correction belongs to the previous glyph; operator glue
                // is a semantic math-unit rule, independent of font or fixture.
                if (kind != AtomKind.Ordinary) { x += previousCorrection; }
                x += Glue(previous, kind) * size / 18d;
            }
            RgbColor color = default;
            RgbColor.TryParse(atom.Run.Text.ColorHex, out color);
            glyphs.Add(new(atom.Source.Value, atom.Render.Value, resource, cid.ToString("X4", CultureInfo.InvariantCulture),
                size, x, 0d, advance, color));
            x += advance;
            previousCorrection = italic * scale;
            previous = kind;
        }
        x += previousCorrection;
        if (!double.IsFinite(x) || !double.IsFinite(above) || !double.IsFinite(below) || x < 0d)
        { fallback = "InvalidMathExtent"; return false; }
        box = new DocxMathBox(x, above, below, glyphs);
        fallback = null;
        return true;
    }

    public static bool TryGetRenderCodePoints(DocxMathRow row, CancellationToken token, out IReadOnlyList<int> codePoints, out string? fallback)
    {
        bool valid = TryGetAtoms(row, token, out List<Atom> atoms, out fallback);
        codePoints = valid ? atoms.Select(atom => atom.Render.Value).Distinct().ToArray() : [];
        return valid;
    }

    private static bool TryGetAtoms(DocxMathRow row, CancellationToken token, out List<Atom> atoms, out string? fallback)
    {
        atoms = []; fallback = null;
        var pending = new Stack<DocxMathNode>(); pending.Push(row);
        while (pending.TryPop(out DocxMathNode? node))
        {
            token.ThrowIfCancellationRequested();
            if (node is DocxMathRow sequence && sequence.PropertiesXml.Count == 0)
            {
                for (int i = sequence.Children.Count - 1; i >= 0; i--) { pending.Push(sequence.Children[i]); }
            }
            else if (node is DocxMathRun run && HasSupportedStyle(run))
            {
                foreach (Rune rune in run.Text.Text.EnumerateRunes())
                {
                    token.ThrowIfCancellationRequested();
                    int rendered = run.MathStyleValue == "p" ? rune.Value : ItalicCodePoint(rune.Value);
                    if (rendered < 0) { fallback = "UnsupportedMathAlphabet"; return false; }
                    OoxConversionBudget.Current?.ChargeSceneNodes(1);
                    atoms.Add(new(run, rune, new Rune(rendered), Kind(rune.Value)));
                }
            }
            else { fallback = node is DocxMathUnsupported unsupported ? "UnsupportedMathNode:" + unsupported.ElementName : "UnsupportedMathStructureOrStyle"; return false; }
        }
        return true;
    }

    private static bool HasSupportedStyle(DocxMathRun run) =>
        !run.HasUnsupportedProperties && run.MathStyleValue is null or "i" or "p" && run.ScriptValue is null or "roman" &&
        !run.Text.Bold && !run.Text.Underline && !run.Text.Strike && !run.Text.DoubleStrike && !run.Text.Hidden &&
        !run.Text.AllCaps && !run.Text.SmallCaps && run.Text.VerticalAlignmentValue is null && run.Text.CharacterSpacingPoints == 0d &&
        run.Text.HighlightValue is null && run.Text.ShadingFillHex is null;

    private static int ItalicCodePoint(int codePoint)
    {
        if (codePoint is >= 'A' and <= 'Z') { return 0x1D434 + codePoint - 'A'; }
        if (codePoint is >= 'a' and <= 'z') { return codePoint == 'h' ? 0x210E : 0x1D44E + codePoint - 'a'; }
        if (codePoint is >= 0x0391 and <= 0x03A9 && codePoint != 0x03A2) { return 0x1D6E2 + codePoint - 0x0391; }
        if (codePoint is >= 0x03B1 and <= 0x03C9) { return 0x1D6FC + codePoint - 0x03B1; }
        if (Rune.IsLetter(new Rune(codePoint)) && codePoint is not (>= 0x1D400 and <= 0x1D7FF) && codePoint != 0x210E) { return -1; }
        return codePoint;
    }

    private static AtomKind Kind(int codePoint) => codePoint switch
    {
        '+' or '-' or 0x2212 or 0x00D7 or 0x00F7 or 0x22C5 => AtomKind.Binary,
        '=' or '<' or '>' or 0x2264 or 0x2265 or 0x2260 or 0x2248 => AtomKind.Relation,
        '(' or '[' or '{' => AtomKind.Open,
        ')' or ']' or '}' => AtomKind.Close,
        ',' or ';' => AtomKind.Punctuation,
        _ => AtomKind.Ordinary
    };
    private static int Glue(AtomKind left, AtomKind right) =>
        left == AtomKind.Open || right == AtomKind.Close ? 0 :
        left == AtomKind.Relation || right == AtomKind.Relation ? 5 :
        left == AtomKind.Binary || right == AtomKind.Binary ? 4 :
        left == AtomKind.Punctuation ? 3 : 0;
}
