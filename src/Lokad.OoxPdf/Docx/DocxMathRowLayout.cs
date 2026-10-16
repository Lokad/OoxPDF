using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

// First component slice: literal rows. Scripts/fractions retain typed roles and
// explicit fallback until they have their own independently qualified boxes.
internal static class DocxMathRowLayout
{
    internal enum AtomKind { Ordinary, Binary, Relation, Open, Close, Punctuation, LargeOperator, Inner }
    internal readonly record struct Neighbors(AtomKind? Before = null, AtomKind? After = null);
    internal sealed record RowEdges(AtomKind First, AtomKind Last, bool FirstNormalText);
    private sealed record Atom(DocxMathRun Run, Rune Source, Rune Render, AtomKind Kind);

    public static bool TryMeasure(DocxMathRow row, Func<DocxMathRun, PdfFontResource?> resolveFont,
        CancellationToken token, out DocxMathBox? box, out string? fallback, int scriptLevel = 0, bool exportFontSize = false,
        bool allowArgumentSize = false, double? operatorSpacingFontSize = null,
        Lokad.OoxPdf.Fonts.OpenTypeFont? scriptMetricsFont = null, Neighbors neighbors = default)
    {
        box = null;
        token.ThrowIfCancellationRequested();
        if (operatorSpacingFontSize is { } spacingSize && (!double.IsFinite(spacingSize) || spacingSize <= 0d))
        { fallback = "InvalidMathOperatorSpacingSize"; return false; }
        if (!TryGetAtoms(row, token, out List<Atom> atoms, out fallback, allowArgumentSize)) { return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(atoms.Count);
        var inlineBreaks = new List<int>();
        Stack<int>? delimiters = null;
        bool balancedDelimiters = true;
        double x = 0d, above = 0d, below = 0d, previousCorrection = 0d, slantedInkRight = 0d;
        AtomKind previous = neighbors.Before ?? AtomKind.Open, firstKind = AtomKind.Ordinary;
        for (int i = 0; i < atoms.Count; i++)
        {
            token.ThrowIfCancellationRequested();
            Atom atom = atoms[i];
            PdfFontResource? resource = resolveFont(atom.Run);
            if (resource is null) { fallback = "MissingMathFont"; return false; }
            var font = resource.Font.Font;
            double size = atom.Run.Text.FontSize;
            if (scriptLevel != 0)
            {
                if (!DocxMathArgumentMetrics.TryGetScriptFontSize(scriptMetricsFont ?? font, size, scriptLevel, out size))
                { fallback = "MissingMathScriptScale"; return false; }
            }
            else if (exportFontSize) { size = OfficePdfTextEmissionProfile.FontSize(size); }
            if (!double.IsFinite(size) || size <= 0d) { fallback = "InvalidMathFontSize"; return false; }
            ushort glyph = font.MapCodePoint(atom.Render.Value);
            if (scriptLevel != 0 && !atom.Run.NormalText && !font.TryGetMathScriptGlyph(glyph, scriptLevel, token, out glyph))
            { fallback = "UnsupportedMathScriptGlyph"; return false; }
            if (glyph == 0) { fallback = "MissingMathGlyph"; return false; }
            if (!resource.Font.TryGetEncodedCid(glyph, out ushort cid)) { fallback = "MissingMathEncoding"; return false; }
            if (!font.TryGetMathItalicCorrection(glyph, out short italic)) { fallback = "InvalidMathGlyphMetrics"; return false; }
            double scale = size / font.UnitsPerEm;
            double advance = font.GetAdvanceWidth(glyph) * scale;
            double shear = atom.Run.NormalText && atom.Run.Text.Italic && Math.Abs(font.Post.ItalicAngle) < 0.01d ? 1d / 3d : 0d;
            if (font.TryGetGlyphBounds(glyph, out var bounds))
            { above = Math.Max(above, bounds.YMax * scale); below = Math.Max(below, -bounds.YMin * scale); }
            else if (!Rune.IsWhiteSpace(atom.Render)) { fallback = "MissingMathGlyphBounds"; return false; }
            AtomKind kind = atom.Kind;
            AtomKind? next = i + 1 == atoms.Count ? neighbors.After : atoms[i + 1].Kind;
            if (kind == AtomKind.Binary && (i == 0 && neighbors.Before is null ||
                previous is AtomKind.Open or AtomKind.Binary or AtomKind.Relation or AtomKind.Punctuation ||
                next is null or AtomKind.Close or AtomKind.Relation or AtomKind.Punctuation))
            { kind = AtomKind.Ordinary; }
            if (i == 0) { firstKind = kind; }
            if (kind == AtomKind.Open) (delimiters ??= new()).Push(atom.Source.Value);
            if (kind == AtomKind.Close)
            {
                int expected = atom.Source.Value switch { ')' => '(', ']' => '[', '}' => '{', _ => -1 };
                if (delimiters is null || !delimiters.TryPop(out int opening) || opening != expected)
                    balancedDelimiters = false;
            }
            if (scriptLevel == 0 && !atom.Run.NormalText &&
                kind is AtomKind.Binary or AtomKind.Relation && i > 0 && i + 1 < atoms.Count)
                inlineBreaks.Add(i + 1);
            if (i != 0)
            {
                // Italic correction belongs to the previous glyph; operator glue
                // is a semantic math-unit rule, independent of font or fixture.
                if (!atom.Run.NormalText && kind != AtomKind.Ordinary) { x += previousCorrection; }
                if (!atom.Run.NormalText && scriptLevel == 0) { x += Glue(previous, kind) * (operatorSpacingFontSize ?? size) / 18d; }
            }
            RgbColor color = default;
            RgbColor.TryParse(atom.Run.Text.ColorHex, out color);
            glyphs.Add(new(atom.Source.Value, atom.Render.Value, resource, cid.ToString("X4", CultureInfo.InvariantCulture),
                size, x, 0d, advance, color) { OriginalGlyph = glyph,
                    TextShear = shear });
            if (shear != 0d && font.TryGetGlyphBounds(glyph, out var slantedBounds))
                slantedInkRight = Math.Max(slantedInkRight, x + (slantedBounds.XMax + shear * slantedBounds.YMax) * scale);
            x += advance;
            previousCorrection = atom.Run.NormalText ? 0d : italic * scale;
            previous = kind;
        }
        if (neighbors.After is not AtomKind.Ordinary) { x += previousCorrection; }
        double advanceWidth = x;
        x = Math.Max(x, slantedInkRight);
        if (!double.IsFinite(x) || !double.IsFinite(above) || !double.IsFinite(below) || x < 0d)
        { fallback = "InvalidMathExtent"; return false; }
        box = new DocxMathBox(x, above, below, glyphs)
        { LogicalAdvance = advanceWidth == x ? null : advanceWidth,
            InlineBreakAfterGlyphs = balancedDelimiters && (delimiters is null || delimiters.Count == 0) ? inlineBreaks : [],
            Edges = atoms.Count == 0 ? null : new(firstKind, previous, atoms[0].Run.NormalText) };
        fallback = null;
        return true;
    }

    public static bool TryGetRenderCodePoints(DocxMathRow row, CancellationToken token, out IReadOnlyList<int> codePoints, out string? fallback,
        bool allowArgumentSize = false)
    {
        bool valid = TryGetAtoms(row, token, out List<Atom> atoms, out fallback, allowArgumentSize);
        codePoints = valid ? atoms.Select(atom => atom.Render.Value).Distinct().ToArray() : [];
        return valid;
    }

    private static bool TryGetAtoms(DocxMathRow row, CancellationToken token, out List<Atom> atoms, out string? fallback,
        bool allowArgumentSize = false)
    {
        atoms = []; fallback = null;
        var pending = new Stack<DocxMathNode>(); pending.Push(row);
        while (pending.TryPop(out DocxMathNode? node))
        {
            token.ThrowIfCancellationRequested();
            if (node is DocxMathRow sequence && (sequence.HasSupportedProperties ||
                ReferenceEquals(sequence, row) && allowArgumentSize && sequence.HasKnownArgumentProperties))
            {
                for (int i = sequence.Children.Count - 1; i >= 0; i--) { pending.Push(sequence.Children[i]); }
            }
            else if (node is DocxMathRun run && HasSupportedStyle(run))
            {
                foreach (Rune rune in run.Text.Text.EnumerateRunes())
                {
                    token.ThrowIfCancellationRequested();
                    int rendered = run.NormalText ? rune.Value : RenderCodePoint(rune.Value, run.MathStyleValue);
                    if (rendered < 0) { fallback = "UnsupportedMathAlphabet"; return false; }
                    OoxConversionBudget.Current?.ChargeSceneNodes(1);
                    atoms.Add(new(run, rune, new Rune(rendered), run.NormalText ? AtomKind.Ordinary : Kind(rune.Value)));
                }
            }
            else { fallback = node is DocxMathUnsupported unsupported ? "UnsupportedMathNode:" + unsupported.ElementName : "UnsupportedMathStructureOrStyle"; return false; }
        }
        return true;
    }

    private static bool HasSupportedStyle(DocxMathRun run) =>
        !run.HasUnsupportedProperties && run.MathStyleValue is null or "i" or "p" or "b" or "bi" && run.ScriptValue is null or "roman" &&
        HasSupportedTextFormatting(run.Text, run.NormalText);

    internal static bool HasSupportedTextFormatting(DocxTextRun text, bool allowBold) =>
        (!text.Bold || allowBold) && !text.Underline && !text.Strike && !text.DoubleStrike && !text.Hidden &&
        !text.AllCaps && !text.SmallCaps && text.VerticalAlignmentValue is null && text.CharacterSpacingPoints == 0d &&
        text.HighlightValue is null && text.ShadingFillHex is null;

    private static int RenderCodePoint(int codePoint, string? style)
    {
        // Word treats a mathematical hyphen as minus, including upright math.
        // Normal-text runs bypass this mapper and retain their punctuation.
        if (codePoint == '-') { return 0x2212; }
        if (style == "p") { return codePoint; }
        if (style is null or "i") { return ItalicCodePoint(codePoint); }
        bool italic = style == "bi";
        if (codePoint is >= 'A' and <= 'Z') { return (italic ? 0x1D468 : 0x1D400) + codePoint - 'A'; }
        if (codePoint is >= 'a' and <= 'z') { return (italic ? 0x1D482 : 0x1D41A) + codePoint - 'a'; }
        if (codePoint is >= 0x0391 and <= 0x03A9 && codePoint != 0x03A2)
            { return (italic ? 0x1D71C : 0x1D6A8) + codePoint - 0x0391; }
        if (codePoint is >= 0x03B1 and <= 0x03C9)
            { return (italic ? 0x1D736 : 0x1D6C2) + codePoint - 0x03B1; }
        if (codePoint is >= '0' and <= '9') { return 0x1D7CE + codePoint - '0'; }
        if (codePoint == 0x2207) { return italic ? 0x1D735 : 0x1D6C1; }
        if (TryGetVariantSymbolOffset(codePoint, out int variant)) { return (italic ? 0x1D74F : 0x1D6DB) + variant; }
        // Operators retain their mathematical symbol glyph. Other alphabets
        // keep fallback until their explicit styled mapping is qualified.
        if (Rune.IsLetter(new Rune(codePoint)) || Rune.IsDigit(new Rune(codePoint))) { return -1; }
        return codePoint;
    }

    private static int ItalicCodePoint(int codePoint)
    {
        if (codePoint is >= 'A' and <= 'Z') { return 0x1D434 + codePoint - 'A'; }
        if (codePoint is >= 'a' and <= 'z') { return codePoint == 'h' ? 0x210E : 0x1D44E + codePoint - 'a'; }
        if (codePoint is >= 0x0391 and <= 0x03A9 && codePoint != 0x03A2) { return 0x1D6E2 + codePoint - 0x0391; }
        if (codePoint is >= 0x03B1 and <= 0x03C9) { return 0x1D6FC + codePoint - 0x03B1; }
        if (codePoint == 0x2207) { return 0x1D6FB; }
        if (TryGetVariantSymbolOffset(codePoint, out int variant)) { return 0x1D715 + variant; }
        if (Rune.IsLetter(new Rune(codePoint)) && codePoint is not (>= 0x1D400 and <= 0x1D7FF) && codePoint != 0x210E) { return -1; }
        return codePoint;
    }

    private static bool TryGetVariantSymbolOffset(int codePoint, out int offset)
    {
        offset = codePoint switch
        {
            0x2202 => 0, // partial differential
            0x03F5 => 1, // epsilon symbol
            0x03D1 => 2, // theta symbol
            0x03F0 => 3, // kappa symbol
            0x03D5 => 4, // phi symbol
            0x03F1 => 5, // rho symbol
            0x03D6 => 6, // pi symbol
            _ => -1
        };
        return offset >= 0;
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
    internal static int Glue(AtomKind left, AtomKind right) =>
        left == AtomKind.Open || right == AtomKind.Close ? 0 :
        left == AtomKind.LargeOperator && right == AtomKind.Punctuation ? 0 :
        left == AtomKind.Relation || right == AtomKind.Relation ? 5 :
        left == AtomKind.Binary || right == AtomKind.Binary ? 4 :
        left == AtomKind.LargeOperator || right == AtomKind.LargeOperator ? 3 :
        left == AtomKind.Inner || right == AtomKind.Inner ? 3 :
        left == AtomKind.Punctuation ? 3 : 0;
}
