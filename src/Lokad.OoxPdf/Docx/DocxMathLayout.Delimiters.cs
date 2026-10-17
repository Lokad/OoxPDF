using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathDelimiter? GetDelimiter(DocxEquation equation) =>
        GetSupportedDelimiter(GetSingleSupportedChild(equation));

    private static DocxMathDelimiter? GetSupportedDelimiter(DocxMathNode? node) =>
        node is DocxMathDelimiter { HasUnsupportedProperties: false, Arguments.Count: > 0,
            ShapeValue: null or "centered" } delimiter &&
        (delimiter.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(delimiter.ControlText, allowBold: false)) &&
        IsCharacter(delimiter.BeginCharacter) && IsCharacter(delimiter.EndCharacter) &&
        (delimiter.Arguments.Count == 1 || IsCharacter(delimiter.SeparatorCharacter)) &&
        delimiter.BeginCharacter.Length + delimiter.EndCharacter.Length > 0 ? delimiter : null;

    private static bool IsCharacter(string value) => value.Length == 0 ||
        Rune.TryGetRuneAt(value, 0, out Rune rune) && rune.Utf16SequenceLength == value.Length;

    private static bool TryGetDelimiterRuns(DocxEquation equation, DocxMathDelimiter delimiter, CancellationToken token,
        out IReadOnlyList<DocxMathRunSource> runs)
    {
        var result = new List<DocxMathRunSource>();
        foreach (var argument in delimiter.Arguments)
        {
            token.ThrowIfCancellationRequested(); var child = equation with { Content = argument };
            // Admission shares the measured literal/structured argument roles.
            // A delimiter or unknown deeper owner cannot
            // recursively rediscover itself through the general composer.
            if (!HasStructuredBoxArgument(child, token) &&
                !DocxMathRowLayout.TryGetRenderCodePoints(argument, token, out _, out _) ||
                !TryGetRuns(child, token, out var sources))
            { runs = []; return false; }
            result.AddRange(sources);
        }
        if (result.Count != 0) result[0] = result[0] with { ControlOwners = [delimiter, .. result[0].ControlOwners] };
        runs = result; return true;
    }

    public static IReadOnlyList<int> GetControlCodePoints(IDocxMathControl control) => control switch
    {
        DocxMathRadical => [0x221A],
        DocxMathNary { CharacterValue: { Length: > 0 } character } => [Rune.GetRuneAt(character, 0).Value],
        DocxMathDelimiter delimiter => (delimiter.BeginCharacter + delimiter.EndCharacter +
            (delimiter.Arguments.Count > 1 ? delimiter.SeparatorCharacter : string.Empty)).EnumerateRunes()
            .Select(rune => rune.Value).Distinct().ToArray(),
        _ => []
    };

    private static bool TryMeasureDelimiter(DocxEquation equation, DocxMathDelimiter delimiter,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources))
        { fallback = "UnsupportedMathDelimiterSource"; return false; }
        PdfFontResource? resource = ResolveControlResource(delimiter, sources, resolveFont, resolveControlFont);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathDelimiterConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = ResolveControlSize(delimiter, equation, sources);
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathDelimiterSize"; return false; }
        var children = new List<DocxMathBox>();
        double above = 0d, below = 0d;
        foreach (var argument in delimiter.Arguments)
        {
            token.ThrowIfCancellationRequested();
            if (!TryMeasureBoxArgument(equation with { Content = argument }, resolveFont, font, size, token,
                out var child, out fallback, resolveControlFont: resolveControlFont, resolveControlMetrics: resolveControlMetrics) || child is null)
            { fallback ??= "UnsupportedMathDelimiterArgument"; return false; }
            children.Add(child);above = Math.Max(above, child.AboveBaseline);below = Math.Max(below, child.BelowBaseline);
        }
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), scale = glyphSize / font.UnitsPerEm;
        double axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * size / font.UnitsPerEm;
        double height = children.Max(child => child.EnclosureHeight ??
            2d * Math.Max(child.AboveBaseline - axis, child.BelowBaseline + axis));
        if (!double.IsFinite(height) || height <= 0d)
        { fallback = "InvalidMathDelimiterExtent"; return false; }
        RgbColor color = children.SelectMany(child => child.Glyphs).FirstOrDefault()?.Color ?? default;
        if (delimiter.ControlText is { } control) { color = default; RgbColor.TryParse(control.ColorHex, out color); }
        var glyphs = new List<DocxMathGlyphPlacement>();
        var rules = new List<DocxMathRulePlacement>();
        double advance = 0d, width = 0d, delimiterHeight = 0d;
        if (!AppendDelimiter(delimiter.BeginCharacter)) { fallback = "UnsupportedMathDelimiterVariant"; return false; }
        for (int i = 0; i < children.Count; i++)
        {
            if (i > 0 && !AppendDelimiter(delimiter.SeparatorCharacter)) { fallback = "UnsupportedMathDelimiterSeparator"; return false; }
            var child = children[i];double childX = advance;
            foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + childX }); }
            foreach (var rule in child.Rules) { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + childX }); }
            width = Math.Max(width, childX + child.Width); advance += child.AdvanceWidth;
        }
        if (!AppendDelimiter(delimiter.EndCharacter)) { fallback = "UnsupportedMathDelimiterVariant"; return false; }
        width = Math.Max(width, advance);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) || width <= 0d)
        { fallback = "InvalidMathDelimiterExtent"; return false; }
        double subformulaMinimum = constants.DelimitedSubFormulaMinHeight * scale;
        // Classification follows the painted enclosing mark, so a small fixed
        // delimiter stays an opening/closing atom even around taller content.
        bool subformula = subformulaMinimum > 0d && delimiterHeight >= subformulaMinimum;
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = new(font, size),
            Edges = new(delimiter.BeginCharacter.Length > 0 ? (subformula ? DocxMathRowLayout.AtomKind.Inner : DocxMathRowLayout.AtomKind.Open) : children[0].Edges?.First ?? DocxMathRowLayout.AtomKind.Ordinary,
                delimiter.EndCharacter.Length > 0 ? (subformula ? DocxMathRowLayout.AtomKind.Inner : DocxMathRowLayout.AtomKind.Close) : children[^1].Edges?.Last ?? DocxMathRowLayout.AtomKind.Ordinary,
                delimiter.BeginCharacter.Length == 0 && children[0].Edges?.FirstNormalText == true) };
        fallback = null;
        return true;

        bool AppendDelimiter(string character)
        {
            if (character.Length == 0) { return true; }
            token.ThrowIfCancellationRequested();
            int point = Rune.GetRuneAt(character, 0).Value;
            ushort glyph = font.MapCodePoint(point);
            if (delimiter.GrowValue != false)
            {
                if (!font.TryGetNearestMathVerticalVariant(glyph, height / scale, token, out var variant)) { return false; }
                glyph = variant.Glyph;
            }
            if (glyph == 0 || !font.TryGetGlyphBounds(glyph, out var bounds) || !resource.Font.TryGetEncodedCid(glyph, out ushort cid))
            { return false; }
            double baseline = axis - (bounds.YMax + bounds.YMin) / 2d * scale;
            double glyphAdvance = font.GetAdvanceWidth(glyph) * scale;
            delimiterHeight = Math.Max(delimiterHeight, (bounds.YMax - bounds.YMin) * scale);
            OoxConversionBudget.Current?.ChargeSceneNodes(1);
            glyphs.Add(new(point, point, resource, cid.ToString("X4", CultureInfo.InvariantCulture), glyphSize,
                advance, baseline, glyphAdvance, color) { OriginalGlyph = glyph, ControlOwner = delimiter });
            width = Math.Max(width, advance + bounds.XMax * scale);
            above = Math.Max(above, baseline + bounds.YMax * scale);
            below = Math.Max(below, -baseline - bounds.YMin * scale);
            advance += glyphAdvance;
            return true;
        }
    }
}
