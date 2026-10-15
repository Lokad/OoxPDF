using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathDelimiter? GetDelimiter(DocxEquation equation) =>
        GetSingleSupportedChild(equation) is DocxMathDelimiter { HasUnsupportedProperties: false, Arguments.Count: 1,
            ShapeValue: null or "centered" } delimiter &&
        (delimiter.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(delimiter.ControlText, allowBold: false)) &&
        IsCharacter(delimiter.BeginCharacter) && IsCharacter(delimiter.EndCharacter) &&
        delimiter.BeginCharacter.Length + delimiter.EndCharacter.Length > 0 ? delimiter : null;

    private static bool IsCharacter(string value) => value.Length == 0 ||
        Rune.TryGetRuneAt(value, 0, out Rune rune) && rune.Utf16SequenceLength == value.Length;

    public static IReadOnlyList<int> GetControlCodePoints(IDocxMathControl control) => control switch
    {
        DocxMathRadical => [0x221A],
        DocxMathDelimiter delimiter => (delimiter.BeginCharacter + delimiter.EndCharacter).EnumerateRunes()
            .Select(rune => rune.Value).Distinct().ToArray(),
        _ => []
    };

    private static bool TryMeasureDelimiter(DocxEquation equation, DocxMathDelimiter delimiter,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathDelimiterSource"; return false; }
        PdfFontResource? resource = resolveControlFont is null ? resolveFont(sources[0].Run) : resolveControlFont(delimiter);
        if (resource is null || !resource.Font.Font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathDelimiterConstants"; return false; }
        OpenTypeFont font = resource.Font.Font;
        double size = delimiter.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathDelimiterSize"; return false; }
        if (!DocxMathRowLayout.TryMeasure(delimiter.Arguments[0], resolveFont, token, out var child, out fallback,
            exportFontSize: true, operatorSpacingFontSize: size) || child is null || child.Glyphs.Count == 0)
        { fallback ??= "UnsupportedMathDelimiterArgument"; return false; }
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size), scale = glyphSize / font.UnitsPerEm;
        double axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * size / font.UnitsPerEm;
        double height = 2d * Math.Max(child.AboveBaseline - axis, child.BelowBaseline + axis);
        if (!double.IsFinite(height) || height <= 0d)
        { fallback = "InvalidMathDelimiterExtent"; return false; }
        RgbColor color = child.Glyphs[0].Color;
        if (delimiter.ControlText is { } control) { color = default; RgbColor.TryParse(control.ColorHex, out color); }
        var glyphs = new List<DocxMathGlyphPlacement>(child.Glyphs.Count + 2);
        double advance = 0d, width = 0d, above = child.AboveBaseline, below = child.BelowBaseline;
        if (!AppendDelimiter(delimiter.BeginCharacter)) { fallback = "UnsupportedMathDelimiterVariant"; return false; }
        double childX = advance;
        foreach (var glyph in child.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + childX }); }
        width = Math.Max(width, childX + child.Width); advance += child.AdvanceWidth;
        if (!AppendDelimiter(delimiter.EndCharacter)) { fallback = "UnsupportedMathDelimiterVariant"; return false; }
        width = Math.Max(width, advance);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) || width <= 0d)
        { fallback = "InvalidMathDelimiterExtent"; return false; }
        var rules = new List<DocxMathRulePlacement>(child.Rules.Count);
        foreach (var rule in child.Rules) { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + childX }); }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = new(font, size) };
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
            OoxConversionBudget.Current?.ChargeSceneNodes(1);
            glyphs.Add(new(point, point, resource, cid.ToString("X4", CultureInfo.InvariantCulture), glyphSize,
                advance, baseline, glyphAdvance, color) { OriginalGlyph = glyph });
            width = Math.Max(width, advance + bounds.XMax * scale);
            above = Math.Max(above, baseline + bounds.YMax * scale);
            below = Math.Max(below, -baseline - bounds.YMin * scale);
            advance += glyphAdvance;
            return true;
        }
    }
}
