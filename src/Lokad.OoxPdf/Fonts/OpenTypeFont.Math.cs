namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    public bool TryGetMathScriptGlyph(ushort glyph, int level, CancellationToken token, out ushort result)
    {
        result = glyph;
        return glyph != 0 && glyph < GlyphCount && mathScriptStyles.TryGetGlyph(glyph, level, token, out result) &&
            result != 0 && result < GlyphCount;
    }

    public bool TryGetMathConstants(out OpenTypeMathConstants constants)
    {
        constants = default;
        return tables.TryGetValue("MATH", out TableRecord table) &&
            OpenTypeMathConstants.TryRead(bytes.AsMemory(table.Offset, table.Length), out constants);
    }
}
