using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    public bool TryGetGlyphBounds(ushort glyph, out FontBounds bounds)
    {
        bounds = default;
        if (glyph >= GlyphCount || !tables.TryGetValue("head", out TableRecord head) ||
            !tables.TryGetValue("loca", out TableRecord loca) || !tables.TryGetValue("glyf", out TableRecord glyf) ||
            !TryGetGlyphTableRange(glyph, I16(bytes, head.Offset + 50), loca, glyf, out int offset, out int length) ||
            length < 10 || offset < 0 || offset > bytes.Length - length)
        { return false; }
        // The glyph header contains its ink bounds. Do not expand outline points
        // or compound descendants just to measure a mathematical box.
        bounds = new FontBounds(I16(bytes, offset + 2), I16(bytes, offset + 4), I16(bytes, offset + 6), I16(bytes, offset + 8));
        return bounds.XMin <= bounds.XMax && bounds.YMin <= bounds.YMax;
    }

    public bool TryGetMathItalicCorrection(ushort glyph, out short correction)
    {
        correction = 0;
        if (!tables.TryGetValue("MATH", out TableRecord math)) { return true; }
        ReadOnlySpan<byte> table = bytes.AsSpan(math.Offset, math.Length);
        if (table.Length < 10 || U16(table, 0) != 1 || U16(table, 2) != 0) { return false; }
        int infoOffset = U16(table, 6);
        if (infoOffset == 0) { return true; }
        if (infoOffset < 10 || table.Length - infoOffset < 8) { return false; }
        ReadOnlySpan<byte> info = table.Slice(infoOffset);
        int italicOffset = U16(info, 0);
        if (italicOffset == 0) { return true; }
        if (italicOffset < 8 || info.Length - italicOffset < 4) { return false; }
        ReadOnlySpan<byte> italic = info.Slice(italicOffset);
        int coverageOffset = U16(italic, 0), count = U16(italic, 2);
        int recordsEnd = 4 + count * 4;
        if (italic.Length < recordsEnd || coverageOffset < recordsEnd || italic.Length - coverageOffset < 4) { return false; }
        if (!TryGetMathCoverageIndex(italic.Slice(coverageOffset), glyph, out int index)) { return false; }
        if (index < 0) { return true; }
        if (index >= count) { return false; }
        correction = BinaryPrimitives.ReadInt16BigEndian(italic.Slice(4 + index * 4, 2));
        return true;
    }

    private static ushort U16(ReadOnlySpan<byte> table, int offset) => BinaryPrimitives.ReadUInt16BigEndian(table.Slice(offset, 2));

    private static bool TryGetMathCoverageIndex(ReadOnlySpan<byte> coverage, ushort glyph, out int index)
    {
        index = -1;
        int format = U16(coverage, 0), count = U16(coverage, 2);
        if (format is not (1 or 2)) { return false; }
        int stride = format == 1 ? 2 : 6;
        if (coverage.Length - 4 < count * stride) { return false; }
        int low = 0, high = count - 1;
        while (low <= high)
        {
            int middle = low + (high - low) / 2, offset = 4 + middle * stride;
            int start = U16(coverage, offset), end = format == 1 ? start : U16(coverage, offset + 2);
            if (end < start) { return false; }
            if (glyph < start) { high = middle - 1; }
            else if (glyph > end) { low = middle + 1; }
            else { index = format == 1 ? middle : U16(coverage, offset + 4) + glyph - start; return true; }
        }
        return true;
    }
}
