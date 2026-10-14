using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

internal enum MathKernCorner { TopRight, TopLeft, BottomRight, BottomLeft }

internal sealed partial class OpenTypeFont
{
    public bool TryGetMathKern(ushort glyph, MathKernCorner corner, double height, out short correction)
    {
        correction = 0;
        if (glyph >= GlyphCount || (uint)corner > 3u || !double.IsFinite(height)) { return false; }
        if (!tables.TryGetValue("MATH", out TableRecord math)) { return true; }
        ReadOnlySpan<byte> table = bytes.AsSpan(math.Offset, math.Length);
        if (table.Length < 10 || U16(table, 0) != 1 || U16(table, 2) != 0) { return false; }
        int infoOffset = U16(table, 6);
        if (infoOffset == 0) { return true; }
        if (infoOffset < 10 || table.Length - infoOffset < 8) { return false; }
        ReadOnlySpan<byte> info = table.Slice(infoOffset);
        int kernInfoOffset = U16(info, 6);
        if (kernInfoOffset == 0) { return true; }
        if (kernInfoOffset < 8 || info.Length - kernInfoOffset < 4) { return false; }
        ReadOnlySpan<byte> records = info.Slice(kernInfoOffset);
        int coverageOffset = U16(records, 0), count = U16(records, 2), recordsEnd = 4 + count * 8;
        if (records.Length < recordsEnd || coverageOffset < recordsEnd || records.Length - coverageOffset < 4 ||
            !TryGetCoverageIndex(records.Slice(coverageOffset), glyph, out int index)) { return false; }
        if (index < 0) { return true; }
        if (index >= count) { return false; }
        int kernOffset = U16(records, 4 + index * 8 + (int)corner * 2);
        if (kernOffset == 0) { return true; }
        if (kernOffset < recordsEnd || records.Length - kernOffset < 2) { return false; }
        ReadOnlySpan<byte> kern = records.Slice(kernOffset);
        int heights = U16(kern, 0);
        if (kern.Length < 6 + heights * 8) { return false; }
        // Upper bound: equality belongs to the interval above its threshold.
        int low = 0, high = heights;
        while (low < high)
        {
            int middle = low + (high - low) / 2;
            short threshold = BinaryPrimitives.ReadInt16BigEndian(kern.Slice(2 + middle * 4, 2));
            if (height >= threshold) { low = middle + 1; }
            else { high = middle; }
        }
        correction = BinaryPrimitives.ReadInt16BigEndian(kern.Slice(2 + heights * 4 + low * 4, 2));
        return true;
    }
}
