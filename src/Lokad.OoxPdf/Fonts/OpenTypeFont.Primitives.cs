using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private static TableRecord Required(Dictionary<string, TableRecord> tables, string tag)
    {
        return tables.TryGetValue(tag, out TableRecord table)
            ? table
            : throw new InvalidDataException($"Font is missing required '{tag}' table.");
    }

    internal static ushort U16(ReadOnlySpan<byte> bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt16BigEndian(bytes.Slice(offset, 2));
    }

    internal static uint U32(ReadOnlySpan<byte> bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(offset, 4));
    }

    internal static short I16(ReadOnlySpan<byte> bytes, int offset)
    {
        return BinaryPrimitives.ReadInt16BigEndian(bytes.Slice(offset, 2));
    }

    private static int I32(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
    }

    internal static void W16(Span<byte> bytes, int offset, ushort value)
    {
        BinaryPrimitives.WriteUInt16BigEndian(bytes.Slice(offset, 2), value);
    }

    internal static void W32(Span<byte> bytes, int offset, uint value)
    {
        BinaryPrimitives.WriteUInt32BigEndian(bytes.Slice(offset, 4), value);
    }

    private static double F2Dot14(byte[] bytes, int offset)
    {
        return I16(bytes, offset) / 16384d;
    }

    internal static bool TryGetGlyphData(ReadOnlySpan<byte> source, ushort glyph, ushort glyphCount, short format,
        TableRecord glyf, TableRecord loca, out int offset, out int length)
    {
        offset = 0;
        length = 0;
        if (glyph >= glyphCount || format is not 0 and not 1) return false;

        int entrySize = format == 0 ? 2 : 4;
        int relativeEntry = glyph * entrySize;
        if (relativeEntry > loca.Length - 2 * entrySize) return false;

        int entry = loca.Offset + relativeEntry;
        uint start = format == 0 ? 2u * U16(source, entry) : U32(source, entry);
        uint end = format == 0 ? 2u * U16(source, entry + entrySize) : U32(source, entry + entrySize);
        // Validate the offsets against the table's length before forming absolute positions.
        // Empty glyphs are valid storage. Outline readers still require a glyph header.
        if (start > end || end > (uint)glyf.Length) return false;

        offset = glyf.Offset + (int)start;
        length = (int)(end - start);
        return true;
    }

    internal static int Align4(int value)
    {
        return (value + 3) & ~3;
    }
}
