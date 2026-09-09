using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private static TableRecord Required(Dictionary<string, TableRecord> tables, string tag)
    {
        return tables.TryGetValue(tag, out TableRecord table)
            ? table
            : throw new InvalidDataException($"Font is missing required '{tag}' table.");
    }

    internal static ushort U16(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));
    }

    internal static uint U32(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4));
    }

    internal static short I16(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadInt16BigEndian(bytes.AsSpan(offset, 2));
    }

    private static int I32(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
    }

    internal static void W16(byte[] bytes, int offset, ushort value)
    {
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset, 2), value);
    }

    internal static void W32(byte[] bytes, int offset, uint value)
    {
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(offset, 4), value);
    }

    private static double F2Dot14(byte[] bytes, int offset)
    {
        return I16(bytes, offset) / 16384d;
    }

    internal static int Align4(int value)
    {
        return (value + 3) & ~3;
    }
}
