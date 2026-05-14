using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed class OpenTypeFont
{
    private readonly byte[] bytes;
    private readonly Dictionary<string, TableRecord> tables;
    private readonly CmapFormat? cmap;
    private readonly ushort[] advances;

    private OpenTypeFont(
        byte[] bytes,
        Dictionary<string, TableRecord> tables,
        string familyName,
        ushort unitsPerEm,
        FontBounds bounds,
        ushort glyphCount,
        Os2Metrics os2,
        PostMetrics post,
        CmapFormat? cmap,
        ushort[] advances)
    {
        this.bytes = bytes;
        this.tables = tables;
        FamilyName = familyName;
        UnitsPerEm = unitsPerEm;
        Bounds = bounds;
        GlyphCount = glyphCount;
        Os2 = os2;
        Post = post;
        this.cmap = cmap;
        this.advances = advances;
    }

    public string FamilyName { get; }

    public ushort UnitsPerEm { get; }

    public FontBounds Bounds { get; }

    public ReadOnlyMemory<byte> Bytes => bytes;

    public ushort GlyphCount { get; }

    public Os2Metrics Os2 { get; }

    public PostMetrics Post { get; }

    public IReadOnlyCollection<string> TableTags => tables.Keys;

    public static OpenTypeFont Load(string path)
    {
        return Load(File.ReadAllBytes(path));
    }

    public static OpenTypeFont Load(byte[] bytes)
    {
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        ushort numTables = U16(bytes, 4);
        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        int offset = 12;
        for (int i = 0; i < numTables; i++)
        {
            string tag = Encoding.ASCII.GetString(bytes, offset, 4);
            uint tableOffset = U32(bytes, offset + 8);
            uint length = U32(bytes, offset + 12);
            if (tableOffset + length > bytes.Length)
            {
                throw new InvalidDataException($"Font table '{tag}' exceeds file length.");
            }

            tables[tag] = new TableRecord((int)tableOffset, (int)length);
            offset += 16;
        }

        ushort unitsPerEm = ReadUnitsPerEm(bytes, tables);
        FontBounds bounds = ReadBounds(bytes, tables);
        ushort glyphCount = ReadGlyphCount(bytes, tables);
        string familyName = ReadFamilyName(bytes, tables);
        Os2Metrics os2 = ReadOs2(bytes, tables);
        PostMetrics post = ReadPost(bytes, tables);
        CmapFormat? cmap = ReadCmap(bytes, tables);
        ushort[] advances = ReadAdvances(bytes, tables);
        return new OpenTypeFont(bytes, tables, familyName, unitsPerEm, bounds, glyphCount, os2, post, cmap, advances);
    }

    public ushort MapCodePoint(int codePoint)
    {
        return cmap?.Map(codePoint, bytes) ?? 0;
    }

    public ushort GetAdvanceWidth(ushort glyphId)
    {
        if (advances.Length == 0)
        {
            return 0;
        }

        return glyphId < advances.Length ? advances[glyphId] : advances[^1];
    }

    private static ushort ReadUnitsPerEm(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord head = Required(tables, "head");
        return U16(bytes, head.Offset + 18);
    }

    private static FontBounds ReadBounds(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord head = Required(tables, "head");
        return new FontBounds(
            I16(bytes, head.Offset + 36),
            I16(bytes, head.Offset + 38),
            I16(bytes, head.Offset + 40),
            I16(bytes, head.Offset + 42));
    }

    private static ushort ReadGlyphCount(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord maxp = Required(tables, "maxp");
        return U16(bytes, maxp.Offset + 4);
    }

    private static Os2Metrics ReadOs2(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord os2 = Required(tables, "OS/2");
        return new Os2Metrics(
            Version: U16(bytes, os2.Offset),
            WeightClass: U16(bytes, os2.Offset + 4),
            WidthClass: U16(bytes, os2.Offset + 6),
            TypographicAscender: I16(bytes, os2.Offset + 68),
            TypographicDescender: I16(bytes, os2.Offset + 70),
            TypographicLineGap: I16(bytes, os2.Offset + 72),
            WindowsAscender: U16(bytes, os2.Offset + 74),
            WindowsDescender: U16(bytes, os2.Offset + 76));
    }

    private static PostMetrics ReadPost(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord post = Required(tables, "post");
        return new PostMetrics(
            ItalicAngle: FixedToDouble(I32(bytes, post.Offset + 4)),
            UnderlinePosition: I16(bytes, post.Offset + 8),
            UnderlineThickness: I16(bytes, post.Offset + 10),
            IsFixedPitch: U32(bytes, post.Offset + 12) != 0);
    }

    private static string ReadFamilyName(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord name = Required(tables, "name");
        ushort count = U16(bytes, name.Offset + 2);
        ushort stringOffset = U16(bytes, name.Offset + 4);
        string? fallback = null;

        for (int i = 0; i < count; i++)
        {
            int record = name.Offset + 6 + i * 12;
            ushort platform = U16(bytes, record);
            ushort nameId = U16(bytes, record + 6);
            ushort length = U16(bytes, record + 8);
            ushort offset = U16(bytes, record + 10);
            if (nameId != 1)
            {
                continue;
            }

            int valueOffset = name.Offset + stringOffset + offset;
            string value = platform == 3
                ? Encoding.BigEndianUnicode.GetString(bytes, valueOffset, length)
                : Encoding.ASCII.GetString(bytes, valueOffset, length);
            if (platform == 3)
            {
                return value.TrimEnd('\0');
            }

            fallback ??= value.TrimEnd('\0');
        }

        return fallback ?? "Unknown";
    }

    private static CmapFormat? ReadCmap(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        if (!tables.TryGetValue("cmap", out TableRecord cmapTable))
        {
            return null;
        }

        ushort count = U16(bytes, cmapTable.Offset + 2);
        int? bestOffset = null;
        for (int i = 0; i < count; i++)
        {
            int record = cmapTable.Offset + 4 + i * 8;
            ushort platform = U16(bytes, record);
            ushort encoding = U16(bytes, record + 2);
            uint subtableOffset = U32(bytes, record + 4);
            int absolute = cmapTable.Offset + (int)subtableOffset;
            ushort format = U16(bytes, absolute);
            if (platform == 3 && encoding is 1 or 10 && format is 4 or 12)
            {
                bestOffset = absolute;
                if (format == 12)
                {
                    break;
                }
            }
        }

        if (bestOffset is null)
        {
            return null;
        }

        return U16(bytes, bestOffset.Value) switch
        {
            4 => CmapFormat4.Read(bytes, bestOffset.Value),
            12 => CmapFormat12.Read(bytes, bestOffset.Value),
            _ => null
        };
    }

    private static ushort[] ReadAdvances(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        if (!tables.TryGetValue("hhea", out TableRecord hhea) || !tables.TryGetValue("hmtx", out TableRecord hmtx))
        {
            return [];
        }

        ushort numberOfHMetrics = U16(bytes, hhea.Offset + 34);
        var advances = new ushort[numberOfHMetrics];
        for (int i = 0; i < advances.Length; i++)
        {
            advances[i] = U16(bytes, hmtx.Offset + i * 4);
        }

        return advances;
    }

    private static TableRecord Required(Dictionary<string, TableRecord> tables, string tag)
    {
        return tables.TryGetValue(tag, out TableRecord table)
            ? table
            : throw new InvalidDataException($"Font is missing required '{tag}' table.");
    }

    private static ushort U16(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));
    }

    private static uint U32(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4));
    }

    private static short I16(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadInt16BigEndian(bytes.AsSpan(offset, 2));
    }

    private static int I32(byte[] bytes, int offset)
    {
        return BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
    }

    private static double FixedToDouble(int value)
    {
        return value / 65536d;
    }

    private readonly record struct TableRecord(int Offset, int Length);

    internal readonly record struct Os2Metrics(
        ushort Version,
        ushort WeightClass,
        ushort WidthClass,
        short TypographicAscender,
        short TypographicDescender,
        short TypographicLineGap,
        ushort WindowsAscender,
        ushort WindowsDescender);

    internal readonly record struct FontBounds(short XMin, short YMin, short XMax, short YMax);

    internal readonly record struct PostMetrics(
        double ItalicAngle,
        short UnderlinePosition,
        short UnderlineThickness,
        bool IsFixedPitch);

    private abstract record CmapFormat
    {
        public abstract ushort Map(int codePoint, byte[] bytes);
    }

    private sealed record CmapFormat4(ushort[] EndCodes, ushort[] StartCodes, short[] IdDeltas, ushort[] IdRangeOffsets, int GlyphArrayOffset) : CmapFormat
    {
        public static CmapFormat4 Read(byte[] bytes, int offset)
        {
            ushort segCount = (ushort)(U16(bytes, offset + 6) / 2);
            int endCodesOffset = offset + 14;
            int startCodesOffset = endCodesOffset + segCount * 2 + 2;
            int idDeltasOffset = startCodesOffset + segCount * 2;
            int idRangeOffsetsOffset = idDeltasOffset + segCount * 2;
            var endCodes = new ushort[segCount];
            var startCodes = new ushort[segCount];
            var idDeltas = new short[segCount];
            var idRangeOffsets = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
            {
                endCodes[i] = U16(bytes, endCodesOffset + i * 2);
                startCodes[i] = U16(bytes, startCodesOffset + i * 2);
                idDeltas[i] = unchecked((short)U16(bytes, idDeltasOffset + i * 2));
                idRangeOffsets[i] = U16(bytes, idRangeOffsetsOffset + i * 2);
            }

            return new CmapFormat4(endCodes, startCodes, idDeltas, idRangeOffsets, idRangeOffsetsOffset + segCount * 2);
        }

        public override ushort Map(int codePoint, byte[] bytes)
        {
            if (codePoint > ushort.MaxValue)
            {
                return 0;
            }

            ushort c = (ushort)codePoint;
            for (int i = 0; i < EndCodes.Length; i++)
            {
                if (c < StartCodes[i] || c > EndCodes[i])
                {
                    continue;
                }

                if (IdRangeOffsets[i] == 0)
                {
                    return (ushort)((c + IdDeltas[i]) & 0xFFFF);
                }

                int idRangeOffsetAddress = GlyphArrayOffset - IdRangeOffsets.Length * 2 + i * 2;
                int glyphOffset = idRangeOffsetAddress + IdRangeOffsets[i] + (c - StartCodes[i]) * 2;
                ushort glyph = U16(bytes, glyphOffset);
                return glyph == 0 ? (ushort)0 : (ushort)((glyph + IdDeltas[i]) & 0xFFFF);
            }

            return 0;
        }
    }

    private sealed record CmapFormat12(CmapGroup[] Groups) : CmapFormat
    {
        public static CmapFormat12 Read(byte[] bytes, int offset)
        {
            uint count = U32(bytes, offset + 12);
            var groups = new CmapGroup[count];
            int groupOffset = offset + 16;
            for (int i = 0; i < groups.Length; i++)
            {
                groups[i] = new CmapGroup(U32(bytes, groupOffset), U32(bytes, groupOffset + 4), U32(bytes, groupOffset + 8));
                groupOffset += 12;
            }

            return new CmapFormat12(groups);
        }

        public override ushort Map(int codePoint, byte[] bytes)
        {
            uint c = (uint)codePoint;
            foreach (CmapGroup group in Groups)
            {
                if (c >= group.StartCode && c <= group.EndCode)
                {
                    uint glyph = group.StartGlyph + c - group.StartCode;
                    return glyph <= ushort.MaxValue ? (ushort)glyph : (ushort)0;
                }
            }

            return 0;
        }
    }

    private readonly record struct CmapGroup(uint StartCode, uint EndCode, uint StartGlyph);
}
