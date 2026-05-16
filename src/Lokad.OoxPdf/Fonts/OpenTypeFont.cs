using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed class OpenTypeFont
{
    private readonly byte[] bytes;
    private readonly Dictionary<string, TableRecord> tables;
    private readonly CmapFormat? cmap;
    private readonly ushort[] advances;
    private readonly IReadOnlyDictionary<uint, short> kerningPairs;

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
        ushort[] advances,
        IReadOnlyDictionary<uint, short> kerningPairs)
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
        this.kerningPairs = kerningPairs;
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

        if (Encoding.ASCII.GetString(bytes, 0, 4) == "ttcf")
        {
            bytes = ExtractCollectionFont(bytes, 0);
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
        IReadOnlyDictionary<uint, short> kerningPairs = ReadKerningPairs(bytes, tables);
        return new OpenTypeFont(bytes, tables, familyName, unitsPerEm, bounds, glyphCount, os2, post, cmap, advances, kerningPairs);
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

    public short GetKerning(ushort leftGlyphId, ushort rightGlyphId)
    {
        uint key = ((uint)leftGlyphId << 16) | rightGlyphId;
        return kerningPairs.TryGetValue(key, out short value) ? value : (short)0;
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
        int? symbolOffset = null;
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
            else if (platform == 3 && encoding == 0 && format is 4 or 12)
            {
                symbolOffset ??= absolute;
            }
        }

        int? selectedOffset = bestOffset ?? symbolOffset;
        if (selectedOffset is null)
        {
            return null;
        }

        return U16(bytes, selectedOffset.Value) switch
        {
            4 => CmapFormat4.Read(bytes, selectedOffset.Value),
            12 => CmapFormat12.Read(bytes, selectedOffset.Value),
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

    private static IReadOnlyDictionary<uint, short> ReadKerningPairs(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        var pairs = new Dictionary<uint, short>();
        ReadLegacyKerningPairs(bytes, tables, pairs);
        ReadGposPairAdjustments(bytes, tables, pairs);
        return pairs;
    }

    private static void ReadLegacyKerningPairs(byte[] bytes, Dictionary<string, TableRecord> tables, Dictionary<uint, short> pairs)
    {
        if (!tables.TryGetValue("kern", out TableRecord kern) || kern.Length < 4)
        {
            return;
        }

        ushort tableCount = U16(bytes, kern.Offset + 2);
        int subtableOffset = kern.Offset + 4;
        for (int table = 0; table < tableCount && subtableOffset + 6 <= kern.Offset + kern.Length; table++)
        {
            ushort length = U16(bytes, subtableOffset + 2);
            ushort coverage = U16(bytes, subtableOffset + 4);
            int format = coverage >> 8;
            bool horizontal = (coverage & 0x0001) != 0;
            if (length >= 14 && format == 0 && horizontal)
            {
                ushort pairCount = U16(bytes, subtableOffset + 6);
                int pairOffset = subtableOffset + 14;
                int pairEnd = Math.Min(subtableOffset + length, kern.Offset + kern.Length);
                for (int i = 0; i < pairCount && pairOffset + 6 <= pairEnd; i++)
                {
                    ushort left = U16(bytes, pairOffset);
                    ushort right = U16(bytes, pairOffset + 2);
                    short value = I16(bytes, pairOffset + 4);
                    if (value != 0)
                    {
                        pairs[((uint)left << 16) | right] = value;
                    }

                    pairOffset += 6;
                }
            }

            if (length == 0)
            {
                break;
            }

            subtableOffset += length;
        }
    }

    private static void ReadGposPairAdjustments(byte[] bytes, Dictionary<string, TableRecord> tables, Dictionary<uint, short> pairs)
    {
        if (!tables.TryGetValue("GPOS", out TableRecord gpos) || gpos.Length < 10)
        {
            return;
        }

        int tableEnd = gpos.Offset + gpos.Length;
        int lookupList = gpos.Offset + U16(bytes, gpos.Offset + 8);
        if (lookupList + 2 > tableEnd)
        {
            return;
        }

        HashSet<ushort> kernLookupIndices = ReadGposKernLookupIndices(bytes, gpos.Offset, tableEnd);
        if (kernLookupIndices.Count == 0)
        {
            return;
        }

        ushort lookupCount = U16(bytes, lookupList);
        for (int i = 0; i < lookupCount && lookupList + 2 + i * 2 + 2 <= tableEnd; i++)
        {
            if (!kernLookupIndices.Contains((ushort)i))
            {
                continue;
            }

            int lookup = lookupList + U16(bytes, lookupList + 2 + i * 2);
            if (lookup + 6 > tableEnd)
            {
                continue;
            }

            ushort lookupType = U16(bytes, lookup);
            ushort subtableCount = U16(bytes, lookup + 4);
            for (int j = 0; j < subtableCount && lookup + 6 + j * 2 + 2 <= tableEnd; j++)
            {
                int subtable = lookup + U16(bytes, lookup + 6 + j * 2);
                ReadGposPairAdjustmentLookupSubtable(bytes, subtable, lookupType, tableEnd, pairs);
            }
        }
    }

    private static HashSet<ushort> ReadGposKernLookupIndices(byte[] bytes, int gposOffset, int tableEnd)
    {
        var lookupIndices = new HashSet<ushort>();
        int featureList = gposOffset + U16(bytes, gposOffset + 6);
        if (featureList + 2 > tableEnd)
        {
            return lookupIndices;
        }

        ushort featureCount = U16(bytes, featureList);
        for (int i = 0; i < featureCount && featureList + 2 + i * 6 + 6 <= tableEnd; i++)
        {
            int record = featureList + 2 + i * 6;
            string tag = Encoding.ASCII.GetString(bytes, record, 4);
            if (!tag.Equals("kern", StringComparison.Ordinal))
            {
                continue;
            }

            int feature = featureList + U16(bytes, record + 4);
            if (feature + 4 > tableEnd)
            {
                continue;
            }

            ushort lookupIndexCount = U16(bytes, feature + 2);
            for (int j = 0; j < lookupIndexCount && feature + 4 + j * 2 + 2 <= tableEnd; j++)
            {
                lookupIndices.Add(U16(bytes, feature + 4 + j * 2));
            }
        }

        return lookupIndices;
    }

    private static void ReadGposPairAdjustmentLookupSubtable(
        byte[] bytes,
        int subtable,
        ushort lookupType,
        int tableEnd,
        Dictionary<uint, short> pairs)
    {
        if (lookupType == 2)
        {
            ReadGposPairAdjustmentSubtable(bytes, subtable, tableEnd, pairs);
            return;
        }

        if (lookupType != 9 || subtable + 8 > tableEnd || U16(bytes, subtable) != 1)
        {
            return;
        }

        ushort extensionLookupType = U16(bytes, subtable + 2);
        uint extensionOffset = U32(bytes, subtable + 4);
        if (extensionLookupType != 2 || extensionOffset > int.MaxValue)
        {
            return;
        }

        int extensionSubtable = subtable + (int)extensionOffset;
        if (extensionSubtable < subtable || extensionSubtable >= tableEnd)
        {
            return;
        }

        ReadGposPairAdjustmentSubtable(bytes, extensionSubtable, tableEnd, pairs);
    }

    private static void ReadGposPairAdjustmentSubtable(byte[] bytes, int subtable, int tableEnd, Dictionary<uint, short> pairs)
    {
        if (subtable + 10 > tableEnd)
        {
            return;
        }

        ushort positionFormat = U16(bytes, subtable);
        int coverage = subtable + U16(bytes, subtable + 2);
        ushort valueFormat1 = U16(bytes, subtable + 4);
        ushort valueFormat2 = U16(bytes, subtable + 6);
        int valueRecordSize1 = ValueRecordSize(valueFormat1);
        int valueRecordSize2 = ValueRecordSize(valueFormat2);
        if (!TryReadCoverage(bytes, coverage, tableEnd, out ushort[] coverageGlyphs))
        {
            return;
        }

        if (positionFormat == 1)
        {
            ushort pairSetCount = U16(bytes, subtable + 8);
            for (int i = 0; i < pairSetCount && i < coverageGlyphs.Length && subtable + 10 + i * 2 + 2 <= tableEnd; i++)
            {
                int pairSet = subtable + U16(bytes, subtable + 10 + i * 2);
                if (pairSet + 2 > tableEnd)
                {
                    continue;
                }

                ushort pairValueCount = U16(bytes, pairSet);
                int pairValue = pairSet + 2;
                for (int j = 0; j < pairValueCount && pairValue + 2 + valueRecordSize1 + valueRecordSize2 <= tableEnd; j++)
                {
                    ushort rightGlyph = U16(bytes, pairValue);
                    short xAdvance = ReadXAdvance(bytes, pairValue + 2, valueFormat1);
                    if (xAdvance != 0)
                    {
                        pairs[((uint)coverageGlyphs[i] << 16) | rightGlyph] = xAdvance;
                    }

                    pairValue += 2 + valueRecordSize1 + valueRecordSize2;
                }
            }
        }
        else if (positionFormat == 2 && subtable + 16 <= tableEnd)
        {
            ushort class1Count = U16(bytes, subtable + 12);
            ushort class2Count = U16(bytes, subtable + 14);
            int classDef1 = subtable + U16(bytes, subtable + 8);
            int classDef2 = subtable + U16(bytes, subtable + 10);
            int classRecord = subtable + 16;
            int classRecordSize = valueRecordSize1 + valueRecordSize2;
            if (classRecordSize == 0)
            {
                return;
            }

            Dictionary<ushort, ushort[]> rightGlyphsByClass = ReadClassGlyphs(bytes, classDef2, tableEnd);
            foreach (ushort leftGlyph in coverageGlyphs)
            {
                ushort leftClass = ReadGlyphClass(bytes, classDef1, tableEnd, leftGlyph);
                if (leftClass >= class1Count)
                {
                    continue;
                }

                foreach (KeyValuePair<ushort, ushort[]> rightClassGlyphs in rightGlyphsByClass)
                {
                    if (rightClassGlyphs.Key == 0 || rightClassGlyphs.Key >= class2Count)
                    {
                        continue;
                    }

                    int record = classRecord + ((leftClass * class2Count + rightClassGlyphs.Key) * classRecordSize);
                    if (record + classRecordSize > tableEnd)
                    {
                        return;
                    }

                    short xAdvance = ReadXAdvance(bytes, record, valueFormat1);
                    if (xAdvance != 0)
                    {
                        foreach (ushort rightGlyph in rightClassGlyphs.Value)
                        {
                            pairs[((uint)leftGlyph << 16) | rightGlyph] = xAdvance;
                        }
                    }
                }
            }
        }
    }

    private static int ValueRecordSize(ushort valueFormat)
    {
        int size = 0;
        for (int bit = 0; bit < 8; bit++)
        {
            if ((valueFormat & (1 << bit)) != 0)
            {
                size += 2;
            }
        }

        return size;
    }

    private static short ReadXAdvance(byte[] bytes, int offset, ushort valueFormat)
    {
        for (int bit = 0; bit < 4; bit++)
        {
            if ((valueFormat & (1 << bit)) == 0)
            {
                continue;
            }

            if (bit == 2)
            {
                return I16(bytes, offset);
            }

            offset += 2;
        }

        return 0;
    }

    private static bool TryReadCoverage(byte[] bytes, int offset, int tableEnd, out ushort[] glyphs)
    {
        glyphs = [];
        if (offset + 4 > tableEnd)
        {
            return false;
        }

        ushort format = U16(bytes, offset);
        ushort count = U16(bytes, offset + 2);
        if (format == 1)
        {
            if (offset + 4 + count * 2 > tableEnd)
            {
                return false;
            }

            glyphs = new ushort[count];
            for (int i = 0; i < glyphs.Length; i++)
            {
                glyphs[i] = U16(bytes, offset + 4 + i * 2);
            }

            return true;
        }

        if (format != 2 || offset + 4 + count * 6 > tableEnd)
        {
            return false;
        }

        var list = new List<ushort>();
        for (int i = 0; i < count; i++)
        {
            int range = offset + 4 + i * 6;
            ushort start = U16(bytes, range);
            ushort end = U16(bytes, range + 2);
            for (int glyph = start; glyph <= end; glyph++)
            {
                list.Add((ushort)glyph);
            }
        }

        glyphs = list.ToArray();
        return true;
    }

    private static ushort ReadGlyphClass(byte[] bytes, int offset, int tableEnd, ushort glyphId)
    {
        if (offset + 4 > tableEnd)
        {
            return 0;
        }

        ushort format = U16(bytes, offset);
        if (format == 1)
        {
            ushort start = U16(bytes, offset + 2);
            ushort count = U16(bytes, offset + 4);
            int classOffset = offset + 6;
            if (glyphId < start || glyphId >= start + count || classOffset + count * 2 > tableEnd)
            {
                return 0;
            }

            return U16(bytes, classOffset + (glyphId - start) * 2);
        }

        if (format != 2)
        {
            return 0;
        }

        ushort rangeCount = U16(bytes, offset + 2);
        if (offset + 4 + rangeCount * 6 > tableEnd)
        {
            return 0;
        }

        for (int i = 0; i < rangeCount; i++)
        {
            int range = offset + 4 + i * 6;
            ushort start = U16(bytes, range);
            ushort end = U16(bytes, range + 2);
            if (glyphId >= start && glyphId <= end)
            {
                return U16(bytes, range + 4);
            }
        }

        return 0;
    }

    private static Dictionary<ushort, ushort[]> ReadClassGlyphs(byte[] bytes, int offset, int tableEnd)
    {
        var classes = new Dictionary<ushort, List<ushort>>();
        if (offset + 4 > tableEnd)
        {
            return [];
        }

        ushort format = U16(bytes, offset);
        if (format == 1)
        {
            ushort start = U16(bytes, offset + 2);
            ushort count = U16(bytes, offset + 4);
            int classOffset = offset + 6;
            if (classOffset + count * 2 > tableEnd)
            {
                return [];
            }

            for (int i = 0; i < count; i++)
            {
                ushort classValue = U16(bytes, classOffset + i * 2);
                if (classValue == 0)
                {
                    continue;
                }

                AddClassGlyph(classes, classValue, (ushort)(start + i));
            }
        }
        else if (format == 2)
        {
            ushort rangeCount = U16(bytes, offset + 2);
            if (offset + 4 + rangeCount * 6 > tableEnd)
            {
                return [];
            }

            for (int i = 0; i < rangeCount; i++)
            {
                int range = offset + 4 + i * 6;
                ushort start = U16(bytes, range);
                ushort end = U16(bytes, range + 2);
                ushort classValue = U16(bytes, range + 4);
                if (classValue == 0)
                {
                    continue;
                }

                for (int glyph = start; glyph <= end; glyph++)
                {
                    AddClassGlyph(classes, classValue, (ushort)glyph);
                }
            }
        }

        return classes.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray());
    }

    private static void AddClassGlyph(Dictionary<ushort, List<ushort>> classes, ushort classValue, ushort glyphId)
    {
        if (!classes.TryGetValue(classValue, out List<ushort>? glyphs))
        {
            glyphs = [];
            classes[classValue] = glyphs;
        }

        glyphs.Add(glyphId);
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

    private static void W32(byte[] bytes, int offset, uint value)
    {
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(offset, 4), value);
    }

    private static double FixedToDouble(int value)
    {
        return value / 65536d;
    }

    private static byte[] ExtractCollectionFont(byte[] bytes, int fontIndex)
    {
        if (bytes.Length < 16)
        {
            throw new InvalidDataException("TrueType collection header is too small.");
        }

        uint fontCount = U32(bytes, 8);
        if (fontIndex < 0 || fontIndex >= fontCount)
        {
            throw new InvalidDataException("TrueType collection font index is out of range.");
        }

        uint fontOffset = U32(bytes, 12 + fontIndex * 4);
        if (fontOffset > bytes.Length - 12)
        {
            throw new InvalidDataException("TrueType collection font offset is invalid.");
        }

        ushort numTables = U16(bytes, (int)fontOffset + 4);
        int directoryLength = 12 + numTables * 16;
        if (fontOffset + directoryLength > bytes.Length)
        {
            throw new InvalidDataException("TrueType collection table directory is invalid.");
        }

        var records = new CollectionTableRecord[numTables];
        int directoryOffset = (int)fontOffset + 12;
        for (int i = 0; i < numTables; i++)
        {
            int recordOffset = directoryOffset + i * 16;
            uint tableOffset = U32(bytes, recordOffset + 8);
            uint tableLength = U32(bytes, recordOffset + 12);
            if (tableOffset + tableLength > bytes.Length)
            {
                throw new InvalidDataException("TrueType collection table exceeds file length.");
            }

            records[i] = new CollectionTableRecord(tableOffset, tableLength);
        }

        int outputLength = directoryLength;
        foreach (CollectionTableRecord record in records)
        {
            outputLength = Align4(outputLength);
            outputLength += (int)record.Length;
        }

        var output = new byte[Align4(outputLength)];
        Array.Copy(bytes, (int)fontOffset, output, 0, directoryLength);

        int writeOffset = directoryLength;
        for (int i = 0; i < records.Length; i++)
        {
            CollectionTableRecord record = records[i];
            writeOffset = Align4(writeOffset);
            Array.Copy(bytes, (int)record.Offset, output, writeOffset, (int)record.Length);
            W32(output, 12 + i * 16 + 8, (uint)writeOffset);
            writeOffset += (int)record.Length;
        }

        return output;
    }

    private static int Align4(int value)
    {
        return (value + 3) & ~3;
    }

    private readonly record struct TableRecord(int Offset, int Length);

    private readonly record struct CollectionTableRecord(uint Offset, uint Length);

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
