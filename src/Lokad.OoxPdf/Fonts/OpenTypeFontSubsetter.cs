using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed record OpenTypeFontSubset(byte[] Bytes, IReadOnlyDictionary<ushort, ushort> CidByOriginalGlyph);

internal static class OpenTypeFontSubsetter
{
    private const ushort Arg1And2AreWords = 0x0001;
    private const ushort MoreComponents = 0x0020;
    private const ushort WeHaveAScale = 0x0008;
    private const ushort WeHaveAnXAndYScale = 0x0040;
    private const ushort WeHaveATwoByTwo = 0x0080;
    private const ushort WeHaveInstructions = 0x0100;

    public static OpenTypeFontSubset? Create(OpenTypeFont font, IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph, CancellationToken cancellationToken = default)
    {
        try
        {
            return CreateCore(font, unicodeByOriginalGlyph, cancellationToken);
        }
        catch (InvalidDataException)
        {
            return null;
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (OverflowException)
        {
            return null;
        }
    }

    private static OpenTypeFontSubset? CreateCore(OpenTypeFont font, IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] source = font.Bytes.ToArray();
        Dictionary<string, TableRecord> tables = ReadTableDirectory(source);
        if (!tables.TryGetValue("glyf", out TableRecord glyf) ||
            !tables.TryGetValue("loca", out TableRecord loca) ||
            !tables.TryGetValue("head", out TableRecord head) ||
            !tables.TryGetValue("hhea", out TableRecord hhea) ||
            !tables.TryGetValue("hmtx", out TableRecord hmtx) ||
            !tables.TryGetValue("maxp", out TableRecord maxp) ||
            !tables.TryGetValue("name", out TableRecord name) ||
            !tables.TryGetValue("OS/2", out TableRecord os2) ||
            !tables.TryGetValue("post", out TableRecord post))
        {
            return null;
        }

        if (head.Length < 54 || hhea.Length < 36 || maxp.Length < 6 || post.Length < 16)
        {
            return null;
        }

        ushort originalGlyphCount = U16(source, maxp.Offset + 4);
        if (originalGlyphCount == 0)
        {
            return null;
        }

        short indexToLocFormat = I16(source, head.Offset + 50);
        if (indexToLocFormat is not 0 and not 1)
        {
            return null;
        }

        ushort[] originalGlyphs = BuildGlyphClosure(source, originalGlyphCount, indexToLocFormat, glyf, loca, unicodeByOriginalGlyph.Keys, cancellationToken);
        if (originalGlyphs.Length <= 1)
        {
            return null;
        }

        var cidByOriginalGlyph = new Dictionary<ushort, ushort>(originalGlyphs.Length);
        for (int i = 0; i < originalGlyphs.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            cidByOriginalGlyph[originalGlyphs[i]] = (ushort)i;
        }

        byte[] subsetGlyf = BuildGlyf(source, originalGlyphs, cidByOriginalGlyph, originalGlyphCount, indexToLocFormat, glyf, loca, out byte[] subsetLoca, cancellationToken);
        byte[] subsetHhea = CopyTable(source, hhea);
        byte[] subsetHmtx = BuildHmtx(source, originalGlyphs, originalGlyphCount, hhea, hmtx, subsetHhea);
        byte[] subsetMaxp = CopyTable(source, maxp);
        W16(subsetMaxp, 4, (ushort)originalGlyphs.Length);
        byte[] subsetHead = CopyTable(source, head);
        W32(subsetHead, 8, 0);
        W16(subsetHead, 50, 1);
        byte[] subsetPost = BuildPost(source, post);
        byte[] subsetCmap = BuildCmap(unicodeByOriginalGlyph, cidByOriginalGlyph);

        var outputTables = new List<TableData>
        {
            new("OS/2", CopyTable(source, os2)),
            new("cmap", subsetCmap),
            new("glyf", subsetGlyf),
            new("head", subsetHead),
            new("hhea", subsetHhea),
            new("hmtx", subsetHmtx),
            new("loca", subsetLoca),
            new("maxp", subsetMaxp),
            new("name", CopyTable(source, name)),
            new("post", subsetPost)
        };

        foreach (string optionalTag in new[] { "cvt ", "fpgm", "prep", "gasp" })
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (tables.TryGetValue(optionalTag, out TableRecord optional))
            {
                outputTables.Add(new TableData(optionalTag, CopyTable(source, optional)));
            }
        }

        byte[] subsetBytes = BuildSfnt(source.AsSpan(0, 4), outputTables, cancellationToken);
        return new OpenTypeFontSubset(subsetBytes, cidByOriginalGlyph);
    }

    private static Dictionary<string, TableRecord> ReadTableDirectory(byte[] bytes)
    {
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        ushort tableCount = U16(bytes, 4);
        if (12 + tableCount * 16 > bytes.Length)
        {
            throw new InvalidDataException("Font table directory exceeds file length.");
        }

        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        int recordOffset = 12;
        for (int i = 0; i < tableCount; i++)
        {
            string tag = Encoding.ASCII.GetString(bytes, recordOffset, 4);
            uint tableOffset = U32(bytes, recordOffset + 8);
            uint tableLength = U32(bytes, recordOffset + 12);
            if (tableOffset > int.MaxValue || tableLength > int.MaxValue || tableOffset + tableLength > bytes.Length)
            {
                throw new InvalidDataException("Font table exceeds file length.");
            }

            tables[tag] = new TableRecord((int)tableOffset, (int)tableLength);
            recordOffset += 16;
        }

        return tables;
    }

    private static ushort[] BuildGlyphClosure(
        byte[] source,
        ushort originalGlyphCount,
        short indexToLocFormat,
        TableRecord glyf,
        TableRecord loca,
        IEnumerable<ushort> seedGlyphs,
        CancellationToken cancellationToken)
    {
        var included = new SortedSet<ushort> { 0 };
        var pending = new Queue<ushort>();
        pending.Enqueue(0);
        foreach (ushort glyph in seedGlyphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (glyph < originalGlyphCount && included.Add(glyph))
            {
                pending.Enqueue(glyph);
            }
        }

        while (pending.Count != 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = pending.Dequeue();
            if (!TryReadCompoundComponents(source, glyph, originalGlyphCount, indexToLocFormat, glyf, loca, out ushort[] components))
            {
                throw new InvalidDataException("Compound glyph is malformed.");
            }

            foreach (ushort component in components)
            {
                if (component < originalGlyphCount && included.Add(component))
                {
                    pending.Enqueue(component);
                }
            }
        }

        return included.ToArray();
    }

    private static byte[] BuildGlyf(
        byte[] source,
        ushort[] originalGlyphs,
        IReadOnlyDictionary<ushort, ushort> cidByOriginalGlyph,
        ushort originalGlyphCount,
        short indexToLocFormat,
        TableRecord glyf,
        TableRecord loca,
        out byte[] subsetLoca,
        CancellationToken cancellationToken)
    {
        using var glyfStream = new MemoryStream();
        subsetLoca = new byte[(originalGlyphs.Length + 1) * 4];
        for (int i = 0; i < originalGlyphs.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int alignedOffset = Align4((int)glyfStream.Position);
            while (glyfStream.Position < alignedOffset)
            {
                glyfStream.WriteByte(0);
            }

            W32(subsetLoca, i * 4, (uint)glyfStream.Position);
            if (!TryGetGlyphData(source, originalGlyphs[i], originalGlyphCount, indexToLocFormat, glyf, loca, out int glyphOffset, out int glyphLength))
            {
                throw new InvalidDataException("Glyph range is invalid.");
            }

            byte[] glyphData = new byte[glyphLength];
            Array.Copy(source, glyphOffset, glyphData, 0, glyphLength);
            if (!RemapCompoundGlyph(glyphData, cidByOriginalGlyph))
            {
                throw new InvalidDataException("Compound glyph remapping failed.");
            }

            glyfStream.Write(glyphData);
        }

        int finalOffset = Align4((int)glyfStream.Position);
        while (glyfStream.Position < finalOffset)
        {
            glyfStream.WriteByte(0);
        }

        W32(subsetLoca, originalGlyphs.Length * 4, (uint)glyfStream.Position);
        return glyfStream.ToArray();
    }

    private static byte[] BuildHmtx(byte[] source, ushort[] originalGlyphs, ushort originalGlyphCount, TableRecord hhea, TableRecord hmtx, byte[] subsetHhea)
    {
        ushort numberOfHMetrics = U16(source, hhea.Offset + 34);
        if (numberOfHMetrics == 0)
        {
            throw new InvalidDataException("Horizontal metrics table is empty.");
        }

        W16(subsetHhea, 34, (ushort)originalGlyphs.Length);
        var output = new byte[originalGlyphs.Length * 4];
        for (int i = 0; i < originalGlyphs.Length; i++)
        {
            ushort glyph = originalGlyphs[i];
            if (!TryReadHorizontalMetric(source, glyph, originalGlyphCount, numberOfHMetrics, hmtx, out ushort advance, out short leftSideBearing))
            {
                throw new InvalidDataException("Horizontal metric is invalid.");
            }

            W16(output, i * 4, advance);
            W16(output, i * 4 + 2, unchecked((ushort)leftSideBearing));
        }

        return output;
    }

    private static bool TryReadHorizontalMetric(
        byte[] source,
        ushort glyph,
        ushort originalGlyphCount,
        ushort numberOfHMetrics,
        TableRecord hmtx,
        out ushort advance,
        out short leftSideBearing)
    {
        advance = 0;
        leftSideBearing = 0;
        if (glyph >= originalGlyphCount)
        {
            return false;
        }

        if (glyph < numberOfHMetrics)
        {
            int offset = hmtx.Offset + glyph * 4;
            if (offset + 4 > hmtx.Offset + hmtx.Length)
            {
                return false;
            }

            advance = U16(source, offset);
            leftSideBearing = I16(source, offset + 2);
            return true;
        }

        int advanceOffset = hmtx.Offset + (numberOfHMetrics - 1) * 4;
        int lsbOffset = hmtx.Offset + numberOfHMetrics * 4 + (glyph - numberOfHMetrics) * 2;
        if (advanceOffset + 2 > hmtx.Offset + hmtx.Length || lsbOffset + 2 > hmtx.Offset + hmtx.Length)
        {
            return false;
        }

        advance = U16(source, advanceOffset);
        leftSideBearing = I16(source, lsbOffset);
        return true;
    }

    private static byte[] BuildPost(byte[] source, TableRecord post)
    {
        var output = new byte[32];
        Array.Copy(source, post.Offset, output, 0, Math.Min(post.Length, output.Length));
        W32(output, 0, 0x00030000);
        return output;
    }

    private static byte[] BuildCmap(IReadOnlyDictionary<ushort, int> unicodeByOriginalGlyph, IReadOnlyDictionary<ushort, ushort> cidByOriginalGlyph)
    {
        var glyphByCodePoint = new SortedDictionary<int, ushort>();
        foreach ((ushort originalGlyph, int codePoint) in unicodeByOriginalGlyph)
        {
            if (codePoint < 0 || codePoint > 0x10FFFF || codePoint is >= 0xD800 and <= 0xDFFF)
            {
                continue;
            }

            if (cidByOriginalGlyph.TryGetValue(originalGlyph, out ushort cid))
            {
                glyphByCodePoint.TryAdd(codePoint, cid);
            }
        }

        byte[] format4 = BuildCmapFormat4(glyphByCodePoint);
        byte[] format12 = BuildCmapFormat12(glyphByCodePoint);
        int format4Offset = 4 + 8 * 2;
        int format12Offset = format4Offset + format4.Length;
        var output = new byte[format12Offset + format12.Length];
        W16(output, 0, 0);
        W16(output, 2, 2);
        W16(output, 4, 3);
        W16(output, 6, 1);
        W32(output, 8, (uint)format4Offset);
        W16(output, 12, 3);
        W16(output, 14, 10);
        W32(output, 16, (uint)format12Offset);
        Array.Copy(format4, 0, output, format4Offset, format4.Length);
        Array.Copy(format12, 0, output, format12Offset, format12.Length);
        return output;
    }

    private static byte[] BuildCmapFormat4(SortedDictionary<int, ushort> glyphByCodePoint)
    {
        KeyValuePair<int, ushort>[] bmp = glyphByCodePoint
            .Where(pair => pair.Key <= 0xFFFF && pair.Key != 0xFFFF)
            .ToArray();
        int segmentCount = bmp.Length + 1;
        int length = 16 + segmentCount * 8;
        var output = new byte[length];
        W16(output, 0, 4);
        W16(output, 2, (ushort)length);
        W16(output, 4, 0);
        W16(output, 6, (ushort)(segmentCount * 2));
        WriteSearchFields(output, 8, segmentCount, 2);

        int endCodes = 14;
        int startCodes = endCodes + segmentCount * 2 + 2;
        int idDeltas = startCodes + segmentCount * 2;
        int idRangeOffsets = idDeltas + segmentCount * 2;
        for (int i = 0; i < bmp.Length; i++)
        {
            int codePoint = bmp[i].Key;
            ushort glyph = bmp[i].Value;
            W16(output, endCodes + i * 2, (ushort)codePoint);
            W16(output, startCodes + i * 2, (ushort)codePoint);
            W16(output, idDeltas + i * 2, unchecked((ushort)(glyph - codePoint)));
            W16(output, idRangeOffsets + i * 2, 0);
        }

        int sentinel = segmentCount - 1;
        W16(output, endCodes + sentinel * 2, 0xFFFF);
        W16(output, startCodes + sentinel * 2, 0xFFFF);
        W16(output, idDeltas + sentinel * 2, 1);
        W16(output, idRangeOffsets + sentinel * 2, 0);
        return output;
    }

    private static byte[] BuildCmapFormat12(SortedDictionary<int, ushort> glyphByCodePoint)
    {
        var groups = new List<CmapGroup>();
        foreach ((int codePoint, ushort glyph) in glyphByCodePoint)
        {
            groups.Add(new CmapGroup((uint)codePoint, (uint)codePoint, glyph));
        }

        int length = 16 + groups.Count * 12;
        var output = new byte[length];
        W16(output, 0, 12);
        W16(output, 2, 0);
        W32(output, 4, (uint)length);
        W32(output, 8, 0);
        W32(output, 12, (uint)groups.Count);
        int offset = 16;
        foreach (CmapGroup group in groups)
        {
            W32(output, offset, group.StartCode);
            W32(output, offset + 4, group.EndCode);
            W32(output, offset + 8, group.StartGlyph);
            offset += 12;
        }

        return output;
    }

    private static byte[] BuildSfnt(ReadOnlySpan<byte> scalerType, List<TableData> tables, CancellationToken cancellationToken)
    {
        tables.Sort((left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        int tableCount = tables.Count;
        int directoryLength = 12 + tableCount * 16;
        int outputLength = directoryLength;
        foreach (TableData table in tables)
        {
            outputLength = Align4(outputLength);
            outputLength += table.Data.Length;
        }

        var output = new byte[Align4(outputLength)];
        scalerType.CopyTo(output);
        W16(output, 4, (ushort)tableCount);
        WriteSearchFields(output, 6, tableCount, 16);

        int dataOffset = directoryLength;
        int directoryOffset = 12;
        int headOffset = -1;
        foreach (TableData table in tables)
        {
            cancellationToken.ThrowIfCancellationRequested();
            dataOffset = Align4(dataOffset);
            Encoding.ASCII.GetBytes(table.Tag, output.AsSpan(directoryOffset, 4));
            W32(output, directoryOffset + 4, Checksum(table.Data));
            W32(output, directoryOffset + 8, (uint)dataOffset);
            W32(output, directoryOffset + 12, (uint)table.Data.Length);
            Array.Copy(table.Data, 0, output, dataOffset, table.Data.Length);
            if (table.Tag.Equals("head", StringComparison.Ordinal))
            {
                headOffset = dataOffset;
            }

            dataOffset += table.Data.Length;
            directoryOffset += 16;
        }

        if (headOffset < 0)
        {
            throw new InvalidDataException("Subset font is missing head table.");
        }

        W32(output, headOffset + 8, 0);
        uint adjustment = unchecked(0xB1B0AFBAu - Checksum(output));
        W32(output, headOffset + 8, adjustment);
        return output;
    }

    private static bool TryReadCompoundComponents(
        byte[] source,
        ushort glyph,
        ushort originalGlyphCount,
        short indexToLocFormat,
        TableRecord glyf,
        TableRecord loca,
        out ushort[] components)
    {
        components = [];
        if (!TryGetGlyphData(source, glyph, originalGlyphCount, indexToLocFormat, glyf, loca, out int glyphOffset, out int glyphLength))
        {
            return false;
        }

        if (glyphLength == 0)
        {
            return true;
        }

        if (glyphLength < 10)
        {
            return false;
        }

        short contourCount = I16(source, glyphOffset);
        if (contourCount >= 0)
        {
            return true;
        }

        var result = new List<ushort>();
        int cursor = glyphOffset + 10;
        int glyphEnd = glyphOffset + glyphLength;
        ushort flags;
        do
        {
            if (cursor + 4 > glyphEnd)
            {
                return false;
            }

            flags = U16(source, cursor);
            result.Add(U16(source, cursor + 2));
            cursor += 4;
            cursor += (flags & Arg1And2AreWords) != 0 ? 4 : 2;
            if ((flags & WeHaveAScale) != 0)
            {
                cursor += 2;
            }
            else if ((flags & WeHaveAnXAndYScale) != 0)
            {
                cursor += 4;
            }
            else if ((flags & WeHaveATwoByTwo) != 0)
            {
                cursor += 8;
            }

            if (cursor > glyphEnd)
            {
                return false;
            }
        }
        while ((flags & MoreComponents) != 0);

        if ((flags & WeHaveInstructions) != 0)
        {
            if (cursor + 2 > glyphEnd)
            {
                return false;
            }

            ushort instructionLength = U16(source, cursor);
            if (cursor + 2 + instructionLength > glyphEnd)
            {
                return false;
            }
        }

        components = result.ToArray();
        return true;
    }

    private static bool RemapCompoundGlyph(byte[] glyphData, IReadOnlyDictionary<ushort, ushort> cidByOriginalGlyph)
    {
        if (glyphData.Length == 0)
        {
            return true;
        }

        if (glyphData.Length < 10)
        {
            return false;
        }

        short contourCount = I16(glyphData, 0);
        if (contourCount >= 0)
        {
            return true;
        }

        int cursor = 10;
        ushort flags;
        do
        {
            if (cursor + 4 > glyphData.Length)
            {
                return false;
            }

            flags = U16(glyphData, cursor);
            ushort originalComponent = U16(glyphData, cursor + 2);
            if (!cidByOriginalGlyph.TryGetValue(originalComponent, out ushort subsetComponent))
            {
                return false;
            }

            W16(glyphData, cursor + 2, subsetComponent);
            cursor += 4;
            cursor += (flags & Arg1And2AreWords) != 0 ? 4 : 2;
            if ((flags & WeHaveAScale) != 0)
            {
                cursor += 2;
            }
            else if ((flags & WeHaveAnXAndYScale) != 0)
            {
                cursor += 4;
            }
            else if ((flags & WeHaveATwoByTwo) != 0)
            {
                cursor += 8;
            }

            if (cursor > glyphData.Length)
            {
                return false;
            }
        }
        while ((flags & MoreComponents) != 0);

        return true;
    }

    private static bool TryGetGlyphData(
        byte[] source,
        ushort glyph,
        ushort originalGlyphCount,
        short indexToLocFormat,
        TableRecord glyf,
        TableRecord loca,
        out int glyphOffset,
        out int glyphLength)
    {
        glyphOffset = 0;
        glyphLength = 0;
        if (glyph >= originalGlyphCount)
        {
            return false;
        }

        int start;
        int end;
        if (indexToLocFormat == 0)
        {
            int entry = loca.Offset + glyph * 2;
            if (entry + 4 > loca.Offset + loca.Length)
            {
                return false;
            }

            start = U16(source, entry) * 2;
            end = U16(source, entry + 2) * 2;
        }
        else
        {
            int entry = loca.Offset + glyph * 4;
            if (entry + 8 > loca.Offset + loca.Length)
            {
                return false;
            }

            uint start32 = U32(source, entry);
            uint end32 = U32(source, entry + 4);
            if (start32 > int.MaxValue || end32 > int.MaxValue)
            {
                return false;
            }

            start = (int)start32;
            end = (int)end32;
        }

        if (end < start || glyf.Offset + end > glyf.Offset + glyf.Length)
        {
            return false;
        }

        glyphOffset = glyf.Offset + start;
        glyphLength = end - start;
        return true;
    }

    private static byte[] CopyTable(byte[] source, TableRecord table)
    {
        var output = new byte[table.Length];
        Array.Copy(source, table.Offset, output, 0, table.Length);
        return output;
    }

    private static void WriteSearchFields(byte[] output, int offset, int count, int unitSize)
    {
        int power = 1;
        int entrySelector = 0;
        while (power * 2 <= count)
        {
            power *= 2;
            entrySelector++;
        }

        int searchRange = power * unitSize;
        W16(output, offset, (ushort)searchRange);
        W16(output, offset + 2, (ushort)entrySelector);
        W16(output, offset + 4, (ushort)(count * unitSize - searchRange));
    }

    private static uint Checksum(byte[] bytes)
    {
        uint sum = 0;
        for (int i = 0; i < bytes.Length; i += 4)
        {
            uint value = 0;
            for (int j = 0; j < 4; j++)
            {
                value <<= 8;
                if (i + j < bytes.Length)
                {
                    value |= bytes[i + j];
                }
            }

            sum = unchecked(sum + value);
        }

        return sum;
    }

    private static int Align4(int value)
    {
        return (value + 3) & ~3;
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

    private static void W16(byte[] bytes, int offset, ushort value)
    {
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset, 2), value);
    }

    private static void W32(byte[] bytes, int offset, uint value)
    {
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(offset, 4), value);
    }

    private readonly record struct TableRecord(int Offset, int Length);

    private readonly record struct TableData(string Tag, byte[] Data);

    private readonly record struct CmapGroup(uint StartCode, uint EndCode, uint StartGlyph);
}
