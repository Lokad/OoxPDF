using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
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
        IReadOnlyDictionary<uint, short> kerningPairs,
        string scalerTag)
    {
        this.bytes = bytes;
        this.tables = tables;
        ScalerTag = scalerTag;
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

    public string ScalerTag { get; }

    public bool HasTrueTypeOutlines => tables.ContainsKey("glyf") && tables.ContainsKey("loca");

    public bool HasCffOutlines => tables.ContainsKey("CFF ") || tables.ContainsKey("CFF2");

    public string FamilyName { get; }

    public ushort UnitsPerEm { get; }

    public FontBounds Bounds { get; }

    public ReadOnlyMemory<byte> Bytes => bytes;

    public ushort GlyphCount { get; }

    public Os2Metrics Os2 { get; }

    public PostMetrics Post { get; }

    public IReadOnlyCollection<string> TableTags => tables.Keys;

    public static OpenTypeFont Load(byte[] bytes)
    {
        return Load(bytes, 0);
    }

    public static OpenTypeFont Load(byte[] bytes, int fontIndex)
    {
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        if (Encoding.ASCII.GetString(bytes, 0, 4) == "ttcf")
        {
            bytes = ExtractCollectionFont(bytes, fontIndex);
        }
        else if (fontIndex != 0)
        {
            throw new InvalidDataException("Font index can only be non-zero for TrueType collections.");
        }

        string scalerTag = ReadScalerTag(bytes);
        try
        {
        Dictionary<string, TableRecord> tables = ReadTableDirectory(bytes, 12, U16(bytes, 4), "font");
        RequireMinimumLength(tables, "head", 54);
        RequireMinimumLength(tables, "hhea", 36);
        RequireMinimumLength(tables, "maxp", 6);
        RequireMinimumLength(tables, "name", 6);
        RequireMinimumLength(tables, "OS/2", 78);
        RequireMinimumLength(tables, "post", 16);
        RequireMinimumLength(tables, "cmap", 4);
        RequireHmtxLength(bytes, tables);

        ushort unitsPerEm = ReadUnitsPerEm(bytes, tables);
        if (unitsPerEm == 0)
        {
            throw new InvalidDataException("Font has zero units-per-em.");
        }
        FontBounds bounds = ReadBounds(bytes, tables);
        ushort glyphCount = ReadGlyphCount(bytes, tables);
        string familyName = ReadFamilyName(bytes, tables);
        Os2Metrics os2 = ReadOs2(bytes, tables);
        PostMetrics post = ReadPost(bytes, tables);
        CmapFormat? cmap = ReadCmap(bytes, tables);
        ushort[] advances = ReadAdvances(bytes, tables);
        IReadOnlyDictionary<uint, short> kerningPairs = ReadKerningPairs(bytes, tables);
        return new OpenTypeFont(bytes, tables, familyName, unitsPerEm, bounds, glyphCount, os2, post, cmap, advances, kerningPairs, scalerTag);

        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
        {
            throw new InvalidDataException("Font file is malformed: " + ex.Message, ex);
        }
    }

    // Discovery headers (G04): everything font discovery needs without kern-pair
    // expansion, the dominant parse transient on big faces. Same validation and header
    // reads as Load, so malformed files are accepted/rejected identically; kern-only
    // defects surface at first full load through the existing fallback instead.
    internal readonly record struct FontDiscoveryHeaders(
        string FamilyName,
        ushort WeightClass,
        double ItalicAngle,
        bool HasMathTable);

    internal static FontDiscoveryHeaders ReadDiscoveryHeaders(byte[] bytes, int fontIndex)
    {
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        if (Encoding.ASCII.GetString(bytes, 0, 4) == "ttcf")
        {
            bytes = ExtractCollectionFont(bytes, fontIndex);
        }
        else if (fontIndex != 0)
        {
            throw new InvalidDataException("Font index can only be non-zero for TrueType collections.");
        }

        ReadScalerTag(bytes);
        try
        {
        Dictionary<string, TableRecord> tables = ReadTableDirectory(bytes, 12, U16(bytes, 4), "font");
        RequireMinimumLength(tables, "head", 54);
        RequireMinimumLength(tables, "hhea", 36);
        RequireMinimumLength(tables, "maxp", 6);
        RequireMinimumLength(tables, "name", 6);
        RequireMinimumLength(tables, "OS/2", 78);
        RequireMinimumLength(tables, "post", 16);
        RequireMinimumLength(tables, "cmap", 4);
        RequireHmtxLength(bytes, tables);

        ushort unitsPerEm = ReadUnitsPerEm(bytes, tables);
        if (unitsPerEm == 0)
        {
            throw new InvalidDataException("Font has zero units-per-em.");
        }
        string familyName = ReadFamilyName(bytes, tables);
        Os2Metrics os2 = ReadOs2(bytes, tables);
        PostMetrics post = ReadPost(bytes, tables);
        // Parsed and discarded: keeps malformed-file rejection identical to Load
        // without retaining anything beyond the discovery headers.
        ReadBounds(bytes, tables);
        ReadGlyphCount(bytes, tables);
        ReadCmap(bytes, tables);
        ReadAdvances(bytes, tables);
        return new FontDiscoveryHeaders(familyName, os2.WeightClass, post.ItalicAngle, tables.ContainsKey("MATH"));
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
        {
            throw new InvalidDataException("Font file is malformed: " + ex.Message, ex);
        }
    }

    private static byte[] ExtractCollectionFont(byte[] bytes, int fontIndex)
        {
            if (bytes.Length < 16)
            {
                throw new InvalidDataException("TrueType collection header is too small.");
            }

            uint fontCount = U32(bytes, 8);
            if (fontCount > 256)
            {
                throw new InvalidDataException("TrueType collection declares too many faces.");
            }
            if (fontIndex < 0 || fontIndex >= fontCount)
            {
                throw new InvalidDataException("TrueType collection font index is out of range.");
            }

            uint fontOffset = U32(bytes, 12 + fontIndex * 4);
            if (fontOffset > bytes.Length - 12)
            {
                throw new InvalidDataException("TrueType collection font offset is invalid.");
            }

            ushort collectionTableCount = U16(bytes, (int)fontOffset + 4);
            int directoryLength = 12 + collectionTableCount * 16;
            if ((ulong)fontOffset + (ulong)(uint)directoryLength > (ulong)bytes.Length)
            {
                throw new InvalidDataException("TrueType collection table directory is invalid.");
            }

            // Validate and dedupe through the shared directory reader, then keep
            // directory order for the byte-exact rebuild below.
            Dictionary<string, TableRecord> faceTables = ReadTableDirectory(bytes, (int)fontOffset + 12, collectionTableCount, "collection face");
            var records = new TableRecord[collectionTableCount];
            for (int i = 0; i < collectionTableCount; i++)
            {
                int recordOffset = (int)fontOffset + 12 + i * 16;
                string tag = Encoding.ASCII.GetString(bytes, recordOffset, 4);
                records[i] = faceTables[tag];
            }

            long outputLength = directoryLength;
            foreach (TableRecord record in records)
            {
                outputLength = (outputLength + 3L) & ~3L;
                outputLength = checked(outputLength + record.Length);
            }

            // Extraction only repackages source bytes plus alignment padding, so a
            // larger claim implies overlapping or hostile table extents.
            if (outputLength > (long)bytes.Length + 3L * collectionTableCount + 4L)
            {
                throw new InvalidDataException("TrueType collection face exceeds file length.");
            }

            var output = new byte[Align4((int)outputLength)];
            Array.Copy(bytes, (int)fontOffset, output, 0, directoryLength);

            int writeOffset = directoryLength;
            for (int i = 0; i < records.Length; i++)
            {
                TableRecord record = records[i];
                writeOffset = Align4(writeOffset);
                Array.Copy(bytes, record.Offset, output, writeOffset, record.Length);
                W32(output, 12 + i * 16 + 8, (uint)writeOffset);
                writeOffset += record.Length;
            }
            return output;
        }

    public static int GetCollectionFontCount(byte[] bytes)
    {
        if (bytes.Length < 12 || Encoding.ASCII.GetString(bytes, 0, 4) != "ttcf")
        {
            return 1;
        }

        if (bytes.Length < 16)
        {
            throw new InvalidDataException("TrueType collection header is too small.");
        }

        uint fontCount = U32(bytes, 8);
        return fontCount > int.MaxValue ? throw new InvalidDataException("TrueType collection font count is too large.") : (int)fontCount;
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

    public bool TryReadGlyphOutline(ushort glyphId, out OpenTypeGlyphOutline outline)
    {
        return TryReadGlyphOutline(glyphId, depth: 0, out outline);
    }

    private bool TryReadGlyphOutline(ushort glyphId, int depth, out OpenTypeGlyphOutline outline)
    {
        outline = default;
        if (depth > 16 ||
            glyphId >= GlyphCount ||
            !tables.TryGetValue("head", out TableRecord head) ||
            !tables.TryGetValue("loca", out TableRecord loca) ||
            !tables.TryGetValue("glyf", out TableRecord glyf) ||
            head.Length < 52)
        {
            return false;
        }

        short indexToLocFormat = I16(bytes, head.Offset + 50);
        if (!TryGetGlyphTableRange(glyphId, indexToLocFormat, loca, glyf, out int glyphOffset, out int glyphLength) ||
            glyphLength < 10)
        {
            return false;
        }

        int glyphEnd = glyphOffset + glyphLength;
        short contourCount = I16(bytes, glyphOffset);
        FontBounds glyphBounds = new(
            I16(bytes, glyphOffset + 2),
            I16(bytes, glyphOffset + 4),
            I16(bytes, glyphOffset + 6),
            I16(bytes, glyphOffset + 8));

        if (contourCount < 0)
        {
            return TryReadCompoundGlyphOutline(glyphOffset + 10, glyphEnd, glyphBounds, depth, out outline);
        }

        return TryReadSimpleGlyphOutline(glyphOffset + 10, glyphEnd, contourCount, glyphBounds, out outline);
    }

    private bool TryGetGlyphTableRange(
        ushort glyphId,
        short indexToLocFormat,
        TableRecord loca,
        TableRecord glyf,
        out int glyphOffset,
        out int glyphLength)
    {
        glyphOffset = 0;
        glyphLength = 0;

        if (indexToLocFormat == 0)
        {
            int locaOffset = loca.Offset + glyphId * 2;
            if (locaOffset + 4 > loca.Offset + loca.Length)
            {
                return false;
            }

            glyphOffset = glyf.Offset + U16(bytes, locaOffset) * 2;
            int nextOffset = glyf.Offset + U16(bytes, locaOffset + 2) * 2;
            glyphLength = nextOffset - glyphOffset;
        }
        else if (indexToLocFormat == 1)
        {
            int locaOffset = loca.Offset + glyphId * 4;
            if (locaOffset + 8 > loca.Offset + loca.Length)
            {
                return false;
            }

            uint relativeOffset = U32(bytes, locaOffset);
            uint nextRelativeOffset = U32(bytes, locaOffset + 4);
            if (relativeOffset > int.MaxValue || nextRelativeOffset > int.MaxValue)
            {
                return false;
            }

            glyphOffset = glyf.Offset + (int)relativeOffset;
            int nextOffset = glyf.Offset + (int)nextRelativeOffset;
            glyphLength = nextOffset - glyphOffset;
        }
        else
        {
            return false;
        }

        return glyphLength > 0 &&
            glyphOffset >= glyf.Offset &&
            glyphOffset + glyphLength <= glyf.Offset + glyf.Length;
    }

    private bool TryReadSimpleGlyphOutline(
        int offset,
        int glyphEnd,
        short contourCount,
        FontBounds bounds,
        out OpenTypeGlyphOutline outline)
    {
        outline = default;
        if (contourCount == 0)
        {
            outline = new OpenTypeGlyphOutline(bounds, [], IsCompound: false);
            return true;
        }

        if (offset + contourCount * 2 + 2 > glyphEnd)
        {
            return false;
        }

        var contourEnds = new ushort[contourCount];
        for (int i = 0; i < contourEnds.Length; i++)
        {
            contourEnds[i] = U16(bytes, offset + i * 2);
        }

        int pointCount = contourEnds[^1] + 1;
        int instructionLengthOffset = offset + contourCount * 2;
        ushort instructionLength = U16(bytes, instructionLengthOffset);
        int flagsOffset = instructionLengthOffset + 2 + instructionLength;
        if (pointCount <= 0 || flagsOffset > glyphEnd)
        {
            return false;
        }

        var flags = new byte[pointCount];
        int cursor = flagsOffset;
        for (int i = 0; i < pointCount;)
        {
            if (cursor >= glyphEnd)
            {
                return false;
            }

            byte flag = bytes[cursor++];
            int repeatCount = 0;
            if ((flag & 0x08) != 0)
            {
                if (cursor >= glyphEnd)
                {
                    return false;
                }

                repeatCount = bytes[cursor++];
            }

            for (int repeat = 0; repeat <= repeatCount && i < pointCount; repeat++)
            {
                flags[i++] = flag;
            }
        }

        if (!TryReadSimpleGlyphCoordinates(flags, glyphEnd, ref cursor, readX: true, out short[] xs) ||
            !TryReadSimpleGlyphCoordinates(flags, glyphEnd, ref cursor, readX: false, out short[] ys))
        {
            return false;
        }

        var contours = new OpenTypeGlyphContour[contourCount];
        int start = 0;
        for (int i = 0; i < contours.Length; i++)
        {
            int end = contourEnds[i];
            if (end < start || end >= pointCount)
            {
                return false;
            }

            var points = new OpenTypeGlyphPoint[end - start + 1];
            for (int point = start; point <= end; point++)
            {
                points[point - start] = new OpenTypeGlyphPoint(xs[point], ys[point], (flags[point] & 0x01) != 0);
            }

            contours[i] = new OpenTypeGlyphContour(points);
            start = end + 1;
        }

        outline = new OpenTypeGlyphOutline(bounds, contours, IsCompound: false);
        return true;
    }

    private bool TryReadCompoundGlyphOutline(
        int offset,
        int glyphEnd,
        FontBounds bounds,
        int depth,
        out OpenTypeGlyphOutline outline)
    {
        outline = default;
        var contours = new List<OpenTypeGlyphContour>();
        int cursor = offset;
        ushort flags;
        do
        {
            if (cursor + 4 > glyphEnd)
            {
                return false;
            }

            flags = U16(bytes, cursor);
            ushort componentGlyphId = U16(bytes, cursor + 2);
            cursor += 4;

            if ((flags & 0x0002) == 0)
            {
                return false;
            }

            double dx;
            double dy;
            if ((flags & 0x0001) != 0)
            {
                if (cursor + 4 > glyphEnd)
                {
                    return false;
                }

                dx = I16(bytes, cursor);
                dy = I16(bytes, cursor + 2);
                cursor += 4;
            }
            else
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                dx = unchecked((sbyte)bytes[cursor]);
                dy = unchecked((sbyte)bytes[cursor + 1]);
                cursor += 2;
            }

            double a = 1d;
            double b = 0d;
            double c = 0d;
            double d = 1d;
            if ((flags & 0x0008) != 0)
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                a = d = F2Dot14(bytes, cursor);
                cursor += 2;
            }
            else if ((flags & 0x0040) != 0)
            {
                if (cursor + 4 > glyphEnd)
                {
                    return false;
                }

                a = F2Dot14(bytes, cursor);
                d = F2Dot14(bytes, cursor + 2);
                cursor += 4;
            }
            else if ((flags & 0x0080) != 0)
            {
                if (cursor + 8 > glyphEnd)
                {
                    return false;
                }

                a = F2Dot14(bytes, cursor);
                c = F2Dot14(bytes, cursor + 2);
                b = F2Dot14(bytes, cursor + 4);
                d = F2Dot14(bytes, cursor + 6);
                cursor += 8;
            }

            if (!TryReadGlyphOutline(componentGlyphId, depth + 1, out OpenTypeGlyphOutline component))
            {
                return false;
            }

            foreach (OpenTypeGlyphContour contour in component.Contours)
            {
                var points = new OpenTypeGlyphPoint[contour.Points.Count];
                for (int i = 0; i < points.Length; i++)
                {
                    OpenTypeGlyphPoint point = contour.Points[i];
                    points[i] = new OpenTypeGlyphPoint(
                        a * point.X + c * point.Y + dx,
                        b * point.X + d * point.Y + dy,
                        point.IsOnCurve);
                }

                contours.Add(new OpenTypeGlyphContour(points));
            }
        }
        while ((flags & 0x0020) != 0);

        if ((flags & 0x0100) != 0)
        {
            if (cursor + 2 > glyphEnd)
            {
                return false;
            }

            ushort instructionLength = U16(bytes, cursor);
            cursor += 2 + instructionLength;
            if (cursor > glyphEnd)
            {
                return false;
            }
        }

        outline = new OpenTypeGlyphOutline(bounds, contours, IsCompound: true);
        return true;
    }

    private bool TryReadSimpleGlyphCoordinates(byte[] flags, int glyphEnd, ref int cursor, bool readX, out short[] coordinates)
    {
        coordinates = new short[flags.Length];
        int value = 0;
        byte shortVector = readX ? (byte)0x02 : (byte)0x04;
        byte sameOrPositive = readX ? (byte)0x10 : (byte)0x20;

        for (int i = 0; i < flags.Length; i++)
        {
            int delta = 0;
            if ((flags[i] & shortVector) != 0)
            {
                if (cursor >= glyphEnd)
                {
                    return false;
                }

                delta = bytes[cursor++];
                if ((flags[i] & sameOrPositive) == 0)
                {
                    delta = -delta;
                }
            }
            else if ((flags[i] & sameOrPositive) == 0)
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                delta = I16(bytes, cursor);
                cursor += 2;
            }

            value += delta;
            if (value < short.MinValue || value > short.MaxValue)
            {
                return false;
            }

            coordinates[i] = (short)value;
        }

        return true;
    }

    private static string ReadScalerTag(byte[] bytes)
    {
        string tag = Encoding.ASCII.GetString(bytes, 0, 4);
        if (tag is "\0\u0001\0\0" or "OTTO" or "true" or "typ1")
        {
            return tag;
        }

        throw new InvalidDataException("Font has an unrecognized sfnt version.");
    }

    private static Dictionary<string, TableRecord> ReadTableDirectory(byte[] bytes, int directoryOffset, ushort tableCount, string what)
    {
        if ((long)directoryOffset + (long)tableCount * 16L > bytes.Length)
        {
            throw new InvalidDataException("Font table directory exceeds file length.");
        }

        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        for (int i = 0; i < tableCount; i++)
        {
            int record = directoryOffset + i * 16;
            string tag = Encoding.ASCII.GetString(bytes, record, 4);
            uint tableOffset = U32(bytes, record + 8);
            uint length = U32(bytes, record + 12);
            if ((ulong)tableOffset + length > (ulong)bytes.Length)
            {
                throw new InvalidDataException("Font table exceeds file length.");
            }

            if (!tables.TryAdd(tag, new TableRecord((int)tableOffset, (int)length)))
            {
                throw new InvalidDataException("Font table directory contains a duplicate tag.");
            }
        }

        return tables;
    }

    private static void RequireMinimumLength(Dictionary<string, TableRecord> tables, string tag, int minimum)
    {
        TableRecord table = Required(tables, tag);
        if (table.Length < minimum)
        {
            throw new InvalidDataException("Font table is truncated.");
        }
    }

    private static void RequireHmtxLength(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord hhea = Required(tables, "hhea");
        TableRecord hmtx = Required(tables, "hmtx");
        ushort numberOfHMetrics = U16(bytes, hhea.Offset + 34);
        if ((long)hmtx.Length < (long)numberOfHMetrics * 4L)
        {
            throw new InvalidDataException("Font horizontal metrics table is truncated.");
        }
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
            StrikeoutSize: I16(bytes, os2.Offset + 26),
            StrikeoutPosition: I16(bytes, os2.Offset + 28),
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

        double FixedToDouble(int value)
        {
            return value / 65536d;
        }
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
        // Collect first so the map below allocates exactly once: kern-pair
        // expansion (especially class-based GPOS format 2) otherwise regrows the
        // dictionary ~18 times on big faces. Insertion order is preserved, so
        // duplicate keys keep last-wins semantics.
        var collected = new List<(uint Key, short Value)>();
        ReadLegacyKerningPairs(bytes, tables, collected);
        ReadGposPairAdjustments(bytes, tables, collected);
        var pairs = new Dictionary<uint, short>(collected.Count);
        foreach ((uint key, short value) in collected)
        {
            pairs[key] = value;
        }

        return pairs;
    }

    private readonly record struct TableRecord(int Offset, int Length);


    internal readonly record struct Os2Metrics(
        ushort Version,
        ushort WeightClass,
        ushort WidthClass,
        short StrikeoutSize,
        short StrikeoutPosition,
        short TypographicAscender,
        short TypographicDescender,
        short TypographicLineGap,
        ushort WindowsAscender,
        ushort WindowsDescender);

    internal readonly record struct FontBounds(short XMin, short YMin, short XMax, short YMax);

    internal readonly record struct OpenTypeGlyphOutline(FontBounds Bounds, IReadOnlyList<OpenTypeGlyphContour> Contours, bool IsCompound);

    internal readonly record struct OpenTypeGlyphContour(IReadOnlyList<OpenTypeGlyphPoint> Points);

    internal readonly record struct OpenTypeGlyphPoint(double X, double Y, bool IsOnCurve);

    internal readonly record struct PostMetrics(
        double ItalicAngle,
        short UnderlinePosition,
        short UnderlineThickness,
        bool IsFixedPitch);

}
