using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

// Builds a minimal but structurally valid TrueType font in memory so core
// font-format, subset, CMap, and width checks run identically on every host
// (Q01): no Windows/Office fonts required. Glyph 0 is .notdef; codepoints
// U+0020..U+009F map to glyphs 1..128; U+0301 maps to zero-advance glyph 129;
// U+1F600 maps to glyph 130 through a format 12 subtable.
internal static class TestFontBuilder
{
    public const int GlyphCount = 131;
    public const int UnitsPerEm = 1000;

    public static int[] AllCodePoints()
    {
        var codePoints = new List<int>();
        for (int codePoint = 0x20; codePoint <= 0x9F; codePoint++)
        {
            codePoints.Add(codePoint);
        }

        codePoints.Add(0x0301);
        codePoints.Add(0x1F600);
        return codePoints.ToArray();
    }

    public static OpenTypeFont LoadTestFont()
    {
        return OpenTypeFont.Load(CreateTestFont());
    }

    public static byte[] CreateTestFont()
    {
        return Assemble(scalerTag: new string(new[] { (char)0, (char)1, (char)0, (char)0 }), includeOutlines: true);
    }

    public static byte[] CreateCffKindFont(string familyName = "TestFont")

    {
        return Assemble(scalerTag: "OTTO", includeOutlines: false, familyName: familyName);
    }

    public static (int Offset, int Length) GetTableRange(byte[] font, string tag)
    {
        int numTables = (font[4] << 8) | font[5];
        for (int i = 0; i < numTables; i++)
        {
            int record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) == tag)
            {
                int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
                int length = (font[record + 12] << 24) | (font[record + 13] << 16) | (font[record + 14] << 8) | font[record + 15];
                return (offset, length);
            }
        }

        throw new InvalidOperationException("Test font is missing table " + tag + ".");
    }

    private static byte[] Assemble(string scalerTag, bool includeOutlines, string familyName = "TestFont")
    {
        byte[] head = BuildHead();
        byte[] hhea = BuildHhea();
        byte[] hmtx = BuildHmtx();
        byte[] maxp = BuildMaxp();
        byte[] name = BuildName(familyName);
        byte[] os2 = BuildOs2();
        byte[] post = BuildPost();
        byte[] gpos = BuildGpos();
        byte[] cmap = BuildCmap();
        var tables = new List<(string Tag, byte[] Data)>
        {
            ("OS/2", os2),
            ("cmap", cmap),
            ("head", head),
            ("hhea", hhea),
            ("hmtx", hmtx),
            ("maxp", maxp),
            ("name", name),
            ("post", post),
            ("GPOS", gpos),
        };

        if (includeOutlines)
        {
            byte[] glyf = BuildGlyf(out byte[] loca);
            tables.Add(("glyf", glyf));
            tables.Add(("loca", loca));
        }
        else
        {
            tables.Add(("CFF ", new byte[] { 1, 0, 4, 0, 0, 0, 0, 0 }));
        }

        tables.Sort((left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        int directoryLength = 12 + tables.Count * 16;
        int dataOffset = directoryLength;
        var directory = new List<byte>();
        var blob = new List<byte>();
        foreach ((string tag, byte[] data) in tables)
        {
            dataOffset = (dataOffset + 3) & ~3;
            foreach (char c in tag)
            {
                directory.Add((byte)c);
            }

            directory.AddRange(new byte[] { 0, 0, 0, 0 });
            directory.Add((byte)(dataOffset >> 24));
            directory.Add((byte)(dataOffset >> 16));
            directory.Add((byte)(dataOffset >> 8));
            directory.Add((byte)dataOffset);
            directory.Add((byte)(data.Length >> 24));
            directory.Add((byte)(data.Length >> 16));
            directory.Add((byte)(data.Length >> 8));
            directory.Add((byte)data.Length);
            while (blob.Count < dataOffset - directoryLength)
            {
                blob.Add(0);
            }

            blob.AddRange(data);
            dataOffset += data.Length;
        }

        var output = new List<byte>();
        foreach (char c in scalerTag)
        {
            output.Add((byte)c);
        }

        int tableCount = tables.Count;
        int maxPower = 1;
        int entrySelector = 0;
        while (maxPower * 2 <= tableCount)
        {
            maxPower *= 2;
            entrySelector++;
        }

        int searchRange = maxPower * 16;
        int rangeShift = tableCount * 16 - searchRange;
        output.Add((byte)(tableCount >> 8));
        output.Add((byte)tableCount);
        output.Add((byte)(searchRange >> 8));
        output.Add((byte)searchRange);
        output.Add((byte)(entrySelector >> 8));
        output.Add((byte)entrySelector);
        output.Add((byte)(rangeShift >> 8));
        output.Add((byte)rangeShift);
        output.AddRange(directory);
        while (output.Count < directoryLength)
        {
            output.Add(0);
        }

        output.AddRange(blob);
        return output.ToArray();
    }

    private static byte[] BuildHead()
    {
        byte[] head = new byte[54];
        head[0] = 0;
        head[1] = 1;
        head[2] = 0;
        head[3] = 0;
        head[8] = 0x5F;
        head[9] = 0x0F;
        head[10] = 0x3C;
        head[11] = 0xF5;
        head[18] = (byte)(UnitsPerEm >> 8);
        head[19] = (byte)(UnitsPerEm & 0xFF);
        head[40] = 0x03;
        head[41] = 0xE8;
        head[42] = 0x03;
        head[43] = 0xE8;
        head[50] = 0;
        head[51] = 1;
        return head;
    }

    private static byte[] BuildHhea()
    {
        byte[] hhea = new byte[36];
        hhea[0] = 0;
        hhea[1] = 1;
        hhea[4] = 0x03;
        hhea[5] = 0x20;
        hhea[6] = 0xFF;
        hhea[7] = 0x38;
        hhea[34] = (byte)(GlyphCount >> 8);
        hhea[35] = (byte)GlyphCount;
        return hhea;
    }

    private static ushort AdvanceForGlyph(int glyph)
    {
        if (glyph == 129)
        {
            return 0;
        }

        if (glyph == 34)
        {
            return 600;
        }

        if (glyph == 35)
        {
            return 620;
        }

        return 500;
    }

    private static byte[] BuildHmtx()
    {
        byte[] hmtx = new byte[GlyphCount * 4];
        for (int glyph = 0; glyph < GlyphCount; glyph++)
        {
            ushort advance = AdvanceForGlyph(glyph);
            hmtx[glyph * 4] = (byte)(advance >> 8);
            hmtx[glyph * 4 + 1] = (byte)advance;
        }

        return hmtx;
    }

    private static byte[] BuildMaxp()
    {
        return new byte[] { 0, 1, 0, 0, (byte)(GlyphCount >> 8), (byte)GlyphCount };
    }

    private static byte[] BuildName(string familyName)
    {
        byte[] storage = System.Text.Encoding.BigEndianUnicode.GetBytes(familyName);
        byte[] name = new byte[6 + 12 + storage.Length];
        name[2] = 0;
        name[3] = 1;
        name[4] = 0;
        name[5] = 18;
        name[6] = 0;
        name[7] = 3;
        name[8] = 0;
        name[9] = 1;
        name[10] = 0x04;
        name[11] = 0x09;
        name[12] = 0;
        name[13] = 1;
        name[14] = (byte)(storage.Length >> 8);
        name[15] = (byte)storage.Length;
        name[16] = 0;
        name[17] = 0;
        storage.CopyTo(name, 18);
        return name;
    }

    private static byte[] BuildOs2()
    {
        byte[] os2 = new byte[78];
        os2[0] = 0;
        os2[1] = 4;
        os2[4] = 0x01;
        os2[5] = 0x90;
        os2[6] = 5;
        os2[26] = 50;
        os2[28] = 0x01;
        os2[29] = 0x2C;
        os2[68] = 0x03;
        os2[69] = 0x20;
        os2[70] = 0xFF;
        os2[71] = 0x38;
        os2[74] = 0x03;
        os2[75] = 0x84;
        os2[76] = 0;
        os2[77] = 200;
        return os2;
    }

    private static byte[] BuildPost()
    {
        byte[] post = new byte[32];
        post[0] = 0;
        post[1] = 3;
        post[8] = 0xFF;
        post[9] = 0x9C;
        post[10] = 50;
        return post;
    }

    private static byte[] BuildCmap()
    {
        var bmp = new List<int>();
        for (int codePoint = 0x20; codePoint <= 0x9F; codePoint++)
        {
            bmp.Add(codePoint);
        }

        bmp.Add(0x0301);
        int[] sorted = bmp.OrderBy(codePoint => codePoint).ToArray();
        int segmentCount = sorted.Length + 1;
        int format4Length = 16 + segmentCount * 8;
        byte[] format4 = new byte[format4Length];
        format4[0] = 0;
        format4[1] = 4;
        format4[2] = (byte)(format4Length >> 8);
        format4[3] = (byte)format4Length;
        format4[6] = (byte)((segmentCount * 2) >> 8);
        format4[7] = (byte)(segmentCount * 2);
        int maxPower = 1;
        int entrySelector = 0;
        while (maxPower * 2 <= segmentCount)
        {
            maxPower *= 2;
            entrySelector++;
        }

        int searchRange = maxPower * 2;
        format4[8] = (byte)(searchRange >> 8);
        format4[9] = (byte)searchRange;
        format4[10] = (byte)(entrySelector >> 8);
        format4[11] = (byte)entrySelector;
        int rangeShift = segmentCount * 2 - searchRange;
        format4[12] = (byte)(rangeShift >> 8);
        format4[13] = (byte)rangeShift;
        int endCodes = 14;
        int startCodes = endCodes + segmentCount * 2 + 2;
        int idDeltas = startCodes + segmentCount * 2;
        int idRangeOffsets = idDeltas + segmentCount * 2;
        for (int i = 0; i < sorted.Length; i++)
        {
            int glyph = GlyphForCodePoint(sorted[i]);
            format4[endCodes + i * 2] = (byte)(sorted[i] >> 8);
            format4[endCodes + i * 2 + 1] = (byte)sorted[i];
            format4[startCodes + i * 2] = (byte)(sorted[i] >> 8);
            format4[startCodes + i * 2 + 1] = (byte)sorted[i];
            int delta = glyph - sorted[i];
            format4[idDeltas + i * 2] = (byte)(delta >> 8);
            format4[idDeltas + i * 2 + 1] = (byte)delta;
        }

        int sentinel = segmentCount - 1;
        format4[endCodes + sentinel * 2] = 0xFF;
        format4[endCodes + sentinel * 2 + 1] = 0xFF;
        format4[startCodes + sentinel * 2] = 0xFF;
        format4[startCodes + sentinel * 2 + 1] = 0xFF;
        format4[idDeltas + sentinel * 2] = 0;
        format4[idDeltas + sentinel * 2 + 1] = 1;
        byte[] format12 = BuildCmapFormat12();
        int format4Offset = 4 + 8 * 2;
        int format12Offset = format4Offset + format4.Length;
        byte[] output = new byte[format12Offset + format12.Length];
        output[2] = 0;
        output[3] = 2;
        output[4] = 0;
        output[5] = 3;
        output[6] = 0;
        output[7] = 1;
        output[8] = (byte)(format4Offset >> 24);
        output[9] = (byte)(format4Offset >> 16);
        output[10] = (byte)(format4Offset >> 8);
        output[11] = (byte)format4Offset;
        output[12] = 0;
        output[13] = 3;
        output[14] = 0;
        output[15] = 10;
        output[16] = (byte)(format12Offset >> 24);
        output[17] = (byte)(format12Offset >> 16);
        output[18] = (byte)(format12Offset >> 8);
        output[19] = (byte)format12Offset;
        Array.Copy(format4, 0, output, format4Offset, format4.Length);
        Array.Copy(format12, 0, output, format12Offset, format12.Length);
        return output;
    }

    private static int GlyphForCodePoint(int codePoint)
    {
        if (codePoint == 0x0301)
        {
            return 129;
        }

        if (codePoint == 0x1F600)
        {
            return 130;
        }

        return codePoint - 0x20 + 1;
    }

    private static byte[] BuildCmapFormat12()
    {
        int[] starts = { 0x20, 0x0301, 0x1F600 };
        int[] ends = { 0x9F, 0x0301, 0x1F600 };
        byte[] output = new byte[16 + starts.Length * 12];
        output[0] = 0;
        output[1] = 12;
        output[2] = 0;
        output[3] = 0;
        output[9] = (byte)(output.Length >> 16);
        output[10] = (byte)(output.Length >> 8);
        output[11] = (byte)output.Length;
        output[12] = 0;
        output[13] = 0;
        output[14] = 0;
        output[15] = (byte)starts.Length;
        for (int i = 0; i < starts.Length; i++)
        {
            int glyph = GlyphForCodePoint(starts[i]);
            WriteUInt32(output, 16 + i * 12, (uint)starts[i]);
            WriteUInt32(output, 16 + i * 12 + 4, (uint)ends[i]);
            WriteUInt32(output, 16 + i * 12 + 8, (uint)glyph);
        }

        return output;
    }

    public static int GetRecordOffset(byte[] font, string tag)
    {
        int numTables = (font[4] << 8) | font[5];
        for (int i = 0; i < numTables; i++)
        {
            int record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) == tag)
            {
                return record;
            }
        }

        throw new InvalidOperationException("Test font is missing table " + tag + ".");
    }

    public static List<(string Tag, int Offset, int Length)> GetDirectoryEntries(byte[] font)
    {
        int numTables = (font[4] << 8) | font[5];
        var entries = new List<(string Tag, int Offset, int Length)>();
        for (int i = 0; i < numTables; i++)
        {
            int record = 12 + i * 16;
            string tag = System.Text.Encoding.ASCII.GetString(font, record, 4);
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            int length = (font[record + 12] << 24) | (font[record + 13] << 16) | (font[record + 14] << 8) | font[record + 15];
            entries.Add((tag, offset, length));
        }

        return entries;
    }

    private static void WriteUInt32To(List<byte> output, uint value)
    {
        output.Add((byte)(value >> 24));
        output.Add((byte)(value >> 16));
        output.Add((byte)(value >> 8));
        output.Add((byte)value);
    }

    private static void WriteUInt32(byte[] output, int offset, uint value)
    {
        output[offset] = (byte)(value >> 24);
        output[offset + 1] = (byte)(value >> 16);
        output[offset + 2] = (byte)(value >> 8);
        output[offset + 3] = (byte)value;
    }

    private static byte[] BuildGlyf(out byte[] loca)
    {
        byte[] glyf = new byte[GlyphCount * 10];
        for (int glyph = 0; glyph < GlyphCount; glyph++)
        {
            glyf[glyph * 10 + 6] = 0x01;
            glyf[glyph * 10 + 7] = 0xF4;
            glyf[glyph * 10 + 8] = 0x02;
            glyf[glyph * 10 + 9] = 0xBC;
        }

        loca = new byte[(GlyphCount + 1) * 4];
        for (int glyph = 0; glyph <= GlyphCount; glyph++)
        {
            WriteUInt32(loca, glyph * 4, (uint)(glyph * 10));
        }

        return glyf;
    }

    // Rebuilds a font with one table swapped, keeping header and directory order.
    // Used for hostile-structure probes (oversized GPOS class kerning, compound
    // nesting) that Assemble cannot express.
    public static byte[] ReplaceTable(byte[] font, string tag, byte[] replacement)
    {
        List<(string Tag, int Offset, int Length)> entries = GetDirectoryEntries(font);
        var tables = new List<(string Tag, byte[] Data)>();
        foreach ((string entryTag, int offset, int length) in entries)
        {
            if (entryTag == tag)
            {
                tables.Add((tag, replacement));
                continue;
            }

            var data = new byte[length];
            Array.Copy(font, offset, data, 0, length);
            tables.Add((entryTag, data));
        }

        if (!tables.Any(table => table.Tag == tag))
        {
            throw new InvalidOperationException("Test font is missing table " + tag + ".");
        }

        int directoryLength = 12 + tables.Count * 16;
        int dataOffset = directoryLength;
        var directory = new List<byte>();
        var blob = new List<byte>();
        foreach ((string entryTag, byte[] data) in tables)
        {
            dataOffset = (dataOffset + 3) & ~3;
            foreach (char c in entryTag)
            {
                directory.Add((byte)c);
            }

            directory.AddRange(new byte[] { 0, 0, 0, 0 });
            directory.Add((byte)(dataOffset >> 24));
            directory.Add((byte)(dataOffset >> 16));
            directory.Add((byte)(dataOffset >> 8));
            directory.Add((byte)dataOffset);
            directory.Add((byte)(data.Length >> 24));
            directory.Add((byte)(data.Length >> 16));
            directory.Add((byte)(data.Length >> 8));
            directory.Add((byte)data.Length);
            while (blob.Count < dataOffset - directoryLength)
            {
                blob.Add(0);
            }

            blob.AddRange(data);
            dataOffset += data.Length;
        }

        var output = new List<byte>();
        // Table count is unchanged, so the sfnt header (searchRange and friends)
        // stays byte-identical.
        for (int i = 0; i < 12; i++)
        {
            output.Add(font[i]);
        }

        output.AddRange(directory);
        while (output.Count < directoryLength)
        {
            output.Add(0);
        }

        output.AddRange(blob);
        return output.ToArray();
    }

    public static byte[] CreateCollection(string firstFamily, string secondFamily)
    {
        byte[] first = Assemble(scalerTag: new string(new[] { (char)0, (char)1, (char)0, (char)0 }), includeOutlines: true, familyName: firstFamily);
        byte[] secondName = BuildName(secondFamily);
        var entries = GetDirectoryEntries(first);
        int tableCount = entries.Count;
        int directoryLength = 12 + tableCount * 16;
        int headerLength = 12 + 2 * 4;
        int firstDirectoryOffset = headerLength;
        int secondDirectoryOffset = headerLength + directoryLength;
        int dataOffset = headerLength + directoryLength * 2;
        int firstDataOffset = -1;
        foreach ((string tag, int offset, int length) in entries)
        {
            if (firstDataOffset < 0 || offset < firstDataOffset)
            {
                firstDataOffset = offset;
            }
        }

        int secondNameOffset = dataOffset + (first.Length - firstDataOffset) + ((4 - ((first.Length - firstDataOffset) % 4)) % 4);
        var output = new List<byte>();
        output.AddRange(new byte[] { 116, 116, 99, 102, 0, 1, 0, 0, 0, 0, 0, 2 });
        WriteUInt32To(output, (uint)firstDirectoryOffset);
        WriteUInt32To(output, (uint)secondDirectoryOffset);
        foreach (bool second in new[] { false, true })
        {
            output.AddRange(new byte[] { 0, 1, 0, 0 });
            output.AddRange(new byte[] { 0, (byte)tableCount });
            output.AddRange(new byte[] { 0, 0, 0, 0, 0, 0 });
            foreach ((string tag, int offset, int length) in entries)
            {
                foreach (char c in tag)
                {
                    output.Add((byte)c);
                }

                output.AddRange(new byte[] { 0, 0, 0, 0 });
                int rebased = tag == "name" && second ? secondNameOffset : dataOffset + (offset - firstDataOffset);
                int finalLength = tag == "name" && second ? secondName.Length : length;
                WriteUInt32To(output, (uint)rebased);
                WriteUInt32To(output, (uint)finalLength);
            }
        }

        while (output.Count < dataOffset)
        {
            output.Add(0);
        }

        for (int i = firstDataOffset; i < first.Length; i++)
        {
            output.Add(first[i]);
        }

        while (output.Count < secondNameOffset)
        {
            output.Add(0);
        }

        output.AddRange(secondName);
        return output.ToArray();
    }

    private static byte[] BuildGpos()
    {
        // Minimal GPOS assembled with patched offsets (all offsets are relative to
        // their containing subtable, per the OpenType Layout specification).
        // Scripts: latn links kern feature 0; arab links kern feature 2 but arab is
        // outside the reader kern-script set, so feature 2 stays inactive.
        // Features: kern 0 -> lookups 0,3; kern 1 -> lookup 1 (unlinked from every
        // accepted script); kern 2 -> lookup 2 (arab-only); liga 3 -> lookup 4.
        // Lookups: 0 pair 34->35 (-50); 1 pair 36->37 (-70); 2 pair 40->41 (-90);
        // 3 extension wrapping pair 38->39 (-30); 4 pair 42->43 (-110).
        var output = new List<byte>();
        void WriteU16(int value)
        {
            output.Add((byte)(value >> 8));
            output.Add((byte)value);
        }
        void WriteI16(int value)
        {
            output.Add((byte)(value >> 8));
            output.Add((byte)value);
        }
        void WriteTag(string tag)
        {
            foreach (char c in tag)
            {
                output.Add((byte)c);
            }
        }
        int ReserveU16()
        {
            int position = output.Count;
            output.Add(0);
            output.Add(0);
            return position;
        }
        void PatchU16(int position, int value)
        {
            output[position] = (byte)(value >> 8);
            output[position + 1] = (byte)value;
        }
        int ReserveU32()
        {
            int position = output.Count;
            output.Add(0);
            output.Add(0);
            output.Add(0);
            output.Add(0);
            return position;
        }
        void PatchU32(int position, int value)
        {
            output[position] = (byte)(value >> 24);
            output[position + 1] = (byte)(value >> 16);
            output[position + 2] = (byte)(value >> 8);
            output[position + 3] = (byte)value;
        }
        int WritePairSet(int leftGlyph, int rightGlyph, int xAdvance)
        {
            int subtable = output.Count;
            WriteU16(1);
            int coverageOffset = ReserveU16();
            WriteU16(4);
            WriteU16(0);
            WriteU16(1);
            int pairSetOffset = ReserveU16();
            int coverage = output.Count;
            PatchU16(coverageOffset, coverage - subtable);
            WriteU16(1);
            WriteU16(1);
            WriteU16(leftGlyph);
            int pairSet = output.Count;
            PatchU16(pairSetOffset, pairSet - subtable);
            WriteU16(1);
            WriteU16(rightGlyph);
            WriteI16(xAdvance);
            return subtable;
        }

        WriteU16(1);
        WriteU16(0);
        int scriptListOffset = ReserveU16();
        int featureListOffset = ReserveU16();
        int lookupListOffset = ReserveU16();

        int scriptList = output.Count;
        PatchU16(scriptListOffset, scriptList);
        WriteU16(2);
        WriteTag("latn");
        int latnScriptOffset = ReserveU16();
        WriteTag("arab");
        int arabScriptOffset = ReserveU16();

        int latnScript = output.Count;
        PatchU16(latnScriptOffset, latnScript - scriptList);
        int latnDefaultOffset = ReserveU16();
        WriteU16(0);
        int latnLangSys = output.Count;
        PatchU16(latnDefaultOffset, latnLangSys - latnScript);
        WriteU16(0);
        WriteU16(0xFFFF);
        WriteU16(1);
        WriteU16(0);

        int arabScript = output.Count;
        PatchU16(arabScriptOffset, arabScript - scriptList);
        int arabDefaultOffset = ReserveU16();
        WriteU16(0);
        int arabLangSys = output.Count;
        PatchU16(arabDefaultOffset, arabLangSys - arabScript);
        WriteU16(0);
        WriteU16(0xFFFF);
        WriteU16(1);
        WriteU16(2);

        string[] featureTags = ["kern", "kern", "kern", "liga"];
        int[][] featureLookups = [[0, 3], [1], [2], [4]];
        int featureList = output.Count;
        PatchU16(featureListOffset, featureList);
        WriteU16(featureTags.Length);
        var featureOffsetSlots = new List<int>();
        for (int i = 0; i < featureTags.Length; i++)
        {
            WriteTag(featureTags[i]);
            featureOffsetSlots.Add(ReserveU16());
        }
        for (int i = 0; i < featureTags.Length; i++)
        {
            PatchU16(featureOffsetSlots[i], output.Count - featureList);
            WriteU16(0);
            WriteU16(featureLookups[i].Length);
            foreach (int lookup in featureLookups[i])
            {
                WriteU16(lookup);
            }
        }

        int lookupList = output.Count;
        PatchU16(lookupListOffset, lookupList);
        WriteU16(5);
        var lookupOffsetSlots = new List<int>();
        for (int i = 0; i < 5; i++)
        {
            lookupOffsetSlots.Add(ReserveU16());
        }

        int[] lookupStarts = new int[5];
        int[] lookupSubtableSlots = new int[5];
        ushort[] lookupTypes = [2, 2, 2, 9, 2];
        for (int i = 0; i < 5; i++)
        {
            lookupStarts[i] = output.Count;
            PatchU16(lookupOffsetSlots[i], lookupStarts[i] - lookupList);
            WriteU16(lookupTypes[i]);
            WriteU16(0);
            WriteU16(1);
            lookupSubtableSlots[i] = ReserveU16();
        }
        int pair0 = WritePairSet(34, 35, -50);
        PatchU16(lookupSubtableSlots[0], pair0 - lookupStarts[0]);
        int pair1 = WritePairSet(36, 37, -70);
        PatchU16(lookupSubtableSlots[1], pair1 - lookupStarts[1]);
        int pair2 = WritePairSet(40, 41, -90);
        PatchU16(lookupSubtableSlots[2], pair2 - lookupStarts[2]);
        int extension = output.Count;
        PatchU16(lookupSubtableSlots[3], extension - lookupStarts[3]);
        WriteU16(1);
        WriteU16(2);
        int extensionOffset = ReserveU32();
        int inner = WritePairSet(38, 39, -30);
        PatchU32(extensionOffset, inner - extension);
        int pair4 = WritePairSet(42, 43, -110);
        PatchU16(lookupSubtableSlots[4], pair4 - lookupStarts[4]);
        return output.ToArray();
    }

}
