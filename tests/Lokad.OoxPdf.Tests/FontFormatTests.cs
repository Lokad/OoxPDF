using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

// Font-format validation (P04) plus platform-independent core checks (Q01):
// every test here runs on any host, with or without installed fonts.
internal static class FontFormatTests
{
    public static void SyntheticFontLoadsWithExpectedIdentity()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.Equal("TestFont", font.FamilyName);
        TestAssert.Equal(TestFontBuilder.GlyphCount, (int)font.GlyphCount);
        TestAssert.Equal(TestFontBuilder.UnitsPerEm, (int)font.UnitsPerEm);
        TestAssert.True(font.HasTrueTypeOutlines, "Synthetic font must report TrueType outlines.");
        TestAssert.True(!font.HasCffOutlines, "Synthetic font must not report CFF outlines.");
        TestAssert.Equal(34, (int)font.MapCodePoint(65));
    }

    public static void SyntheticZeroAdvanceGlyphIsPreserved()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        ushort combining = font.MapCodePoint(0x0301);

        TestAssert.True(combining != 0, "Synthetic font must map U+0301.");
        TestAssert.Equal(0, (int)font.GetAdvanceWidth(combining));

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, TestFontBuilder.AllCodePoints(), CancellationToken.None);
        TestAssert.True(embedded.TryGetEncodedCid(combining, out ushort cid), "Combining mark must have an encoded CID.");
        TestAssert.True(
            embedded.BuildWidthArray(CancellationToken.None).Contains(cid.ToString() + " [0]", StringComparison.Ordinal),
            "Zero-advance glyphs must emit an explicit zero width.");
    }

    public static void SyntheticSubsetRoundTripKeepsCmapAndWidths()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, TestFontBuilder.AllCodePoints(), CancellationToken.None);

        TestAssert.True(embedded.UsesSubsetFontProgram, "Synthetic font must subset.");
        TestAssert.True(embedded.UnicodeByCid.Count > 100, "Synthetic probe must exceed one CMap block.");
        string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
        foreach (string line in cmap.Split((char)10))
        {
            string trimmed = line.Trim();
            if (trimmed.EndsWith("beginbfchar", StringComparison.Ordinal))
            {
                int count = int.Parse(trimmed.Substring(0, trimmed.Length - "beginbfchar".Length).Trim(), System.Globalization.CultureInfo.InvariantCulture);
                TestAssert.True(count >= 1 && count <= 100, "Each bfchar block must hold 1..100 mappings.");
            }
        }

        OpenTypeFont subset = OpenTypeFont.Load(embedded.FontProgramBytes.ToArray());
        TestAssert.True(subset.MapCodePoint(65) != 0, "Subset must keep basic mappings.");
        TestAssert.True(subset.MapCodePoint(0x1F600) != 0, "Subset must keep supplementary mappings.");
        TestAssert.Equal("TestFont", subset.FamilyName);
    }

    public static void SyntheticSubsetTagIsDeterministicAndDistinct()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, new[] { 65, 66 }, CancellationToken.None);
        PdfEmbeddedFont repeat = PdfEmbeddedFont.Create(font, new[] { 65, 66 }, CancellationToken.None);
        PdfEmbeddedFont other = PdfEmbeddedFont.Create(font, new[] { 67, 68 }, CancellationToken.None);

        TestAssert.Equal(first.BaseFontName, repeat.BaseFontName);
        TestAssert.True(first.BaseFontName != other.BaseFontName, "Differing subsets must have distinct names.");
        TestAssert.True(first.BaseFontName.Length > 7 && first.BaseFontName[6] == (char)43, "Subset name must carry a six-letter tag plus separator.");
    }

    public static void CffKindFontRefusesCidFontType2Embedding()
    {
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateCffKindFont());

        TestAssert.True(!font.HasTrueTypeOutlines, "CFF-kind fixture must lack TrueType outlines.");
        TestAssert.True(font.HasCffOutlines, "CFF-kind fixture must report CFF outlines.");
        TestAssert.True(
            !OpenTypeFontSubsetter.TryCreate(font, new[] { 65 }.ToDictionary(codePoint => font.MapCodePoint(codePoint), codePoint => codePoint), CancellationToken.None, out _, out string? reason) && reason is not null,
            "Subsetter must explain CFF refusal.");
        InvalidDataException cffFailure = TestAssert.Throws<InvalidDataException>(() => PdfEmbeddedFont.Create(font, new[] { 65 }, CancellationToken.None));
        TestAssert.Contains("TestFont", cffFailure.Message);
    }

    public static void RejectsTruncatedTableDirectory()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        byte[] truncated = new byte[20];
        Array.Copy(font, truncated, 20);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(truncated));
    }

    public static void RejectsWrappingTableOffsets()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        int record = TestFontBuilder.GetRecordOffset(font, "head");
        WriteBigEndianUInt32(font, record + 8, 0xFFFFFFF0u);
        WriteBigEndianUInt32(font, record + 12, 0x20u);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    public static void RejectsDuplicateTableTags()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        Array.Copy(font, 12, font, 28, 16);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    public static void RejectsShortHeadTable()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        int headRecord = TestFontBuilder.GetRecordOffset(font, "head");
        (int _, int headLength) = TestFontBuilder.GetTableRange(font, "head");
        WriteBigEndianUInt32(font, headRecord + 12, (uint)(headLength - 44));
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    public static void RejectsZeroUnitsPerEm()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        (int offset, _) = TestFontBuilder.GetTableRange(font, "head");
        font[offset + 18] = 0;
        font[offset + 19] = 0;
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    public static void RejectsUnrecognizedScaler()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        font[0] = (byte)119;
        font[1] = (byte)79;
        font[2] = (byte)70;
        font[3] = (byte)70;
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    private static void WriteBigEndianUInt32(byte[] font, int offset, uint value)
    {
        font[offset] = (byte)(value >> 24);
        font[offset + 1] = (byte)(value >> 16);
        font[offset + 2] = (byte)(value >> 8);
        font[offset + 3] = (byte)value;
    }

    public static void CollectionFacesKeepDistinctIdentities()
    {
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");

        TestAssert.Equal(2, OpenTypeFont.GetCollectionFontCount(collection));
        TestAssert.Equal("TestFontA", OpenTypeFont.Load(collection, 0).FamilyName);
        TestAssert.Equal("TestFontB", OpenTypeFont.Load(collection, 1).FamilyName);
    }

    public static void FontProgramLoaderSharedProgramsRetainIndependentOwnership()
    {
        foreach (byte[] program in new[] { TestFontBuilder.CreateTestFont(), TestFontBuilder.CreateCollection("TestFontA", "TestFontB") })
        {
            var backing = new byte[program.Length + 73];
            program.CopyTo(backing, 37);
            var source = new MemoryFontProgramSource("ownership:shared-slice", backing.AsMemory(37, program.Length));
            int faceCount = OpenTypeFont.GetCollectionFontCount(program);
            OpenTypeFont[] expected = Enumerable.Range(0, faceCount).Select(i => OpenTypeFont.Load(program.ToArray(), i)).ToArray();
            var fonts = new OpenTypeFont[8];
            Parallel.For(0, fonts.Length, i =>
            {
                var resolution = new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, i % faceCount, false), source, IsFallback: false);
                fonts[i] = TestAssert.NotNull(FontProgramLoader.Load(resolution, CancellationToken.None));
            });
            TestAssert.True(program.AsSpan().SequenceEqual(backing.AsSpan(37, program.Length)), "Concurrent parsing must leave shared source bytes unchanged.");
            for (int i = 0; i < fonts.Length; i++)
            {
                OpenTypeFont font = fonts[i];
                OpenTypeFont reference = expected[i % faceCount];
                TestAssert.True(font.Bytes.Span.SequenceEqual(reference.Bytes.Span), "Selected SFNT must retain byte-exact extraction.");
                TestAssert.Equal(reference.FamilyName, font.FamilyName);
                TestAssert.Equal(reference.MapCodePoint('A'), font.MapCodePoint('A'));
                TestAssert.Equal(reference.GetKerning(34, 35), font.GetKerning(34, 35));
                TestAssert.True(System.Runtime.InteropServices.MemoryMarshal.TryGetArray(font.Bytes, out ArraySegment<byte> owned), "Parsed font must own an array.");
                TestAssert.True(!ReferenceEquals(backing, owned.Array), "Parsed font must not retain the shared source array.");
                for (int j = 0; j < i; j++)
                {
                    TestAssert.True(System.Runtime.InteropServices.MemoryMarshal.TryGetArray(fonts[j].Bytes, out ArraySegment<byte> other) && !ReferenceEquals(owned.Array, other.Array), "Independent loads must own independent SFNT buffers.");
                }
            }
        }
    }

    public static void FontProgramLoaderCollectionExtractionSkipsUnusedSourceCopy()
    {
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        var padded = new byte[4 * 1024 * 1024];
        collection.CopyTo(padded, 0);
        var resolution = new FontFaceResolution("TestFontB", "TestFontB", new FontStyleKey(false, false, 400, 1, false), new MemoryFontProgramSource("allocation:padded-collection", padded), IsFallback: false);
        TestAssert.NotNull(FontProgramLoader.Load(resolution, CancellationToken.None));
        long before = GC.GetAllocatedBytesForCurrentThread();
        OpenTypeFont font = TestAssert.NotNull(FontProgramLoader.Load(resolution, CancellationToken.None));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal("TestFontB", font.FamilyName);
        TestAssert.True(allocated < padded.Length / 2, $"Loading a small face must not allocate a copy of its4MiB collection; allocated {allocated} bytes.");
        TestAssert.True(font.Bytes.Span.SequenceEqual(OpenTypeFont.Load(collection, 1).Bytes.Span), "Unused collection bytes must not change the selected face.");
    }

    public static void FontProgramLoaderSharedCollectionRejectsInvalidExtentsAndCancellation()
    {
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        int faceOffset = (int)OpenTypeFont.U32(collection, 12);
        var invalid = new List<byte[]> { collection[..16] };
        foreach (Action<byte[]> mutate in new Action<byte[]>[]
        {
            bytes => WriteBigEndianUInt32(bytes, 8, 257),
            bytes => WriteBigEndianUInt32(bytes, 12, uint.MaxValue),
            bytes => { bytes[faceOffset + 4] = 255; bytes[faceOffset + 5] = 255; },
            bytes => Array.Copy(bytes, faceOffset + 12, bytes, faceOffset + 28, 4),
            bytes => WriteBigEndianUInt32(bytes, faceOffset + 12 + 8, uint.MaxValue)
        })
        {
            byte[] bytes = collection.ToArray();
            mutate(bytes);
            invalid.Add(bytes);
        }
        foreach (byte[] bytes in invalid)
        {
            TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(bytes));
            TestAssert.True(FontProgramLoader.Load(Resolve(bytes, 0), CancellationToken.None) is null, "Malformed shared collections must retain the font fallback.");
        }
        foreach (int index in new[] { -1, 2 })
        {
            TestAssert.True(FontProgramLoader.Load(Resolve(collection, index), CancellationToken.None) is null, "Invalid collection face index must retain the font fallback.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => FontProgramLoader.Load(Resolve(collection, 1), cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(() => OpenTypeFont.LoadShared(collection, 1, cancelled.Token));

        static FontFaceResolution Resolve(byte[] bytes, int index) => new("TestFont", "TestFont", new FontStyleKey(false, false, 400, index, false), new MemoryFontProgramSource("invalid:collection", bytes), IsFallback: false);
    }

    public static void RejectsCollectionFaceCountOverflow()
    {
        byte[] header = new byte[16];
        header[0] = (byte)116;
        header[1] = (byte)116;
        header[2] = (byte)99;
        header[3] = (byte)102;
        header[10] = 1;
        header[11] = 1;
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(header));
    }

    public static void RejectsCollectionFaceIndexOutOfRange()
    {
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(collection, 5));
    }

    public static void RejectsTruncatedCollectionFaceDirectory()
    {
        byte[] header = new byte[28];
        header[0] = (byte)116;
        header[1] = (byte)116;
        header[2] = (byte)99;
        header[3] = (byte)102;
        header[4] = 0;
        header[5] = 1;
        header[6] = 0;
        header[7] = 0;
        header[11] = 1;
        header[15] = 16;
        header[20] = 0;
        header[21] = 1;
        header[22] = 0;
        header[23] = 0;
        header[25] = 100;
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(header));
    }

    public static void RejectsCmapFormat12GroupOverflow()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        (int cmapOffset, _) = TestFontBuilder.GetTableRange(font, "cmap");
        int format12Offset = cmapOffset + 20 + 1056;
        font[format12Offset + 12] = 0xFF;
        font[format12Offset + 13] = 0xFF;
        font[format12Offset + 14] = 0xFF;
        font[format12Offset + 15] = 0xFF;
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(font));
    }

    public static void SyntheticGposPairAdjustmentIsRead()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.True(font.TableTags.Contains("GPOS"), "Synthetic font must expose a GPOS table.");
        TestAssert.Equal(-50, (int)font.GetKerning(34, 35));
        TestAssert.Equal(0, (int)font.GetKerning(35, 34));
        TestAssert.Equal(0, (int)font.GetKerning(1, 2));
    }

    public static void SyntheticUnlinkedKernFeatureIsIgnored()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.Equal(0, (int)font.GetKerning(36, 37));
        TestAssert.Equal(-50, (int)font.GetKerning(34, 35));
    }

    public static void SyntheticExtensionKernLookupIsRead()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.Equal(-30, (int)font.GetKerning(38, 39));
        TestAssert.Equal(0, (int)font.GetKerning(39, 38));
    }

    public static void SyntheticRejectedScriptKernIsIgnored()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.Equal(0, (int)font.GetKerning(40, 41));
        TestAssert.Equal(-50, (int)font.GetKerning(34, 35));
    }

    public static void SyntheticNonKernFeatureIsIgnored()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();

        TestAssert.Equal(0, (int)font.GetKerning(42, 43));
    }

    public static void SyntheticKerningAdjustsMeasuredAdvance()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "AB".Select(c => (int)c), CancellationToken.None);
        const double size = 10d;

        double plain = embedded.MeasureTextPoints("AB", size, kerningEnabled: false);
        double kerned = embedded.MeasureTextPoints("AB", size, kerningEnabled: true);
        TestAssert.Equal(plain - 50d * size / TestFontBuilder.UnitsPerEm, kerned);
    }
}
