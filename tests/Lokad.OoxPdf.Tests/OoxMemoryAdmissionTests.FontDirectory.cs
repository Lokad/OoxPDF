using System.Buffers.Binary;
using System.Text;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetFontDirectoryRefusalPrecedesMaterialization()
    {
        byte[] sfnt = DirectoryFont();
        foreach (byte[] source in new[] { sfnt, DirectoryCollection(sfnt) })
        {
            OpenTypeFont.Load(source); // Warm ordinary parsing outside the counter.
            var pool = new TestMemoryBudget(256 * 1024);
            using (OoxConversionBudget.BeginScope(null, pool))
            {
                long start = GC.GetAllocatedBytesForCurrentThread();
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OpenTypeFont.LoadShared(source, 0, default));
                long allocated = GC.GetAllocatedBytesForCurrentThread() - start;
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                TestAssert.True(allocated < source.Length + 128 * 1024,
                    $"Directory refusal must precede bulk tag/dictionary construction; allocated {allocated} for {source.Length} source bytes.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetFontDirectoryCancellationRefundsBeforeParsing()
    {
        byte[] sfnt = DirectoryFont();
        foreach (byte[] source in new[] { sfnt, DirectoryCollection(sfnt) })
        {
            bool collection = OpenTypeFont.IsTrueTypeCollectionHeader(source), reached = false;
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (before, after) => {
                if (after > before && (collection ? before == 0 : before == source.Length))
                { reached = true; cancellation.Cancel(); }
                return false;
            }};
            using (OoxConversionBudget.BeginScope(null, pool))
            {
                var failure = TestAssert.Throws<OperationCanceledException>(() => OpenTypeFont.LoadShared(source, 0, cancellation.Token));
                TestAssert.True(reached, "Cancel when the directory's existing owner acquires capacity.");
                TestAssert.Equal(cancellation.Token, failure.CancellationToken);
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void MemoryBudgetFontDirectoryRetainsTagsAndCollectionBytes()
    {
        byte[] sfnt = DirectoryFont(), collection = DirectoryCollection(sfnt);
        var baseline = OpenTypeFont.Load(sfnt);
        long retained = 0;
        foreach (byte[] source in new[] { sfnt, collection })
        {
            byte[] original = source.ToArray();
            var expected = OpenTypeFont.Load(source);
            var pool = new TestMemoryBudget();
            using (OoxConversionBudget.BeginScope(null, pool))
            {
                var font = OpenTypeFont.LoadShared(source, 0, default);
                TestAssert.Equal(4095, font.TableTags.Count);
                TestAssert.True(font.TableTags.Contains("Z000"), "Keep unknown vendor tags available to useful inspection.");
                TestAssert.True(font.Bytes.Span.SequenceEqual(expected.Bytes.Span), "Admission preserves the established standalone/selected SFNT bytes.");
                TestAssert.Equal(baseline.MapCodePoint('A'), font.MapCodePoint('A'));
                TestAssert.Equal(baseline.GetAdvanceWidth(1), font.GetAdvanceWidth(1));
                TestAssert.Equal(baseline.GetKerning(1, 2), font.GetKerning(1, 2));
                TestAssert.True(pool.Reserved > font.Bytes.Length + 128 * 1024,
                    "Significant retained directory storage stays admitted beside the private program.");
                if (retained == 0) retained = pool.Reserved;
                else TestAssert.Equal(retained, pool.Reserved); // Extraction scratch has already refunded.
                TestAssert.True(source.AsSpan().SequenceEqual(original), "Borrowed font bytes remain unchanged.");
                GC.KeepAlive(font);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void FontDirectoryDiscoveryKeepsMetadataAndMalformedRejection()
    {
        byte[] sfnt = DirectoryFont();
        var expected = OpenTypeFont.ReadDiscoveryHeaders(sfnt, 0);
        foreach (byte[] source in new[] { sfnt, DirectoryCollection(sfnt) })
        {
            bool collection = OpenTypeFont.IsTrueTypeCollectionHeader(source);
            int face = collection ? 16 : 0;
            OpenTypeFont.FontDiscoveryHeaders Discover(byte[] bytes) => collection
                ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(bytes, 0, null)
                : OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
            var actual = Discover(source);
            TestAssert.True(expected.FamilyNames.SequenceEqual(actual.FamilyNames), "In-place and standalone discovery keep the same aliases.");
            TestAssert.Equal(expected.WeightClass, actual.WeightClass);
            TestAssert.Equal(expected.ItalicAngle, actual.ItalicAngle);
            TestAssert.Equal(expected.HasMathTable, actual.HasMathTable);
            foreach (int defect in new[] { 0, 1, 2, 3 })
            {
                byte[] invalid = source.ToArray();
                if (defect == 0) invalid.AsSpan(face + 12, 4).CopyTo(invalid.AsSpan(face + 28));
                if (defect == 1) BinaryPrimitives.WriteUInt32BigEndian(invalid.AsSpan(face + 20), uint.MaxValue);
                if (defect == 2) invalid = invalid[..(face + 40)];
                if (defect == 3) "bad!"u8.CopyTo(invalid.AsSpan(face));
                TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(invalid));
                TestAssert.Throws<InvalidDataException>(() => Discover(invalid));
            }
        }
    }

    // Parser fixture: extend the existing synthetic face with empty unique vendor
    // tables. Independently checked whole-face controls remain in qualification.
    private static byte[] DirectoryFont()
    {
        byte[] seed = TestFontBuilder.CreateTestFont();
        int count = BinaryPrimitives.ReadUInt16BigEndian(seed.AsSpan(4)), oldEnd = 12 + count * 16;
        int newEnd = 12 + 4095 * 16, shift = newEnd - oldEnd;
        byte[] output = new byte[seed.Length + shift];
        seed.AsSpan(0, 12).CopyTo(output);
        BinaryPrimitives.WriteUInt16BigEndian(output.AsSpan(4), 4095);
        BinaryPrimitives.WriteUInt16BigEndian(output.AsSpan(6), 32768);
        BinaryPrimitives.WriteUInt16BigEndian(output.AsSpan(8), 11);
        BinaryPrimitives.WriteUInt16BigEndian(output.AsSpan(10), 32752);
        seed.AsSpan(oldEnd).CopyTo(output.AsSpan(newEnd));
        var records = new List<byte[]>();
        for (int i = 0; i < count; i++)
        {
            byte[] record = seed.AsSpan(12 + 16 * i, 16).ToArray();
            uint offset = BinaryPrimitives.ReadUInt32BigEndian(record.AsSpan(8));
            BinaryPrimitives.WriteUInt32BigEndian(record.AsSpan(8), offset + (uint)shift);
            records.Add(record);
        }
        for (int i = 0; records.Count < 4095; i++)
        {
            byte[] record = new byte[16];
            Encoding.ASCII.GetBytes($"Z{i:X3}").CopyTo(record, 0);
            BinaryPrimitives.WriteUInt32BigEndian(record.AsSpan(8), (uint)output.Length);
            records.Add(record);
        }
        int index = 12;
        foreach (byte[] record in records.OrderBy(r => Encoding.ASCII.GetString(r, 0, 4), StringComparer.Ordinal))
        { record.CopyTo(output, index); index += 16; }
        return output;
    }

    private static byte[] DirectoryCollection(byte[] sfnt)
    {
        byte[] output = new byte[sfnt.Length + 16];
        "ttcf"u8.CopyTo(output);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(4), 0x10000);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(8), 1);
        BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(12), 16);
        sfnt.CopyTo(output, 16);
        int count = BinaryPrimitives.ReadUInt16BigEndian(sfnt.AsSpan(4));
        for (int i = 0; i < count; i++)
        {
            int record = 16 + 12 + i * 16;
            uint offset = BinaryPrimitives.ReadUInt32BigEndian(output.AsSpan(record + 8));
            BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(record + 8), offset + 16);
        }
        return output;
    }
}
