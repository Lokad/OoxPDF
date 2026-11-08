using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetFontSubsetAdmitsMetadataBeforeBuildingAndRetainsMappings()
    {
        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(Case("PortableFonts/DejaVuSans.ttf")));
        var wanted = new Dictionary<ushort, int>();
        for (int cp = 32; cp < 65536; cp++)
        {
            ushort glyph = font.MapCodePoint(cp);
            if (glyph != 0) wanted.TryAdd(glyph, cp);
        }
        TestAssert.True(wanted.Count > 1000, "Exercise a broad source-supported glyph set.");
        TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out var baseline, out _),
            "The pinned face produces the ungoverned control subset.");
        var expected = TestAssert.NotNull(baseline);
        long before = 0, beforeAdmission = -1;
        var pool = new TestMemoryBudget
        {
            RefuseChange = (_, next) =>
            {
                if (next > 0 && beforeAdmission < 0)
                    beforeAdmission = GC.GetAllocatedBytesForCurrentThread() - before;
                return false;
            }
        };
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            before = GC.GetAllocatedBytesForCurrentThread();
            TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out var result, out _),
                "Admitted preparation retains the source-supported glyph set.");
            var subset = TestAssert.NotNull(result);
            TestAssert.True(beforeAdmission >= 0 && beforeAdmission < 64 * 1024,
                "Substantial glyph metadata must not be built before the first host grant.");
            TestAssert.True(pool.Reserved >= subset.Bytes.Length + 4L * subset.CidByOriginalGlyph.Count,
                "Retained admission must include the CID map's keys/values beside the program.");
            TestAssert.True(subset.Bytes.AsSpan().SequenceEqual(expected.Bytes), "Admission preserves every subset byte.");
            TestAssert.True(subset.CidByOriginalGlyph.OrderBy(p => p.Key).SequenceEqual(expected.CidByOriginalGlyph.OrderBy(p => p.Key)),
                "Admission preserves every source-to-CID association.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetFontSubsetMetadataGrowthRefusesAndRefunds()
    {
        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(Case("PortableFonts/DejaVuSans.ttf")));
        var wanted = new Dictionary<ushort, int> { [font.MapCodePoint('é')] = 'é' };
        var pool = new TestMemoryBudget { RefuseChange = (before, after) => before > 0 && after > before };
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                OpenTypeFontSubsetter.TryCreate(font, wanted, default, out _, out _));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out _, out _),
            "Refusal leaves the caller-owned font usable.");
    }

    public static void MemoryBudgetFontSubsetMetadataGrowthCancellationPreservesToken()
    {
        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(Case("PortableFonts/DejaVuSans.ttf")));
        var wanted = new Dictionary<ushort, int> { [font.MapCodePoint('é')] = 'é' };
        using var cancellation = new CancellationTokenSource();
        var pool = new TestMemoryBudget
        {
            RefuseChange = (before, after) =>
            {
                if (before > 0 && after > before) cancellation.Cancel();
                return false;
            }
        };
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() =>
                OpenTypeFontSubsetter.TryCreate(font, wanted, cancellation.Token, out _, out _));
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.True(pool.Peak > 0, "Cancel after the admitted metadata family grows.");
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out _, out _),
            "Cancellation leaves the original font usable for another subset.");
    }

    public static void MemoryBudgetFontSubsetNoUsableGlyphsRetainsUnchargedFallback()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        var pool = new TestMemoryBudget(0);
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            foreach (var wanted in new[] { new Dictionary<ushort, int>(), new Dictionary<ushort, int> { [0] = 65, [ushort.MaxValue] = 66 } })
                TestAssert.True(!OpenTypeFontSubsetter.TryCreate(font, wanted, default, out _, out _),
                    "Empty or unusable glyph requests keep the no-subset fallback.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0L, pool.Peak);
        }
    }

    public static void MemoryBudgetFontSubsetEmptyMappingKeepsCompoundNotdefClosure()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        int glyphs = OpenTypeFont.Load(bytes).GlyphCount;
        byte[] glyf = new byte[16 + (glyphs - 1) * 10];
        BinaryPrimitives.WriteInt16BigEndian(glyf.AsSpan(0, 2), -1);
        BinaryPrimitives.WriteUInt16BigEndian(glyf.AsSpan(10, 2), 2); // Byte-sized XY arguments.
        BinaryPrimitives.WriteUInt16BigEndian(glyf.AsSpan(12, 2), 1); // .notdef references glyph1.
        byte[] loca = new byte[(glyphs + 1) * 4];
        for (int i = 1; i <= glyphs; i++)
            BinaryPrimitives.WriteUInt32BigEndian(loca.AsSpan(i * 4, 4), (uint)(16 + (i - 1) * 10));
        bytes = TestFontBuilder.ReplaceTable(TestFontBuilder.ReplaceTable(bytes, "glyf", glyf), "loca", loca);
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        var wanted = new Dictionary<ushort, int>();
        TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out var baseline, out _),
            "A compound .notdef retains its existing component closure without mapped characters.");
        var expected = TestAssert.NotNull(baseline);
        TestAssert.Equal(2, expected.CidByOriginalGlyph.Count);
        TestAssert.Equal((ushort)0, expected.CidByOriginalGlyph[0]);
        TestAssert.Equal((ushort)1, expected.CidByOriginalGlyph[1]);
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            TestAssert.True(OpenTypeFontSubsetter.TryCreate(font, wanted, default, out var result, out _),
                "Governed preparation preserves compound .notdef behavior.");
            var subset = TestAssert.NotNull(result);
            TestAssert.True(subset.Bytes.AsSpan().SequenceEqual(expected.Bytes), "The compound control keeps every subset byte.");
            TestAssert.True(pool.Reserved > subset.Bytes.Length, "Its retained component mapping remains funded.");
        }
        TestAssert.Equal(0L, pool.Reserved);
    }
}
