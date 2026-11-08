using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetKerningCoverageRefusesBeforeBulkSparseConstruction()
    {
        byte[] bytes = SparseCoverageFont(4096, 128, range: true, portable: true);
        OpenTypeFont.Load(bytes); // Warm parser/refusal prerequisites outside the counter.
        long allocated = 0, start = 0;
        var pool = new TestMemoryBudget { RefuseChange = (before, after) =>
        {
            if (before != 0 || after <= 0 || after >= bytes.Length) return false;
            allocated = GC.GetAllocatedBytesForCurrentThread() - start;
            return true;
        }};
        using (OoxConversionBudget.BeginScope(null, pool))
        {
            start = GC.GetAllocatedBytesForCurrentThread();
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OpenTypeFont.LoadShared(bytes, 0, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.True(allocated > 0 && allocated < bytes.Length + 512 * 1024,
                $"Font construction must request admission before bulk coverage/callback work; allocated before refusal: {allocated}, source: {bytes.Length}.");
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryBudgetKerningSparseBatchesRefuseAndRefund()
    {
        byte[] bytes = SparseCoverageFont(3, 4096, range: false);
        var face = ProgramResolution(new MemoryFontProgramSource("sparse-batch-owner", bytes));
        var pool = new TestMemoryBudget(256 * 1024);
        using (OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(face, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(0L, pool.Reserved);
        }
        var accepted = TestAssert.NotNull(FontProgramLoader.Load(face, default));
        TestAssert.Equal((short)-25, accepted.GetKerning(1, 2));
        TestAssert.True(accepted.Bytes.Span.SequenceEqual(bytes), "Refusal leaves the caller-owned source unchanged and reusable.");
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetKerningCoverageCancellationAndZeroBatchReuse()
    {
        byte[] bytes = SparseCoverageFont(64, 128, range: true, nonzero: false);
        using var cancellation = new CancellationTokenSource();
        bool reached = false;
        var pool = new TestMemoryBudget { RefuseChange = (before, after) =>
        {
            if (before == 0 && after > 0 && after < bytes.Length) { reached = true; cancellation.Cancel(); }
            return false;
        }};
        using (OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() => OpenTypeFont.LoadShared(bytes, 0, cancellation.Token));
            TestAssert.True(reached, "Cancellation reaches the construction owner before coverage materialization.");
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.Equal(0L, pool.Reserved);
        }
        long Peak(int aliases)
        {
            byte[] source = SparseCoverageFont(64, aliases, range: true, nonzero: false);
            var budget = new TestMemoryBudget(128 * 1024);
            using (OoxConversionBudget.BeginScope(null, budget))
            {
                var font = OpenTypeFont.LoadShared(source, 0, default);
                TestAssert.Equal((short)0, font.GetKerning(1, 2));
                GC.KeepAlive(font);
            }
            TestAssert.Equal(0L, budget.Reserved);
            return budget.Peak;
        }
        TestAssert.True(Peak(128) < Peak(1) + 4096, "Zero batches reuse one construction window rather than accumulating a work charge.");
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryBudgetKerningCoverageFormatsKeepSignedPairs()
    {
        foreach (bool range in new[] { false, true })
        {
            var font = OpenTypeFont.Load(SparseCoverageFont(64, 16, range));
            TestAssert.Equal((short)-25, font.GetKerning(1, 2));
            TestAssert.Equal((short)0, font.GetKerning(2, 1));
            TestAssert.Equal((short)0, font.GetKerning(63, 2));
        }
    }

    // Correctly counted/ordered GPOS data, sharing one subtable through the
    // declared lookup references. Whole-face FontTools controls are independent
    // of this parser fixture and are retained with the qualification evidence.
    private static byte[] SparseCoverageFont(int glyphs, int aliases, bool range, bool nonzero = true, bool portable = false)
    {
        int lookupList = 44, lookup = lookupList + 4, pair = lookup + 6 + 2 * aliases;
        int coverage = pair + 10 + 2 * glyphs;
        int nonzeroSet = coverage + (range ? 10 : 4 + 2 * glyphs), emptySet = nonzeroSet + 6;
        byte[] data = new byte[emptySet + 2];
        void Put(int offset, int value) => BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(offset, 2), checked((ushort)value));
        Put(0, 1); Put(4, 10); Put(6, 30); Put(8, lookupList);
        Put(10, 1); "DFLT"u8.CopyTo(data.AsSpan(12)); Put(16, 8);
        Put(18, 4); Put(24, ushort.MaxValue); Put(26, 1); // LangSys reserved word at22 remains zero.
        Put(30, 1); "kern"u8.CopyTo(data.AsSpan(32)); Put(36, 8); Put(40, 1);
        Put(lookupList, 1); Put(lookupList + 2, lookup - lookupList);
        Put(lookup, 2); Put(lookup + 4, aliases);
        for (int i = 0; i < aliases; i++) Put(lookup + 6 + 2 * i, pair - lookup);
        Put(pair, 1); Put(pair + 2, coverage - pair); Put(pair + 4, 4); Put(pair + 8, glyphs);
        for (int i = 0; i < glyphs; i++) Put(pair + 10 + 2 * i, (nonzero && i == 1 ? nonzeroSet : emptySet) - pair);
        Put(coverage, range ? 2 : 1); Put(coverage + 2, range ? 1 : glyphs);
        if (range) Put(coverage + 6, glyphs - 1);
        else for (int i = 0; i < glyphs; i++) Put(coverage + 4 + 2 * i, i);
        Put(nonzeroSet, 1); Put(nonzeroSet + 2, 2); BinaryPrimitives.WriteInt16BigEndian(data.AsSpan(nonzeroSet + 4, 2), -25);
        byte[] seed = portable ? TestFontBuilder.ReplaceTable(File.ReadAllBytes(Case("PortableFonts/DejaVuSans.ttf")), "kern", [0, 0, 0, 0])
            : TestFontBuilder.CreateTestFont();
        return TestFontBuilder.ReplaceTable(seed, "GPOS", data);
    }
}
