using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetKerningLeftRowsAdmitsReachedConstructionAndReusesAliases()
    {
        long Peak(int aliases, int covered)
        {
            byte[] bytes = OoxLimitsTests.BuildLeftClassFont(15000, aliases: aliases, coveredGlyphs: covered);
            var pool = new TestMemoryBudget();
            using (var scope = OoxConversionBudget.BeginScope(null, pool))
            {
                var font = OpenTypeFont.LoadShared(bytes, 0, default);
                TestAssert.Equal((short)0, font.GetKerning(0, 0));
                TestAssert.Equal((short)0, font.GetKerning(14999, 0));
                if (covered == 15000)
                    TestAssert.True(pool.Peak > pool.Reserved + 1024 * 1024,
                        "Reached zero-adjustment rows need construction capacity before their cache is built.");
                TestAssert.True(font.Bytes.Span.SequenceEqual(bytes), "Admission preserves the source program.");
                GC.KeepAlive(font);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
            return pool.Peak;
        }
        long one = Peak(1, 15000), repeated = Peak(2, 15000), sparse = Peak(1, 2);
        TestAssert.True(Math.Abs(repeated - one) < 4096,
            "Repeated source aliases reuse the construction window instead of accumulating a work charge.");
        TestAssert.True(sparse + 1024 * 1024 < one,
            "Declared but unreached classes must not consume the reached-row allowance.");
    }

    public static void MemoryBudgetKerningLeftRowsRefusesBeforeConstructionAndRefundsCancellation()
    {
        byte[] sparse = OoxLimitsTests.BuildLeftClassFont(15000, coveredGlyphs: 2);
        byte[] dense = OoxLimitsTests.BuildLeftClassFont(15000, nonzero: true);
        var pool = new TestMemoryBudget(1536 * 1024);
        using (var scope = OoxConversionBudget.BeginScope(null, pool))
            TestAssert.Equal((short)0, OpenTypeFont.LoadShared(sparse, 0, default).GetKerning(0, 0));
        TestAssert.Equal(0L, pool.Reserved);
        long before = GC.GetAllocatedBytesForCurrentThread();
        using (var scope = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OpenTypeFont.LoadShared(dense, 0, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.True(allocated < 2 * 1024 * 1024,
            $"Refusal must precede the multi-megabyte reached-row cache, rather than follow it: {allocated}.");
        TestAssert.Equal(0L, pool.Reserved);
        using var cancellation = new CancellationTokenSource();
        var cancelPool = new TestMemoryBudget { RefuseChange = (_, after) =>
        {
            if (after > 2 * 1024 * 1024) cancellation.Cancel();
            return false;
        }};
        byte[] larger = OoxLimitsTests.BuildLeftClassFont(16000);
        using (var scope = OoxConversionBudget.BeginScope(null, cancelPool))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() => OpenTypeFont.LoadShared(larger, 0, cancellation.Token));
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
        }
        TestAssert.Equal(0L, cancelPool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(cancelPool.ScopesCreated, cancelPool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
        TestAssert.Equal(0, cancelPool.CallerDisposals);
    }

    public static void MemoryBudgetKerningMetadataFundsConstructionAndRefundsBeforeUse()
    {
        byte[] bytes = OoxLimitsTests.BuildRightClassFont(16384, nonzero: true);
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var font = OpenTypeFont.LoadShared(bytes, 0, default);
            TestAssert.True(pool.Peak > pool.Reserved + 1024 * 1024,
                "Significant class construction is admitted before decoding and refunded after its last batch user.");
            TestAssert.True(pool.Reserved >= bytes.Length, "The private program remains owned while its font is in use.");
            TestAssert.Equal((short)-50, font.GetKerning(0, 0));
            TestAssert.Equal((short)-50, font.GetKerning(0, 65532));
            TestAssert.True(font.Bytes.Span.SequenceEqual(bytes), "Admission does not change the source program.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetKerningMetadataReusesZeroBatchWindow()
    {
        long Peak(int aliases)
        {
            byte[] bytes = OoxLimitsTests.BuildRightClassFont(16384, aliases: aliases);
            var pool = new TestMemoryBudget(4 * 1024 * 1024);
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var font = OpenTypeFont.LoadShared(bytes, 0, default);
                TestAssert.Equal((short)0, font.GetKerning(0, 0));
                GC.KeepAlive(font);
            }
            TestAssert.Equal(0L, pool.Reserved);
            return pool.Peak;
        }
        long one = Peak(1), repeated = Peak(8);
        TestAssert.True(repeated < one + 4096, "Repeated zero-adjustment subtables reuse construction capacity rather than accumulating a work charge.");
    }

    public static void MemoryBudgetKerningMetadataRefusalAndCancellationStayTerminal()
    {
        byte[] small = OoxLimitsTests.BuildRightClassFont(2), dense = OoxLimitsTests.BuildRightClassFont(16384);
        var pool = new TestMemoryBudget(1536 * 1024);
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            TestAssert.Equal((short)0, OpenTypeFont.LoadShared(small, 0, default).GetKerning(0, 0));
        TestAssert.Equal(0L, pool.Reserved);
        var face = ProgramResolution(new MemoryFontProgramSource("right-class-metadata", dense));
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(face, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(0L, pool.Reserved);
        }
        using var cancellation = new CancellationTokenSource();
        var cancelPool = new TestMemoryBudget { RefuseChange = (before, after) =>
        {
            if (before <= 2 * 1024 * 1024 && after > 2 * 1024 * 1024) cancellation.Cancel();
            return false;
        }};
        using (var conversion = OoxConversionBudget.BeginScope(null, cancelPool))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() => FontProgramLoader.Load(face, cancellation.Token));
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.Equal(0L, cancelPool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(cancelPool.ScopesCreated, cancelPool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
        TestAssert.Equal(0, cancelPool.CallerDisposals);
    }
}
