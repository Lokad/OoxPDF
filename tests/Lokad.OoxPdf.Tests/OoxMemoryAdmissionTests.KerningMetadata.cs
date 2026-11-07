using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
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
            if (before == 0 && after > 2 * 1024 * 1024) cancellation.Cancel();
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
