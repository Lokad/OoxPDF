using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetFontCharacterMapsRefuseBeforeRetainedExpansion()
    {
        byte[] bytes = DenseCharacterMapProgram();
        var face = ProgramResolution(new MemoryFontProgramSource("caller-owned-dense-cmap", bytes));
        TestAssert.Equal((ushort)1, TestAssert.NotNull(FontProgramLoader.Load(face, default)).MapCodePoint(0x10000));
        // This host can fund the private SFNT, but not its 1.2 MB character map.
        var pool = new TestMemoryBudget(bytes.Length + 64 * 1024);
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(face, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
        TestAssert.Equal((ushort)1, TestAssert.NotNull(FontProgramLoader.Load(face, default)).MapCodePoint(0x20000));
    }

    public static void MemoryBudgetFontCharacterMapsRetainFundingAndRefundCancellation()
    {
        byte[] bytes = DenseCharacterMapProgram();
        var face = ProgramResolution(new MemoryFontProgramSource("caller-owned-dense-cmap", bytes));
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var first = TestAssert.NotNull(FontProgramLoader.Load(face, default));
            long retained = pool.Reserved;
            TestAssert.True(retained >= first.Bytes.Length + 1_200_000L,
                "The private SFNT and its retained character-map payload share the font owner's funding.");
            var second = TestAssert.NotNull(FontProgramLoader.Load(face, default));
            TestAssert.Equal(2L * retained, pool.Reserved);
            TestAssert.Equal((ushort)1, second.MapCodePoint(0x20000));
            TestAssert.True(first.Bytes.Span.SequenceEqual(bytes), "Parsing preserves the caller's source.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        using var cancellation = new CancellationTokenSource();
        bool reachedMetadata = false;
        var canceled = new TestMemoryBudget { RefuseChange = (before, after) =>
        {
            if (before == bytes.Length && after > before) { reachedMetadata = true; cancellation.Cancel(); }
            return false;
        }};
        using (var operation = OoxConversionBudget.BeginScope(null, canceled))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() => FontProgramLoader.Load(face, cancellation.Token));
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.True(reachedMetadata, "Cancel after private-program acquisition, at derived-storage admission.");
            TestAssert.Equal(0L, canceled.Reserved);
        }
        TestAssert.Equal(canceled.ScopesCreated, canceled.ScopesDisposed);
        TestAssert.Equal(0, canceled.CallerDisposals);
    }

    private static byte[] DenseCharacterMapProgram()
    {
        const int groups = 100_000;
        var cmap = new byte[28 + groups * 12];
        void U16(int offset, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(offset, 2), value);
        void U32(int offset, uint value) => BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset, 4), value);
        U16(2, 1); U16(4, 3); U16(6, 10); U32(8, 12);
        U16(12, 12); U32(16, (uint)cmap.Length - 12); U32(24, groups);
        for (uint index = 0; index < groups; index++)
        {
            int offset = 28 + (int)index * 12;
            U32(offset, 0x10000 + index); U32(offset + 4, 0x10000 + index); U32(offset + 8, 1);
        }
        return TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "cmap", cmap);
    }
}
