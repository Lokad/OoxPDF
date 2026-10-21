using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxLimitsTests
{
    public static void ClassKerningZeroAdvanceScanCompletesWithoutWorkQuota()
    {
        // Repeated zero-adjustment class ranges scan beyond the retired work
        // ceiling without constructing any new pair storage.
        var font = OpenTypeFont.Load(BuildZeroAdvanceScanFont(1025, 1025));
        TestAssert.Equal((short)0, font.GetKerning(0, 0));
    }

    public static void ClassKerningZeroAdvanceScanObservesCancellation()
    {
        byte[] bytes = BuildZeroAdvanceScanFont(30_000, 8_000);
        var resolution = new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(),
            new MemoryFontProgramSource("test:scan-cancellation", bytes), IsFallback: false);
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(25));
        var watch = System.Diagnostics.Stopwatch.StartNew();
        OperationCanceledException failure = TestAssert.Throws<OperationCanceledException>(
            () => FontProgramLoader.Load(resolution, cancellation.Token));
        TestAssert.Equal(cancellation.Token, failure.CancellationToken);
        TestAssert.True(watch.Elapsed < TimeSpan.FromSeconds(10), "Cancellation must interrupt a long scan, without becoming font fallback.");
    }

    public static void ClassKerningStorageRefusalStillUsesFontFallback()
    {
        byte[] bytes = TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS",
            BuildClassKernGpos(2048, 2048, rangeCoverage: false));
        var resolution = new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(),
            new MemoryFontProgramSource("test:pair-refusal", bytes), IsFallback: false);
        TestAssert.Equal<OpenTypeFont?>(null, FontProgramLoader.Load(resolution, default));
    }

    private static byte[] BuildZeroAdvanceScanFont(int aliases, int rangeCount)
    {
        byte[] gpos = BuildAliasedClassKernGpos(128, 128, aliases);
        int lookupList = BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(8, 2));
        int lookup = lookupList + BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookupList + 2, 2));
        int subtable = lookup + BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookup + 6, 2));
        int classDef = subtable + BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(subtable + 8, 2));
        const int oldLength = 6 + 128 * 2;
        byte[] replacement = new byte[4 + 6 * rangeCount];
        BinaryPrimitives.WriteUInt16BigEndian(replacement.AsSpan(0, 2), 2);
        BinaryPrimitives.WriteUInt16BigEndian(replacement.AsSpan(2, 2), checked((ushort)rangeCount));
        for (int index = 0; index < rangeCount; index++)
        {
            int offset = 4 + 6 * index;
            bool last = index == rangeCount - 1;
            BinaryPrimitives.WriteUInt16BigEndian(replacement.AsSpan(offset, 2), last ? (ushort)0 : (ushort)129);
            BinaryPrimitives.WriteUInt16BigEndian(replacement.AsSpan(offset + 2, 2), last ? (ushort)127 : (ushort)129);
            BinaryPrimitives.WriteUInt16BigEndian(replacement.AsSpan(offset + 4, 2), 1);
        }
        int delta = replacement.Length - oldLength;
        byte[] changed = new byte[gpos.Length + delta];
        gpos.AsSpan(0, classDef).CopyTo(changed);
        replacement.CopyTo(changed, classDef);
        gpos.AsSpan(classDef + oldLength).CopyTo(changed.AsSpan(classDef + replacement.Length));
        ushort classDef2 = BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(subtable + 10, 2));
        BinaryPrimitives.WriteUInt16BigEndian(changed.AsSpan(subtable + 10, 2), checked((ushort)(classDef2 + delta)));
        BinaryPrimitives.WriteInt16BigEndian(changed.AsSpan(subtable + 22, 2), 0);
        return TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", changed);
    }

    public static void ClassKerningSharedReadsRetainSignedLastWinsAndOwnedSource()
    {
        foreach (int size in new[] { 8, 256 })
        foreach (short later in new[] { short.MinValue, (short)0, short.MaxValue })
        {
            byte[] gpos = BuildAliasedClassKernGpos(size, size, 2);
            int lookupList = BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(8, 2));
            int lookup = lookupList + BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookupList + 2, 2));
            int subtable = lookup + BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookup + 6, 2));
            byte[] second = gpos[subtable..];
            BinaryPrimitives.WriteInt16BigEndian(second.AsSpan(22, 2), later);
            BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(lookup + 8, 2), checked((ushort)(gpos.Length - lookup)));
            gpos = gpos.Concat(second).ToArray();
            byte[] bytes = TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos);
            var font = OpenTypeFont.Load(bytes, 0);
            // The existing reader skips zero-advance records rather than
            // replacing a previously nonzero pair. Both signed limits survive.
            short expected = later == 0 ? (short)-50 : later;
            var range = TestFontBuilder.GetTableRange(bytes, "GPOS");
            Array.Clear(bytes, range.Offset, range.Length);
            Parallel.For(0, 4, _ =>
            {
                for (int left = 0; left < size; left++)
                for (int right = 0; right < size; right++)
                    TestAssert.Equal(expected, font.GetKerning((ushort)left, (ushort)right));
            });
            TestAssert.Equal((short)0, font.GetKerning((ushort)(size + 1), 0));
            TestAssert.Equal((short)0, font.GetKerning(0, (ushort)(size + 1)));
        }
    }
}
