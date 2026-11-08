using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxLimitsTests
{
    public static void ClassKerningUnusableRightClassesSkipLeftRowConstruction()
    {
        byte[] small = BuildLeftClassFont(128, rightClass: false), dense = BuildLeftClassFont(30000, rightClass: false);
        OpenTypeFont.LoadShared(small, 0, default); OpenTypeFont.LoadShared(dense, 0, default);
        long Measure(byte[] bytes)
        {
            long before = GC.GetAllocatedBytesForCurrentThread();
            var font = OpenTypeFont.LoadShared(bytes, 0, default);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            TestAssert.Equal((short)0, font.GetKerning(0, 0));
            TestAssert.Equal((short)0, font.GetKerning(29999, 0));
            GC.KeepAlive(font);
            return allocated;
        }
        long sparse = Measure(small), expanded = Measure(dense);
        TestAssert.True(expanded - sparse < 512 * 1024,
            $"No eligible right class means no reached-left row cache, even with many source rows: {sparse} -> {expanded}.");
    }

    public static void ClassKerningReachedLeftRowsKeepNonzeroAndAliasSemantics()
    {
        foreach (int aliases in new[] { 1, 2 })
        {
            byte[] bytes = BuildLeftClassFont(15000, nonzero: true, aliases: aliases);
            var font = OpenTypeFont.LoadShared(bytes, 0, default);
            TestAssert.Equal((short)-50, font.GetKerning(0, 0));
            TestAssert.Equal((short)-50, font.GetKerning(14999, 0));
            TestAssert.Equal((short)0, font.GetKerning(15000, 0));
            TestAssert.Equal((short)0, font.GetKerning(0, 1));
            TestAssert.True(font.Bytes.Span.SequenceEqual(bytes), "Source aliases and admission preserve the same original program.");
        }
    }

    public static void ClassKerningRightClassesAvoidGrowableListsAndKeepSparseOutput()
    {
        byte[] small = BuildRightClassFont(2), dense = BuildRightClassFont(16384);
        OpenTypeFont.Load(small); OpenTypeFont.Load(dense);
        long Measure(byte[] bytes)
        {
            long before = GC.GetAllocatedBytesForCurrentThread();
            var font = OpenTypeFont.Load(bytes);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            TestAssert.Equal((ushort)65535, font.GlyphCount);
            TestAssert.Equal((short)0, font.GetKerning(0, 0));
            GC.KeepAlive(font);
            return allocated;
        }
        long sparseBytes = Measure(small), denseBytes = Measure(dense);
        TestAssert.True(denseBytes - sparseBytes < 2 * 1024 * 1024,
            $"Many source classes must not each construct growable lists and copied arrays: {sparseBytes} -> {denseBytes}.");
        var nonzero = OpenTypeFont.Load(BuildRightClassFont(16384, nonzero: true));
        TestAssert.Equal((short)-50, nonzero.GetKerning(0, 0));
        TestAssert.Equal((short)-50, nonzero.GetKerning(0, 65532));
        TestAssert.Equal((short)0, nonzero.GetKerning(0, 1));
        TestAssert.Equal((short)0, nonzero.GetKerning(1, 0));
    }

    public static void ClassKerningRightRangesKeepClassAndDuplicateSourceOrder()
    {
        byte[] gpos = BuildClassKernGpos(1, 1, rangeCoverage: false);
        int list = U(gpos, 8), lookup = list + U(gpos, list + 2), subtable = lookup + U(gpos, lookup + 6);
        int left = subtable + U(gpos, subtable + 8), right = subtable + U(gpos, subtable + 10);
        W(gpos, subtable + 12, 1); W(gpos, subtable + 14, 3); W(gpos, left + 6, 0);
        BinaryPrimitives.WriteInt16BigEndian(gpos.AsSpan(subtable + 18, 2), 3);
        BinaryPrimitives.WriteInt16BigEndian(gpos.AsSpan(subtable + 20, 2), -7);
        Array.Resize(ref gpos, right + 22);
        ushort[] definition = [2, 3, 0, 2, 2, 1, 1, 1, 1, 1, 2];
        for (int i = 0; i < definition.Length; i++) W(gpos, right + 2 * i, definition[i]);
        var font = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos));
        TestAssert.Equal((short)-7, font.GetKerning(0, 0));
        TestAssert.Equal((short)3, font.GetKerning(0, 1));
        TestAssert.Equal((short)-7, font.GetKerning(0, 2));
    }

    internal static byte[] BuildRightClassFont(int classes, bool nonzero = false, int aliases = 1)
    {
        const int glyphs = 65535;
        byte[] font = TestFontBuilder.CreateTestFont();
        byte[] maxp = Table(font, "maxp"); W(maxp, 4, glyphs); font = TestFontBuilder.ReplaceTable(font, "maxp", maxp);
        byte[] hhea = Table(font, "hhea"); W(hhea, 34, glyphs); font = TestFontBuilder.ReplaceTable(font, "hhea", hhea);
        byte[] metrics = new byte[glyphs * 4];
        for (int i = 0; i < glyphs; i++) W(metrics, 4 * i, 600);
        font = TestFontBuilder.ReplaceTable(font, "hmtx", metrics);
        byte[] oldLoca = Table(font, "loca"); int stride = U(Table(font, "head"), 50) == 0 ? 2 : 4;
        byte[] loca = new byte[(glyphs + 1) * stride]; oldLoca.CopyTo(loca, 0);
        for (int i = oldLoca.Length; i < loca.Length; i += stride)
            oldLoca.AsSpan(oldLoca.Length - stride, stride).CopyTo(loca.AsSpan(i, stride));
        font = TestFontBuilder.ReplaceTable(font, "loca", loca);
        byte[] seed = BuildClassKernGpos(1, 1, rangeCoverage: false);
        int list = U(seed, 8), lookup = list + U(seed, list + 2), subtable = lookup + 6 + 2 * aliases;
        int coverage = 16 + classes * 2, left = coverage + 6, right = left + 8;
        byte[] gpos = new byte[subtable + right + 6 + glyphs * 2]; seed.AsSpan(0, lookup + 6).CopyTo(gpos);
        W(gpos, lookup + 4, aliases);
        for (int i = 0; i < aliases; i++) W(gpos, lookup + 6 + 2 * i, subtable - lookup);
        W(gpos, subtable, 2); W(gpos, subtable + 2, coverage); W(gpos, subtable + 4, 4);
        W(gpos, subtable + 8, left); W(gpos, subtable + 10, right);
        W(gpos, subtable + 12, 1); W(gpos, subtable + 14, classes);
        if (nonzero) BinaryPrimitives.WriteInt16BigEndian(gpos.AsSpan(subtable + 18, 2), -50);
        W(gpos, subtable + coverage, 1); W(gpos, subtable + coverage + 2, 1);
        W(gpos, subtable + left, 1); W(gpos, subtable + left + 4, 1);
        W(gpos, subtable + right, 1); W(gpos, subtable + right + 4, glyphs);
        for (int i = 0; i < glyphs; i++) W(gpos, subtable + right + 6 + i * 2, 1 + i % (classes - 1));
        return TestFontBuilder.ReplaceTable(font, "GPOS", gpos);

        static byte[] Table(byte[] source, string name)
        {
            var table = TestFontBuilder.GetTableRange(source, name);
            return source.AsSpan(table.Offset, table.Length).ToArray();
        }
    }

    internal static byte[] BuildLeftClassFont(int classes, bool rightClass = true, bool nonzero = false,
        int aliases = 1, int? coveredGlyphs = null)
    {
        byte[] font = BuildRightClassFont(2);
        byte[] seed = BuildClassKernGpos(1, 1, rangeCoverage: false);
        int list = U(seed, 8), lookup = list + U(seed, list + 2), subtable = lookup + 6 + 2 * aliases;
        int columns = rightClass ? 2 : 1;
        int coverage = 16 + classes * columns * 2, right = coverage + 10, left = right + 8;
        int covered = coveredGlyphs ?? classes;
        if (covered < 1 || covered > classes) throw new ArgumentOutOfRangeException(nameof(coveredGlyphs));
        byte[] gpos = new byte[subtable + left + 6 + classes * 2];
        seed.AsSpan(0, lookup + 6).CopyTo(gpos);
        W(gpos, lookup + 4, aliases);
        for (int i = 0; i < aliases; i++) W(gpos, lookup + 6 + 2 * i, subtable - lookup);
        W(gpos, subtable, 2); W(gpos, subtable + 2, coverage); W(gpos, subtable + 4, 4);
        W(gpos, subtable + 8, left); W(gpos, subtable + 10, right);
        W(gpos, subtable + 12, classes); W(gpos, subtable + 14, columns);
        if (nonzero)
            for (int i = 0; i < classes; i++)
                BinaryPrimitives.WriteInt16BigEndian(gpos.AsSpan(subtable + 16 + 2 * (i * columns + 1)), -50);
        W(gpos, subtable + coverage, 2); W(gpos, subtable + coverage + 2, 1);
        W(gpos, subtable + coverage + 6, covered - 1);
        W(gpos, subtable + right, 1); W(gpos, subtable + right + 4, 1);
        W(gpos, subtable + right + 6, rightClass ? 1 : 0);
        W(gpos, subtable + left, 1); W(gpos, subtable + left + 4, classes);
        for (int i = 0; i < classes; i++) W(gpos, subtable + left + 6 + i * 2, i);
        return TestFontBuilder.ReplaceTable(font, "GPOS", gpos);
    }

    private static ushort U(byte[] b, int i) => BinaryPrimitives.ReadUInt16BigEndian(b.AsSpan(i, 2));
    private static void W(byte[] b, int i, int value) => BinaryPrimitives.WriteUInt16BigEndian(b.AsSpan(i, 2), checked((ushort)value));
}
