using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxLimitsTests
{
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
