using System.Buffers.Binary;
using System.Reflection;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void GposSelectionPreservesDeclaredFeatureSemantics()
    {
        short Read(byte[] gpos) => OpenTypeFont.Load(
            TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos)).GetKerning(1, 2);
        byte[] original = SelectionTable(1, 1, nonzero: true);
        TestAssert.Equal((short)-25, Read(original));

        byte[] Change(Action<byte[]> change)
        {
            byte[] copy = original.ToArray(); change(copy); return copy;
        }
        // Empty/unsupported language declarations retain the established fallback.
        TestAssert.Equal((short)-25, Read(Change(b => SelectionWrite(b, 26, 0))));
        TestAssert.Equal((short)-25, Read(Change(b => "zzzz"u8.CopyTo(b.AsSpan(12)))));
        // An explicit nonmatching feature declaration keeps that distinction.
        TestAssert.Equal((short)0, Read(Change(b => SelectionWrite(b, 28, 1))));
        TestAssert.Equal((short)0, Read(Change(b => SelectionWrite(b, 28, ushort.MaxValue))));
        int featureList = SelectionRead(original, 6);
        int feature = featureList + SelectionRead(original, featureList + 6);
        TestAssert.Equal((short)0, Read(Change(b => SelectionWrite(b, feature + 4, 1))));
        TestAssert.Equal((short)0, Read(Change(b => SelectionWrite(b, feature + 4, ushort.MaxValue))));
        TestAssert.Equal((short)0, Read(Change(b => "mark"u8.CopyTo(b.AsSpan(featureList + 2)))));
    }

    public static void GposLookupSelectionStorageStaysWithinIndexDomain()
    {
        byte[] large = SelectionTable(1, 30000, nonzero: false);
        ReadSelection(large, default, out _); // Warm the same reflection/parser path.
        int[] selected = ReadSelection(large, default, out long allocated);
        TestAssert.Equal(30000, selected.Length);
        TestAssert.Equal(0, selected[0]);
        TestAssert.Equal(29999, selected[^1]);
        TestAssert.True(allocated < 128 * 1024,
            "A16-bit selection must not construct megabyte-sized growing index sets.");

        byte[] sparse = SelectionTable(1, 3, nonzero: false);
        int featureList = SelectionRead(sparse, 6);
        int feature = featureList + SelectionRead(sparse, featureList + 6);
        SelectionWrite(sparse, feature + 4, 1);
        SelectionWrite(sparse, feature + 6, 1);
        SelectionWrite(sparse, feature + 8, ushort.MaxValue);
        int[] sparseSelection = ReadSelection(sparse, default, out _);
        TestAssert.Equal(1, sparseSelection.Length);
        TestAssert.Equal(1, sparseSelection[0]);
    }

    public static void GposMetadataSelectionObservesCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            ReadSelection(SelectionTable(4, 16, nonzero: false), cancellation.Token, out _));
    }

    private static int[] ReadSelection(byte[] gpos, CancellationToken token, out long allocated)
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        // Keep the regression harness usable against the committed parent too.
        MethodInfo method = typeof(OpenTypeFont).GetMethod("ReadGposKernLookupSelection", flags)
            ?? TestAssert.NotNull(typeof(OpenTypeFont).GetMethod("ReadGposKernLookupIndices", flags));
        Type builderType = TestAssert.NotNull(typeof(OpenTypeFont).GetNestedType("KerningPairBuilder", BindingFlags.NonPublic));
        using var builder = (IDisposable)TestAssert.NotNull(Activator.CreateInstance(builderType,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, [token, null], null));
        int count = SelectionRead(gpos, SelectionRead(gpos, 8));
        object?[] arguments = method.GetParameters().Length == 3
            ? [gpos, 0, gpos.Length] : [gpos, 0, gpos.Length, count, builder];
        long before = GC.GetAllocatedBytesForCurrentThread();
        object? result;
        try { result = method.Invoke(null, arguments); }
        catch (TargetInvocationException ex) when (ex.InnerException is not null) { throw ex.InnerException; }
        allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        return result switch
        {
            null => [],
            bool[] mask => Enumerable.Range(0, mask.Length).Where(i => mask[i]).ToArray(),
            IEnumerable<ushort> indices => indices.Where(i => i < count).Select(i => (int)i).Order().ToArray(),
            _ => throw new InvalidOperationException("Unexpected kerning selection result.")
        };
    }

    // Only the GPOS data under test is synthesized here. Independent whole-face
    // probe controls use the pinned portable face and rebuilt font directory.
    private static byte[] SelectionTable(int features, int lookups, bool nonzero)
    {
        int lookupList = 28 + 2 * features;
        int lookup = lookupList + 2 + 2 * lookups, pair = lookup + 8;
        int featureList = pair + (nonzero ? 24 : 20), feature = featureList + 2 + 6 * features;
        byte[] bytes = new byte[feature + 4 + 2 * lookups];
        SelectionWrite(bytes, 0, 1); SelectionWrite(bytes, 4, 10);
        SelectionWrite(bytes, 6, featureList); SelectionWrite(bytes, 8, lookupList);
        SelectionWrite(bytes, 10, 1); "DFLT"u8.CopyTo(bytes.AsSpan(12)); SelectionWrite(bytes, 16, 8);
        SelectionWrite(bytes, 18, 4); SelectionWrite(bytes, 24, ushort.MaxValue); SelectionWrite(bytes, 26, features);
        for (int i = 0; i < features; i++) SelectionWrite(bytes, 28 + i * 2, i);
        SelectionWrite(bytes, lookupList, lookups);
        for (int i = 0; i < lookups; i++) SelectionWrite(bytes, lookupList + 2 + i * 2, lookup - lookupList);
        SelectionWrite(bytes, lookup, 2); SelectionWrite(bytes, lookup + 4, 1); SelectionWrite(bytes, lookup + 6, 8);
        SelectionWrite(bytes, pair, 1); SelectionWrite(bytes, pair + 2, 12); SelectionWrite(bytes, pair + 4, 4);
        SelectionWrite(bytes, pair + 8, 1); SelectionWrite(bytes, pair + 10, 18);
        SelectionWrite(bytes, pair + 12, 1); SelectionWrite(bytes, pair + 14, 1); SelectionWrite(bytes, pair + 16, 1);
        SelectionWrite(bytes, pair + 18, nonzero ? 1 : 0);
        if (nonzero) { SelectionWrite(bytes, pair + 20, 2); BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(pair + 22, 2), -25); }
        SelectionWrite(bytes, featureList, features);
        for (int i = 0; i < features; i++)
        {
            int row = featureList + 2 + i * 6;
            "kern"u8.CopyTo(bytes.AsSpan(row)); SelectionWrite(bytes, row + 4, feature - featureList);
        }
        SelectionWrite(bytes, feature + 2, lookups);
        for (int i = 0; i < lookups; i++) SelectionWrite(bytes, feature + 4 + i * 2, i);
        return bytes;
    }

    private static int SelectionRead(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));
    private static void SelectionWrite(byte[] bytes, int offset, int value) => BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset, 2), checked((ushort)value));
}
