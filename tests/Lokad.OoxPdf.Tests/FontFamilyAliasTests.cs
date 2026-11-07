using Lokad.OoxPdf.Fonts;
using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class FontFamilyAliasTests
{
    private static readonly (ushort Language, string Name)[] Names = [(0x0404, "字體"), (0x0409, "Review Family"), (0x0411, "字體"), (0x0409, "REVIEW FAMILY")];

    public static void FontFamilyAliasesPreserveCanonicalIdentityAndSharedSource() => InDirectory(directory =>
    {
        byte[] bytes = TestFontBuilder.CreateFontWithFamilyNames(Names);
        File.WriteAllBytes(Path.Combine(directory, "face.ttf"), bytes);
        OpenTypeFont.FontDiscoveryHeaders headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
        TestAssert.Equal("字體", headers.FamilyName);
        TestAssert.Equal("字體|Review Family", string.Join("|", headers.FamilyNames));
        TestAssert.Equal("字體", OpenTypeFont.Load(bytes).FamilyName);
        var resolver = new WindowsFontResolver(directory);
        FontFaceResolution english = resolver.Resolve(new("Review Family"));
        FontFaceResolution localized = resolver.Resolve(new("字體"));
        TestAssert.True(!english.IsFallback && !localized.IsFallback, "Both declared aliases must resolve exactly.");
        TestAssert.Equal("Review Family", english.RequestedFamily);
        TestAssert.Equal("REVIEW FAMILY", resolver.Resolve(new("REVIEW FAMILY")).RequestedFamily);
        TestAssert.True(ReferenceEquals(english.Source, localized.Source), "Aliases must borrow the same owned program.");
        TestAssert.Equal(1, resolver.GetDiscoveredFonts().Count);
        TestAssert.True(FontProgramLoader.Load(english, default)!.MapCodePoint(0x1F600) != 0, "Alias lookup must retain supplementary glyph coverage.");
    });

    public static void FontFamilyAliasesKeepCollectionFaceOwnership() => InDirectory(directory =>
    {
        byte[] bytes = TestFontBuilder.CreateCollection("字體", "Other Family", Names);
        File.WriteAllBytes(Path.Combine(directory, "faces.ttc"), bytes);
        var resolver = new WindowsFontResolver(directory);
        FontFaceResolution alias = resolver.Resolve(new("Review Family"));
        FontFaceResolution other = resolver.Resolve(new("Other Family"));
        TestAssert.True(!alias.IsFallback && !other.IsFallback, "Collection names must resolve their own faces.");
        TestAssert.Equal(0, alias.FontFaceIndex);
        TestAssert.Equal(1, other.FontFaceIndex);
        TestAssert.True(ReferenceEquals(alias.Source, other.Source), "Collection faces share the file source, not the face index.");
        TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
        TestAssert.Equal("字體", FontProgramLoader.Load(alias, default)!.FamilyName);
        TestAssert.Equal("Other Family", FontProgramLoader.Load(other, default)!.FamilyName);
    });

    public static void FontFamilyAliasesRetainStyleRanking() => InDirectory(directory =>
    {
        byte[] regular = TestFontBuilder.CreateFontWithFamilyNames(Names);
        byte[] bold = regular.ToArray();
        int os2 = TestFontBuilder.GetTableRange(bold, "OS/2").Offset;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(bold.AsSpan(os2+4, 2), 700);
        File.WriteAllBytes(Path.Combine(directory, "regular.ttf"), regular);
        File.WriteAllBytes(Path.Combine(directory, "bold.ttf"), bold);
        var resolver = new WindowsFontResolver(directory);
        foreach (string family in new[] { "Review Family", "字體" }) {
            FontFaceResolution normal = resolver.Resolve(new(family));
            FontFaceResolution heavy = resolver.Resolve(new(family, Bold: true));
            TestAssert.Equal(400, normal.WeightClass);
            TestAssert.Equal(700, heavy.WeightClass);
            TestAssert.True(!normal.Bold && heavy.Bold, "Aliases keep ordinary style selection.");
        }
        TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
    });

    public static void FontCanonicalLoadDoesNotConstructUnusedDiscoveryAliases()
    {
        byte[] bytes = SharedNameProgram();
        var headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
        TestAssert.Equal(128, headers.FamilyNames.Count);
        TestAssert.Equal(10_000, headers.FamilyName.Length);
        TestAssert.Equal('\u1000', headers.FamilyName[0]);
        TestAssert.Equal('\u107F', headers.FamilyNames[^1][0]);
        OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        long before = GC.GetAllocatedBytesForCurrentThread();
        var font = OpenTypeFont.Load(bytes);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal(headers.FamilyName, font.FamilyName);
        TestAssert.True(font.MapCodePoint(0x1F600) != 0, "Canonical selection preserves supplementary glyph coverage.");
        TestAssert.True(allocated < 128 * 1024,
            $"Private parsing needs one family name, not the discovery alias catalog: allocated={allocated}.");
    }

    public static void FontCanonicalLoadRetainsLateNameBoundsAndMalformedRecovery()
    {
        byte[] bytes = SharedNameProgram();
        int name = TestFontBuilder.GetTableRange(bytes, "name").Offset;
        // The first canonical name is valid. A later unused alias still has to
        // satisfy the same source bounds before the parser may accept the face.
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(name + 6 + 127 * 12 + 10, 2), ushort.MaxValue);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(bytes));
        var resolution = new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(),
            new MemoryFontProgramSource("late-name-bounds", bytes), IsFallback: false);
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            TestAssert.True(FontProgramLoader.Load(resolution, default) is null,
                "Malformed name storage keeps ordinary font fallback rather than accepting an unchecked later record.");
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void FontCanonicalSelectionPreservesFallbackAndEmptyWindowsNames()
    {
        foreach (bool windows in new[] { false, true })
        {
            var names = windows ? new[] { ((ushort)0x0409, "Fallback"), ((ushort)0x0409, "Windows Family") }
                : new[] { ((ushort)0x0409, "Fallback") };
            byte[] bytes = TestFontBuilder.CreateFontWithFamilyNames(names);
            int name = TestFontBuilder.GetTableRange(bytes, "name").Offset;
            int storage = name + BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(name + 4, 2));
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(name + 6, 2), 1);
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(name + 6 + 8, 2), 8);
            Encoding.ASCII.GetBytes("Fallback").CopyTo(bytes, storage);
            string expected = windows ? "Windows Family" : "Fallback";
            TestAssert.Equal(expected, OpenTypeFont.Load(bytes).FamilyName);
            TestAssert.Equal(expected, OpenTypeFont.ReadDiscoveryHeaders(bytes, 0).FamilyName);
        }
        byte[] empty = TestFontBuilder.CreateFontWithFamilyNames((0x0409, ""), (0x0411, "Later Family"));
        TestAssert.Equal("", OpenTypeFont.Load(empty).FamilyName);
        TestAssert.Equal("|Later Family", string.Join("|", OpenTypeFont.ReadDiscoveryHeaders(empty, 0).FamilyNames));
    }

    internal static byte[] SharedNameProgram()
    {
        const int aliases = 128, characters = 10_000;
        int storage = 6 + aliases * 12;
        byte[] name = new byte[storage + (characters + aliases - 1) * 2];
        void U16(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(name.AsSpan(at, 2), value);
        U16(2, aliases); U16(4, (ushort)storage);
        for (ushort index = 0; index < aliases; index++)
        {
            int at = 6 + index * 12;
            U16(at, 3); U16(at + 2, 1); U16(at + 4, (ushort)(0x0400 + index));
            U16(at + 6, 1); U16(at + 8, characters * 2); U16(at + 10, (ushort)(index * 2));
        }
        for (int index = 0; index < characters + aliases - 1; index++)
            U16(storage + index * 2, (ushort)(0x1000 + index));
        return TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "name", name);
    }

    private static void InDirectory(Action<string> action)
    {
        string directory = Path.Combine(Path.GetTempPath(), "font-alias-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try { action(directory); }
        finally {
            WindowsFontResolver.InvalidateDiscoveryCaches();
            Directory.Delete(directory, recursive: true);
        }
    }
}
