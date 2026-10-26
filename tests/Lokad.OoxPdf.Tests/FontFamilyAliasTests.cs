using Lokad.OoxPdf.Fonts;

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
