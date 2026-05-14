using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class FontTests
{
    public static void WindowsFontResolverFindsInstalledFonts()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        IReadOnlyList<FontResolution> fonts = resolver.GetDiscoveredFonts();
        if (fonts.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one discoverable Windows font.");
        }

        FontResolution resolved = resolver.Resolve(new FontRequest("Arial"));
        TestAssert.NotNull(resolved.FontFilePath);
    }

    public static void OpenTypeParserMapsBasicLatinGlyphs()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);

        TestAssert.Equal("Arial", font.FamilyName);
        TestAssert.True(font.UnitsPerEm > 0, "Expected a positive units-per-em value.");
        TestAssert.True(font.GlyphCount > 0, "Expected a positive glyph count from maxp.");
        TestAssert.True(font.Os2.WeightClass > 0, "Expected a positive OS/2 weight class.");
        TestAssert.True(font.Os2.WindowsAscender > 0, "Expected a positive OS/2 Windows ascender.");
        TestAssert.True(font.Post.UnderlineThickness != 0, "Expected a non-zero post underline thickness.");
        TestAssert.True(font.TableTags.Contains("cmap"), "Expected cmap table.");
        ushort glyph = font.MapCodePoint('A');
        TestAssert.True(glyph > 0, "Expected a glyph mapping for 'A'.");
        TestAssert.True(font.GetAdvanceWidth(glyph) > 0, "Expected a positive advance width for 'A'.");
    }
}
