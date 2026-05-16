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

        FontResolution bold = resolver.Resolve(new FontRequest("Arial", Bold: true));
        TestAssert.NotNull(bold.FontFilePath);
        TestAssert.True(bold.WeightClass >= resolved.WeightClass, "Expected bold font resolution to prefer a heavier face when one is available.");
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

    public static void OpenTypeParserReadsGposPairAdjustments()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        if (!font.TableTags.Contains("GPOS"))
        {
            return;
        }

        ushort left = font.MapCodePoint('T');
        ushort right = font.MapCodePoint('o');
        TestAssert.True(font.GetKerning(left, right) != 0, "Expected Arial GPOS pair adjustment for 'To'.");
    }

    public static void OpenTypeParserIgnoresInactiveGposPairAdjustments()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambria = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(cambria);
        AssertNoKerning(font, 'L', 'o');
        AssertNoKerning(font, 'o', 'k');
        AssertNoKerning(font, 'q', 'u');
        AssertNoKerning(font, 'D', 'é');
    }

    public static void OpenTypeParserMapsWindowsSymbolCmap()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string symbol = Path.Combine(fontsDirectory, "symbol.ttf");
        if (!File.Exists(symbol))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(symbol);
        ushort glyph = font.MapCodePoint(0xF0B7);

        TestAssert.True(glyph > 0, "Expected a glyph mapping for the common Symbol bullet code point.");
    }

    private static void AssertNoKerning(OpenTypeFont font, char leftChar, char rightChar)
    {
        ushort left = font.MapCodePoint(leftChar);
        ushort right = font.MapCodePoint(rightChar);

        TestAssert.Equal((short)0, font.GetKerning(left, right));
    }

    public static void WindowsFontResolverMapsCambriaMathToCambria()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory) ||
            !Directory.EnumerateFiles(fontsDirectory, "cambria*.ttf", SearchOption.TopDirectoryOnly).Any())
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontResolution resolved = resolver.Resolve(new FontRequest("Cambria Math"));

        TestAssert.Equal("Cambria", resolved.FamilyName);
        TestAssert.NotNull(resolved.FontFilePath);
    }

    public static void WindowsFontResolverMapsAptosThemeFontsToCalibri()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory) ||
            !Directory.EnumerateFiles(fontsDirectory, "calibri*.ttf", SearchOption.TopDirectoryOnly).Any())
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontResolution display = resolver.Resolve(new FontRequest("Aptos Display"));
        FontResolution body = resolver.Resolve(new FontRequest("Aptos"));

        TestAssert.Equal("Calibri Light", display.FamilyName);
        TestAssert.Equal("Calibri", body.FamilyName);
        TestAssert.NotNull(display.FontFilePath);
        TestAssert.NotNull(body.FontFilePath);
    }

    public static void OpenTypeParserLoadsTrueTypeCollections()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambriaCollection = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambriaCollection))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(cambriaCollection);

        TestAssert.True(font.FamilyName.Length > 0, "Expected a family name from the first TTC face.");
        TestAssert.True(font.GlyphCount > 0, "Expected glyphs from the first TTC face.");
        TestAssert.True(font.TableTags.Contains("cmap"), "Expected cmap table from the first TTC face.");
    }
}
