using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

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
        AssertNoKerning(font, 'D', 'é');
    }

    public static void OpenTypeParserReadsGposExtensionKerning()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambria = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(cambria);
        TestAssert.True(
            HasAnyKerning(font, "Lokad en quelques mots") ||
            HasAnyKerning(font, "Dépendance à l'offre") ||
            HasAnyKerning(font, "Large Global Supply Network") ||
            HasAnyKerning(font, "The scale and growth"),
            "Expected at least one Cambria extension GPOS kerning pair in the typography probe words.");
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

    private static void AssertHasKerning(OpenTypeFont font, char leftChar, char rightChar)
    {
        ushort left = font.MapCodePoint(leftChar);
        ushort right = font.MapCodePoint(rightChar);

        TestAssert.True(font.GetKerning(left, right) != 0, $"Expected kerning for '{leftChar}{rightChar}'.");
    }

    private static bool HasAnyKerning(OpenTypeFont font, string text)
    {
        ushort previous = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = font.MapCodePoint(rune.Value);
            if (previous != 0 && glyph != 0 && font.GetKerning(previous, glyph) != 0)
            {
                return true;
            }

            previous = glyph;
        }

        return false;
    }

    public static void WindowsFontResolverKeepsExactMathCollectionFace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory) ||
            !Directory.EnumerateFiles(fontsDirectory, "*.ttc", SearchOption.TopDirectoryOnly).Any())
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontResolution? mathFace = resolver.GetDiscoveredFonts()
            .FirstOrDefault(f => f.HasMathTable && f.FontFilePath is not null && f.FontFilePath.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase));
        if (mathFace is null)
        {
            return;
        }

        FontResolution resolved = resolver.Resolve(new FontRequest(mathFace.FamilyName));

        TestAssert.NotNull(resolved.FontFilePath);
        TestAssert.True(resolved.HasMathTable, "Expected exact font resolution to keep the requested math-table face.");
        TestAssert.Equal(mathFace.FontFilePath, resolved.FontFilePath);
        TestAssert.True(!resolved.IsFallback, "Expected exact font resolution not to be marked as fallback.");
    }

    public static void WindowsFontResolverUsesMetadataRatherThanFontNameAliases()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontResolution display = resolver.Resolve(new FontRequest("Aptos Display"));
        FontResolution body = resolver.Resolve(new FontRequest("Aptos"));

        TestAssert.NotNull(display.FontFilePath);
        TestAssert.NotNull(body.FontFilePath);
        if (!resolver.GetDiscoveredFonts().Any(f => f.FamilyName.Equals("Aptos Display", StringComparison.OrdinalIgnoreCase)))
        {
            TestAssert.True(
                !display.FamilyName.Equals("Calibri Light", StringComparison.OrdinalIgnoreCase),
                "Expected missing display font resolution to avoid font-name aliases.");
        }

        if (!resolver.GetDiscoveredFonts().Any(f => f.FamilyName.Equals("Aptos", StringComparison.OrdinalIgnoreCase)))
        {
            TestAssert.True(
                !body.FamilyName.Equals("Calibri", StringComparison.OrdinalIgnoreCase),
                "Expected missing body font resolution to avoid font-name aliases.");
        }
    }

    public static void WindowsFontResolverDiscoversMicrosoftCloudFonts()
    {
        string cloudFonts = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft",
            "FontCache",
            "4",
            "CloudFonts");
        if (!Directory.Exists(cloudFonts) ||
            !Directory.EnumerateFiles(cloudFonts, "*.ttf", SearchOption.AllDirectories).Any())
        {
            return;
        }

        var resolver = new WindowsFontResolver();
        IReadOnlyList<FontResolution> fonts = resolver.GetDiscoveredFonts();

        TestAssert.True(
            fonts.Any(f => f.FontFilePath is not null && f.FontFilePath.StartsWith(cloudFonts, StringComparison.OrdinalIgnoreCase)),
            "Expected the default Windows resolver to scan Microsoft cloud font cache directories.");
    }

    public static void PdfEmbeddedFontWidthsCoverEncodedGlyphs()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambriaCollection = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambriaCollection))
        {
            return;
        }

        byte[] bytes = File.ReadAllBytes(cambriaCollection);
        if (OpenTypeFont.GetCollectionFontCount(bytes) < 2)
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(bytes, 1);
        ushort glyph = font.MapCodePoint('h');
        if (glyph == 0)
        {
            return;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "The scale".EnumerateRunes().Select(rune => rune.Value));
        string positioning = TestAssert.NotNull(embedded.EncodeGlyphPositioningArray("The scale", 0d, 18d, forcePositioningArray: true));
        string widths = embedded.BuildWidthArray();

        TestAssert.Contains($"<{glyph:X4}>", positioning);
        TestAssert.Contains(glyph.ToString(CultureInfo.InvariantCulture) + " [", widths);
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

    public static void OpenTypeParserLoadsSpecificTrueTypeCollectionFace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambriaCollection = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambriaCollection))
        {
            return;
        }

        byte[] bytes = File.ReadAllBytes(cambriaCollection);
        if (OpenTypeFont.GetCollectionFontCount(bytes) < 2)
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(bytes, 1);

        TestAssert.Equal("Cambria Math", font.FamilyName);
        TestAssert.True(font.GlyphCount > 0, "Expected glyphs from the selected TTC face.");
    }
}
