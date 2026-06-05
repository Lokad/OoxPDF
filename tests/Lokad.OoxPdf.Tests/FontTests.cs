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
        IReadOnlyList<FontFaceResolution> fonts = resolver.GetDiscoveredFonts();
        if (fonts.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one discoverable Windows font.");
        }

        FontFaceResolution resolved = resolver.Resolve(new FontRequest("Arial"));
        TestAssert.NotNull(resolved.Source);

        FontFaceResolution bold = resolver.Resolve(new FontRequest("Arial", Bold: true));
        TestAssert.NotNull(bold.Source);
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
        TestAssert.True(font.Os2.StrikeoutSize > 0, "Expected a positive OS/2 strikeout size.");
        TestAssert.True(font.Os2.StrikeoutPosition > 0, "Expected a positive OS/2 strikeout position.");
        TestAssert.True(font.Post.UnderlineThickness != 0, "Expected a non-zero post underline thickness.");
        TestAssert.True(font.TableTags.Contains("cmap"), "Expected cmap table.");
        ushort glyph = font.MapCodePoint('A');
        TestAssert.True(glyph > 0, "Expected a glyph mapping for 'A'.");
        TestAssert.True(font.GetAdvanceWidth(glyph) > 0, "Expected a positive advance width for 'A'.");
    }

    public static void PresentationFontResolverLoadsMemoryBackedFontSource()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        byte[] bytes = File.ReadAllBytes(arial);
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        var resolution = new FontFaceResolution(
            font.FamilyName,
            font.FamilyName,
            new FontStyleKey(
                Bold: font.Os2.WeightClass >= 600,
                Italic: Math.Abs(font.Post.ItalicAngle) > 0.01d,
                WeightClass: font.Os2.WeightClass,
                HasMathTable: font.TableTags.Contains("MATH")),
            new MemoryFontProgramSource("memory:test-arial", bytes),
            IsFallback: false);

        var resolver = new PresentationFontResolver(new SingleFontResolver(resolution));
        (FontFaceResolution Resolution, OpenTypeFont Font)? resolved = resolver.ResolvePresentationOpenTypeFont(new FontRequest(font.FamilyName));

        if (resolved is null)
        {
            throw new InvalidOperationException("Expected memory-backed font source to load through presentation resolver.");
        }

        TestAssert.Equal("memory:test-arial", resolved.Value.Resolution.Source.StableId);
        TestAssert.Equal(font.FamilyName, resolved.Value.Font.FamilyName);
        TestAssert.True(resolved.Value.Font.MapCodePoint('A') != 0, "Expected memory-backed font to map capital A.");
    }

    public static void OpenTypeParserReadsSimpleGlyphOutlines()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('A');
        if (glyph == 0)
        {
            return;
        }

        TestAssert.True(font.TryReadGlyphOutline(glyph, out var outline), "Expected a readable TrueType outline for Arial 'A'.");
        TestAssert.True(!outline.IsCompound, "Expected Arial 'A' to be a simple glyph outline.");
        TestAssert.True(outline.Contours.Count > 0, "Expected at least one contour in Arial 'A'.");
        TestAssert.True(outline.Contours.Sum(contour => contour.Points.Count) > 0, "Expected outline points in Arial 'A'.");
        TestAssert.True(
            outline.Bounds.XMax > outline.Bounds.XMin && outline.Bounds.YMax > outline.Bounds.YMin,
            "Expected non-empty glyph outline bounds for Arial 'A'.");
    }

    public static void OpenTypeParserReadsCurvedGlyphOutlinePoints()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('O');
        if (glyph == 0)
        {
            return;
        }

        TestAssert.True(font.TryReadGlyphOutline(glyph, out var outline), "Expected a readable TrueType outline for Arial 'O'.");
        TestAssert.True(
            outline.Contours.SelectMany(contour => contour.Points).Any(point => !point.IsOnCurve),
            "Expected at least one off-curve point in Arial 'O'.");
    }

    public static void OpenTypeParserExpandsCompoundGlyphOutlines()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('é');
        if (glyph == 0)
        {
            return;
        }

        TestAssert.True(font.TryReadGlyphOutline(glyph, out var outline), "Expected a readable TrueType outline for Arial 'é'.");
        if (!outline.IsCompound)
        {
            return;
        }

        TestAssert.True(outline.Contours.Count > 1, "Expected expanded contours from an Arial compound glyph.");
        TestAssert.True(outline.Contours.Sum(contour => contour.Points.Count) > 10, "Expected expanded compound glyph points.");
    }

    public static void OpenTypeParserRejectsOutOfRangeGlyphOutline()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);

        TestAssert.True(!font.TryReadGlyphOutline(ushort.MaxValue, out _), "Expected out-of-range glyph outline reads to fail.");
    }

    public static void PdfGlyphOutlinePathConvertsSimpleGlyphContours()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('A');
        if (glyph == 0)
        {
            return;
        }

        var graphics = new PdfGraphicsBuilder();
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 10d, 20d, 12d), "Expected PDF glyph path conversion for Arial 'A'.");
        string pdf = graphics.ToString();

        TestAssert.Contains(" m", pdf);
        TestAssert.Contains(" l", pdf);
        TestAssert.Contains("h", pdf);
        TestAssert.True(!pdf.Contains("NaN", StringComparison.Ordinal), "Expected finite glyph path coordinates.");
    }

    public static void PdfGlyphOutlinePathConvertsQuadraticCurvesToCubics()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('O');
        if (glyph == 0)
        {
            return;
        }

        var graphics = new PdfGraphicsBuilder();
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 0d, 0d, 18d), "Expected PDF glyph path conversion for Arial 'O'.");
        string pdf = graphics.ToString();

        TestAssert.Contains(" c", pdf);
        TestAssert.True(!pdf.Contains("Infinity", StringComparison.Ordinal), "Expected finite cubic control points.");
    }

    public static void PdfGlyphOutlinePathConvertsCompoundGlyphContours()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort glyph = font.MapCodePoint('é');
        if (glyph == 0 || !font.TryReadGlyphOutline(glyph, out var outline) || !outline.IsCompound)
        {
            return;
        }

        var graphics = new PdfGraphicsBuilder();
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 0d, 0d, 14d), "Expected PDF glyph path conversion for compound Arial 'é'.");
        string pdf = graphics.ToString();

        TestAssert.True(pdf.Split(" m", StringSplitOptions.None).Length > 2, "Expected multiple contour starts for compound glyph path.");
        TestAssert.True(!pdf.Contains("NaN", StringComparison.Ordinal), "Expected finite compound glyph path coordinates.");
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
        FontFaceResolution? mathFace = resolver.GetDiscoveredFonts()
            .FirstOrDefault(f => f.HasMathTable &&
                f.Source is FileFontProgramSource source &&
                source.Path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase));
        if (mathFace is null)
        {
            return;
        }

        FontFaceResolution resolved = resolver.Resolve(new FontRequest(mathFace.FamilyName));

        TestAssert.True(resolved.HasMathTable, "Expected exact font resolution to keep the requested math-table face.");
        TestAssert.Equal(mathFace.Source.StableId, resolved.Source.StableId);
        TestAssert.True(!resolved.IsFallback, "Expected exact font resolution not to be marked as fallback.");
    }

    public static void WindowsFontResolverPreservesPresentationMathTextFace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory) ||
            !Directory.EnumerateFiles(fontsDirectory, "*.ttc", SearchOption.TopDirectoryOnly).Any())
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontFaceResolution? mathFace = resolver.GetDiscoveredFonts()
            .FirstOrDefault(f => f.HasMathTable &&
                f.Source is FileFontProgramSource source &&
                source.Path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase) &&
                resolver.GetDiscoveredFonts().Any(other =>
                    !other.HasMathTable &&
                    other.Source.StableId.Equals(f.Source.StableId, StringComparison.OrdinalIgnoreCase)));
        if (mathFace is null)
        {
            return;
        }

        FontFaceResolution resolved = resolver.ResolvePresentationTextFace(new FontRequest(mathFace.FamilyName));

        TestAssert.Equal(mathFace.Source.StableId, resolved.Source.StableId);
        TestAssert.True(resolved.HasMathTable, "Expected PPTX presentation text to preserve the requested math-table face.");
        TestAssert.True(!resolved.IsFallback, "Expected the requested presentation math face not to be marked as fallback.");
    }

    public static void WindowsFontResolverUsesMetadataRatherThanFontNameAliases()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontFaceResolution display = resolver.Resolve(new FontRequest("Aptos Display"));
        FontFaceResolution body = resolver.Resolve(new FontRequest("Aptos"));

        TestAssert.NotNull(display.Source);
        TestAssert.NotNull(body.Source);
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
        IReadOnlyList<FontFaceResolution> fonts = resolver.GetDiscoveredFonts();

        TestAssert.True(
            fonts.Any(f => f.Source is FileFontProgramSource source && source.Path.StartsWith(cloudFonts, StringComparison.OrdinalIgnoreCase)),
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
        string encodedGlyph = embedded.EncodeGlyphHex("h");
        TestAssert.True(encodedGlyph.Length == 4, "Expected a single encoded CID for 'h'.");
        int cid = int.Parse(encodedGlyph, NumberStyles.HexNumber, CultureInfo.InvariantCulture);

        TestAssert.Contains(encodedGlyph, positioning);
        TestAssert.Contains(cid.ToString(CultureInfo.InvariantCulture) + " [", widths);
    }

    public static void PdfEmbeddedFontBuildsLoadableSubsetFontProgram()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c));

        TestAssert.True(embedded.UsesSubsetFontProgram, "Expected PDF font embedding to use a subset font program.");
        TestAssert.True(embedded.FontProgramBytes.Length < font.Bytes.Length / 4, "Expected a tiny two-glyph subset compared to the source font.");
        TestAssert.Equal("0001", embedded.EncodeGlyphHex("A"));
        TestAssert.Equal("0002", embedded.EncodeGlyphHex("z"));
        TestAssert.Contains("1 [", embedded.BuildWidthArray());
        TestAssert.Contains("2 [", embedded.BuildWidthArray());

        OpenTypeFont subset = OpenTypeFont.Load(embedded.FontProgramBytes.ToArray());
        TestAssert.True(subset.GlyphCount < font.GlyphCount, "Expected the subset font to expose fewer glyphs than the source font.");
        TestAssert.True(subset.MapCodePoint('A') != 0, "Expected the subset cmap to map capital A.");
        TestAssert.True(subset.MapCodePoint('z') != 0, "Expected the subset cmap to map lowercase z.");
    }

    public static void PdfEmbeddedFontSubsetKeepsCompoundGlyphComponents()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        ushort originalGlyph = font.MapCodePoint('é');
        if (originalGlyph == 0 ||
            !font.TryReadGlyphOutline(originalGlyph, out OpenTypeFont.OpenTypeGlyphOutline originalOutline) ||
            !originalOutline.IsCompound)
        {
            return;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "é".EnumerateRunes().Select(rune => rune.Value));
        OpenTypeFont subset = OpenTypeFont.Load(embedded.FontProgramBytes.ToArray());
        ushort subsetGlyph = subset.MapCodePoint('é');

        TestAssert.True(embedded.UsesSubsetFontProgram, "Expected compound glyph fixture to use a subset font program.");
        TestAssert.True(subsetGlyph != 0, "Expected the subset cmap to map the compound glyph.");
        TestAssert.True(subset.TryReadGlyphOutline(subsetGlyph, out OpenTypeFont.OpenTypeGlyphOutline subsetOutline), "Expected remapped compound glyph outline to be readable.");
        TestAssert.True(subsetOutline.Contours.Count > 0, "Expected remapped compound glyph to keep component contours.");
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

    private sealed class SingleFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution with
            {
                RequestedFamily = request.FamilyName,
                IsFallback = !request.FamilyName.Equals(resolution.FamilyName, StringComparison.OrdinalIgnoreCase),
                Style = resolution.Style with { Bold = request.Bold, Italic = request.Italic }
            };
        }
    }
}
