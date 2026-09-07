using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Docx;
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

        FontFaceResolution bold = resolver.Resolve(new FontRequest("Arial", true, false));
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
                WeightClass: font.Os2.WeightClass, FaceIndex: 0,
                HasMathTable: font.TableTags.Contains("MATH")),
            new MemoryFontProgramSource("memory:test-arial", bytes),
            IsFallback: false);

        var resolver = new PresentationFontResolver(new SingleFontResolver(resolution));
        (FontFaceResolution Resolution, OpenTypeFont Font)? resolved = resolver.ResolvePresentationOpenTypeFont(new FontRequest(font.FamilyName), CancellationToken.None);

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
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 10d, 20d, 12d, 0d), "Expected PDF glyph path conversion for Arial 'A'.");
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
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 0d, 0d, 18d, 0d), "Expected PDF glyph path conversion for Arial 'O'.");
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
        TestAssert.True(PdfGlyphOutlinePath.TryAppendGlyphPath(graphics, font, glyph, 0d, 0d, 14d, 0d), "Expected PDF glyph path conversion for compound Arial 'é'.");
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

    public static void OpenTypeParserIgnoresNonLatinScriptGposKerning()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string calibri = Path.Combine(fontsDirectory, "calibri.ttf");
        if (!File.Exists(calibri))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(calibri);
        AssertNoKerning(font, (char)34, (char)44);
        AssertHasKerning(font, (char)84, (char)111);
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

    public static void FontCoverageFallbackSplitsByGlyphCoverage()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string calibri = Path.Combine(fontsDirectory, "calibri.ttf");
        string symbol = Path.Combine(fontsDirectory, "symbol.ttf");
        if (!File.Exists(calibri) || !File.Exists(symbol))
        {
            return;
        }
        OpenTypeFont primary = OpenTypeFont.Load(calibri);
        OpenTypeFont fallback = OpenTypeFont.Load(symbol);
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage("a" + (char)0xF0B7 + "b", [primary, fallback], CancellationToken.None);
        TestAssert.Equal(3, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(0, 0, 1), spans[0]);
        TestAssert.Equal(new FontCoverageSpan(1, 1, 1), spans[1]);
        TestAssert.Equal(new FontCoverageSpan(0, 2, 1), spans[2]);
    }
    public static void FontCoverageFallbackReportsUncoveredRunes()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string calibri = Path.Combine(fontsDirectory, "calibri.ttf");
        if (!File.Exists(calibri))
        {
            return;
        }
        OpenTypeFont primary = OpenTypeFont.Load(calibri);
        string text = "a" + char.ConvertFromUtf32(0x1F600) + "b";
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(text, [primary], CancellationToken.None);
        TestAssert.Equal(3, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(0, 0, 1), spans[0]);
        TestAssert.Equal(new FontCoverageSpan(-1, 1, 2), spans[1]);
        TestAssert.Equal(new FontCoverageSpan(0, 3, 1), spans[2]);
    }

    public static void FontCoverageFallbackReportsSingleFallbackSpan()
    {
        var resolver = new WindowsFontResolver();
        FontFaceResolution primaryResolution = resolver.Resolve(new FontRequest("Aptos"));
        FontFaceResolution symbolResolution = resolver.Resolve(new FontRequest("Symbol"));
        if (primaryResolution.IsFallback || symbolResolution.IsFallback)
        {
            return;
        }

        OpenTypeFont? primary = FontProgramLoader.Load(primaryResolution, CancellationToken.None);
        OpenTypeFont? symbol = FontProgramLoader.Load(symbolResolution, CancellationToken.None);
        if (primary is null || symbol is null || primary.MapCodePoint(0xF0B7) != 0 || symbol.MapCodePoint(0xF0B7) == 0)
        {
            return;
        }

        string text = ((char)0xF0B7).ToString();
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(text, [primary, symbol], CancellationToken.None);
        TestAssert.Equal(1, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(1, 0, 1), spans[0]);
    }

    public static void DocxMeasurerFallsBackPerCharacterForMissingGlyphs()
    {
        var resolver = new WindowsFontResolver();
        FontFaceResolution symbolResolution = resolver.Resolve(new FontRequest("Symbol"));
        if (symbolResolution.IsFallback)
        {
            return;
        }

        OpenTypeFont? symbolFont = FontProgramLoader.Load(symbolResolution, CancellationToken.None);
        if (symbolFont is null || symbolFont.MapCodePoint(0xF0B7) == 0)
        {
            return;
        }

        string body = """
            <?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">a[BULLET]b</w:t></w:r></w:p>
                <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
              </w:body>
            </w:document>
            """.Replace("[BULLET]", ((char)0xF0B7).ToString());
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = body
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        DocxResolvedRunTypeface resolved = plan.Runs.Single(run => run.Run.Text.IndexOf((char)0xF0B7) >= 0);
        if (resolved.Resolution is not FontFaceResolution primaryResolution)
        {
            return;
        }

        OpenTypeFont? primaryFont = FontProgramLoader.Load(primaryResolution, CancellationToken.None);
        if (primaryFont is null || primaryFont.MapCodePoint(0xF0B7) != 0)
        {
            return;
        }

        string text = "a" + (char)0xF0B7 + "b";
        var measurer = new DocxFontPlanTextMeasurer(plan, null, CancellationToken.None, resolver);
        double actual = measurer.MeasureText(resolved.Run, text, 9d);
        double expected =
            primaryFont.GetAdvanceWidth(primaryFont.MapCodePoint(97)) * 9d / primaryFont.UnitsPerEm +
            symbolFont.GetAdvanceWidth(symbolFont.MapCodePoint(0xF0B7)) * 9d / symbolFont.UnitsPerEm +
            primaryFont.GetAdvanceWidth(primaryFont.MapCodePoint(98)) * 9d / primaryFont.UnitsPerEm;
        TestAssert.True(Math.Abs(actual - expected) < 0.000001d, "Expected per-character fallback to measure the bullet in the Symbol face.");
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

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "The scale".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);
        string positioning = TestAssert.NotNull(embedded.EncodeGlyphPositioningArray("The scale", 0d, 18d, forcePositioningArray: true, kerningEnabled: true));
        string widths = embedded.BuildWidthArray(CancellationToken.None);
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c), CancellationToken.None);

        TestAssert.True(embedded.UsesSubsetFontProgram, "Expected PDF font embedding to use a subset font program.");
        TestAssert.True(embedded.FontProgramBytes.Length < font.Bytes.Length / 4, "Expected a tiny two-glyph subset compared to the source font.");
        TestAssert.Equal("0001", embedded.EncodeGlyphHex("A"));
        TestAssert.Equal("0002", embedded.EncodeGlyphHex("z"));
        TestAssert.Contains("1 [", embedded.BuildWidthArray(CancellationToken.None));
        TestAssert.Contains("2 [", embedded.BuildWidthArray(CancellationToken.None));

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

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "é".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);
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

    public static void FontPackResolverDownloadsManifestOnlyOnCreate()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", fontBytes, "Aptos", "Aptos", "Aptos");

        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, fontBytes: null, out StubHttpMessageHandler handler);

        TestAssert.Equal("test-pack", resolver.PackId);
        TestAssert.Equal("https://example.test/ooxpdf-fonts/test-pack/", resolver.PackRootUri.AbsoluteUri);
        TestAssert.Equal(1, handler.Requests.Count);
        TestAssert.Equal("ooxpdf-fonts/test-pack/manifest.json", handler.Requests[0]);
    }

    public static void FontPackResolverDownloadsSelectedFontAndValidatesHash()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", fontBytes, "Aptos", "Aptos", "Aptos");
        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, fontBytes, out StubHttpMessageHandler handler);

        FontFaceResolution resolution = resolver.Resolve(new FontRequest("Aptos"));
        ReadOnlyMemory<byte> loaded = resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();
        ReadOnlyMemory<byte> cached = resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();

        TestAssert.Equal("Aptos", resolution.RequestedFamily);
        TestAssert.Equal("Aptos", resolution.ResolvedFamily);
        TestAssert.True(!resolution.IsFallback, "Expected exact font pack resolution not to be marked as fallback.");
        TestAssert.True(loaded.ToArray().SequenceEqual(fontBytes), "Expected font pack source to return downloaded font bytes.");
        TestAssert.True(cached.ToArray().SequenceEqual(fontBytes), "Expected cached font pack bytes to remain stable.");
        TestAssert.Equal(2, handler.Requests.Count);
        TestAssert.Equal("ooxpdf-fonts/test-pack/files/aptos.ttf", handler.Requests[1]);
    }

    public static void FontPackResolverRejectsHashMismatch()
    {
        byte[] expectedFontBytes = [1, 2, 3, 4];
        byte[] servedFontBytes = [1, 2, 3, 5];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", expectedFontBytes, "Aptos", "Aptos", "Aptos");
        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, servedFontBytes, out _);
        FontFaceResolution resolution = resolver.Resolve(new FontRequest("Aptos"));

        OoxPdfFontPackException ex = TestAssert.Throws<OoxPdfFontPackException>(
            () => resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult());

        TestAssert.Equal(OoxPdfFontPackDiagnosticIds.FontPackHashMismatch, ex.DiagnosticId);
    }

    public static void FontPackResolverRejectsUnsafeRelativePath()
    {
        byte[] manifest = Encoding.UTF8.GetBytes("""
            {
              "packId": "test-pack",
              "files": [
                { "relativePath": "%2e%2e/aptos.ttf", "byteSize": 4, "sha256": "0000000000000000000000000000000000000000000000000000000000000000" }
              ],
              "families": [
                { "requestedFamily": "Aptos", "resolvedFamily": "Aptos", "relativeFontFile": "%2e%2e/aptos.ttf" }
              ]
            }
            """);

        OoxPdfFontPackException ex = TestAssert.Throws<OoxPdfFontPackException>(
            () => CreateFontPackResolver(manifest, fontBytes: null, out _));

        TestAssert.Equal(OoxPdfFontPackDiagnosticIds.FontPackInvalid, ex.DiagnosticId);
    }

    public static void FontPackResolverUsesConfiguredFallbackFamilies()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest(
            "test-pack",
            "files/fallback.ttf",
            fontBytes,
            requestedFamily: "Fallback Sans",
            resolvedFamily: "Fallback Sans",
            fallbackFamily: "Fallback Sans");
        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, fontBytes, out _);

        FontFaceResolution resolution = resolver.Resolve(new FontRequest("Missing Sans"));

        TestAssert.Equal("Missing Sans", resolution.RequestedFamily);
        TestAssert.Equal("Fallback Sans", resolution.ResolvedFamily);
        TestAssert.True(resolution.IsFallback, "Expected missing font requests to use the configured font pack fallback.");
    }

    public static void PresentationFontResolverUsesFontPackCatalog()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", fontBytes, "Aptos", "Aptos", "Aptos");
        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, fontBytes, out _);
        var presentationResolver = new PresentationFontResolver(resolver);

        IReadOnlyList<FontFaceResolution> fonts = presentationResolver.GetDiscoveredFonts();

        TestAssert.Equal(1, fonts.Count);
        TestAssert.Equal("Aptos", fonts[0].ResolvedFamily);
        TestAssert.True(
            fonts[0].Source.StableId.StartsWith("ooxpdf-font-pack:test-pack:", StringComparison.Ordinal),
            "Expected presentation font discovery to expose font pack sources.");
    }

    private static OoxPdfFontPackResolver CreateFontPackResolver(byte[] manifest, byte[]? fontBytes, out StubHttpMessageHandler handler)
    {
        var responses = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["ooxpdf-fonts/test-pack/manifest.json"] = manifest
        };
        if (fontBytes is not null)
        {
            responses["ooxpdf-fonts/test-pack/files/aptos.ttf"] = fontBytes;
            responses["ooxpdf-fonts/test-pack/files/fallback.ttf"] = fontBytes;
        }

        handler = new StubHttpMessageHandler(responses);
        var httpClient = new HttpClient(handler);
        return OoxPdfFontPackResolver.CreateHttpAsync(
            "test-pack",
            new Uri("https://example.test/ooxpdf-fonts"),
            httpClient,
            CancellationToken.None).GetAwaiter().GetResult();
    }

    private static byte[] BuildFontPackManifest(
        string packId,
        string relativeFontPath,
        byte[] fontBytes,
        string requestedFamily,
        string resolvedFamily,
        string fallbackFamily)
    {
        string sha256 = Convert.ToHexString(SHA256.HashData(fontBytes));
        return Encoding.UTF8.GetBytes($$"""
            {
              "packId": "{{packId}}",
              "files": [
                { "relativePath": "{{relativeFontPath}}", "byteSize": {{fontBytes.Length}}, "sha256": "{{sha256}}" }
              ],
              "families": [
                {
                  "requestedFamily": "{{requestedFamily}}",
                  "resolvedFamily": "{{resolvedFamily}}",
                  "relativeFontFile": "{{relativeFontPath}}",
                  "weight": 400,
                  "italic": false,
                  "faceIndex": 0,
                  "hasMathTable": false
                }
              ],
              "fallbacks": [
                { "family": "{{fallbackFamily}}" }
              ]
            }
            """);
    }


    public static void PdfEmbeddedFontSubsetsWithDifferentCodepointsDoNotShareResourceKey()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, "ABC".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(font, "XYZ".Select(c => (int)c), CancellationToken.None);

        TestAssert.True(first.ResourceKey != second.ResourceKey, "Subsets over different codepoint sets must not merge: subset CID assignment is only valid for its own set.");
    }

    public static void PdfEmbeddedFontSubsetsWithSameCodepointsShareResourceKey()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, "ABCDEF".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(font, "ABCDEF".Select(c => (int)c), CancellationToken.None);

        TestAssert.Equal(first.ResourceKey, second.ResourceKey);
        PdfEmbeddedFont merged = PdfEmbeddedFont.Merge([first, second], CancellationToken.None);
        TestAssert.Equal(first.ResourceKey, merged.ResourceKey);
        TestAssert.Equal(first.EncodeGlyphHex("ABCDEF"), merged.EncodeGlyphHex("ABCDEF"));
        TestAssert.Equal(first.BuildWidthArray(CancellationToken.None), merged.BuildWidthArray(CancellationToken.None));
    }
    public static void PdfEmbeddedFontMapsSharedSpaceGlyphToPlainSpace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string calibri = Path.Combine(fontsDirectory, "calibri.ttf");
        if (!File.Exists(calibri))
        {
            return;
        }
        OpenTypeFont font = OpenTypeFont.Load(calibri);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, new[] { 65, 32, 160, 66 }, CancellationToken.None);
        string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
        TestAssert.Contains("<0020>", cmap);
        TestAssert.DoesNotContain("<00A0>", cmap);
        TestAssert.True(embedded.EncodeGlyphHex(" \u00A0").Length > 0, "Both space variants must still encode (shared subset glyph).");
    }


    private sealed class StubHttpMessageHandler(IReadOnlyDictionary<string, byte[]> responses) : HttpMessageHandler
    {
        public List<string> Requests { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string path = request.RequestUri?.AbsolutePath.TrimStart('/') ?? "";
            Requests.Add(path);
            if (!responses.TryGetValue(path, out byte[]? bytes))
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            });
        }
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