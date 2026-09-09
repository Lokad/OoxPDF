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
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));

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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('A');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('O');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('é');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
        }

        TestAssert.True(font.TryReadGlyphOutline(glyph, out var outline), "Expected a readable TrueType outline for Arial 'é'.");
        if (!outline.IsCompound)
        {
            TestAssert.Skip("Environmental precondition not met: (!outline.IsCompound)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));

        TestAssert.True(!font.TryReadGlyphOutline(ushort.MaxValue, out _), "Expected out-of-range glyph outline reads to fail.");
    }

    public static void PdfGlyphOutlinePathConvertsSimpleGlyphContours()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('A');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('O');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort glyph = font.MapCodePoint('é');
        if (glyph == 0 || !font.TryReadGlyphOutline(glyph, out var outline) || !outline.IsCompound)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0 || !font.TryReadGlyphOutline(glyph, out var outline) || !outline.IsCompound)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        if (!font.TableTags.Contains("GPOS"))
        {
            TestAssert.Skip("Environmental precondition not met: (!font.TableTags.Contains(\"GPOS\"))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(cambria));
        AssertNoKerning(font, 'D', 'é');
    }

    public static void OpenTypeParserReadsGposExtensionKerning()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cambria = Path.Combine(fontsDirectory, "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(cambria));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(calibri));
        AssertNoKerning(font, (char)34, (char)44);
        AssertHasKerning(font, (char)84, (char)111);
    }

    public static void OpenTypeParserMapsWindowsSymbolCmap()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string symbol = Path.Combine(fontsDirectory, "symbol.ttf");
        if (!File.Exists(symbol))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(symbol))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(symbol));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri) || !File.Exists(symbol))");
        }
        OpenTypeFont primary = OpenTypeFont.Load(File.ReadAllBytes(calibri));
        OpenTypeFont fallback = OpenTypeFont.Load(File.ReadAllBytes(symbol));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri))");
        }
        OpenTypeFont primary = OpenTypeFont.Load(File.ReadAllBytes(calibri));
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
            TestAssert.Skip("Environmental precondition not met: (primaryResolution.IsFallback || symbolResolution.IsFallback)");
        }

        OpenTypeFont? primary = FontProgramLoader.Load(primaryResolution, CancellationToken.None);
        OpenTypeFont? symbol = FontProgramLoader.Load(symbolResolution, CancellationToken.None);
        if (primary is null || symbol is null || primary.MapCodePoint(0xF0B7) != 0 || symbol.MapCodePoint(0xF0B7) == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (primary is null || symbol is null || primary.MapCodePoint(0xF0B7) != 0 || symbol.MapCodePoint(0xF0B7) == 0)");
        }

        string text = ((char)0xF0B7).ToString();
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(text, [primary, symbol], CancellationToken.None);
        TestAssert.Equal(1, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(1, 0, 1), spans[0]);
    }

    public static void SyntheticCoverageFallbackSplitsByGlyphCoverage()
    {
        OpenTypeFont primary = LoadSubsetFont(new[] { 65, 66 });
        OpenTypeFont fallback = TestFontBuilder.LoadTestFont();
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage("ACB", [primary, fallback], CancellationToken.None);
        TestAssert.Equal(3, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(0, 0, 1), spans[0]);
        TestAssert.Equal(new FontCoverageSpan(1, 1, 1), spans[1]);
        TestAssert.Equal(new FontCoverageSpan(0, 2, 1), spans[2]);
    }

    public static void SyntheticCoverageFallbackReportsUncoveredRunes()
    {
        OpenTypeFont primary = TestFontBuilder.LoadTestFont();
        string text = "A" + (char)160 + "B";
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(text, [primary], CancellationToken.None);
        TestAssert.Equal(3, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(0, 0, 1), spans[0]);
        TestAssert.Equal(new FontCoverageSpan(-1, 1, 1), spans[1]);
        TestAssert.Equal(new FontCoverageSpan(0, 2, 1), spans[2]);
    }

    public static void SyntheticCoverageFallbackReportsSingleFallbackSpan()
    {
        OpenTypeFont primary = LoadSubsetFont(new[] { 65 });
        OpenTypeFont fallback = TestFontBuilder.LoadTestFont();
        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(((char)0x301).ToString(), [primary, fallback], CancellationToken.None);
        TestAssert.Equal(1, spans.Count);
        TestAssert.Equal(new FontCoverageSpan(1, 0, 1), spans[0]);
    }

    private static byte[] CreateSubsetFontBytes(int[] codePoints)
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, codePoints, CancellationToken.None);
        return embedded.FontProgramBytes.ToArray();
    }

    private static OpenTypeFont LoadSubsetFont(int[] codePoints)
    {
        return OpenTypeFont.Load(CreateSubsetFontBytes(codePoints));
    }

    public static void SyntheticDocxMeasurerFallsBackPerCharacterForMissingGlyphs()
    {
        byte[] primaryBytes = CreateSubsetFontBytes(new[] { 65, 66 });
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        var resolver = new FamilyFontResolver(
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["Primary"] = primaryBytes,
                ["Fallback"] = fallbackBytes,
            },
            "Fallback");
        var run = new DocxTextRun("ACB", 12d, null, false, false, false, null, "Primary")
        {
            Fonts = new DocxRunFonts("Primary", null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        DocxTextRun planRun = plan.Runs.Single(candidate => candidate.Run.Text == "ACB").Run;
        var measurer = new DocxFontPlanTextMeasurer(plan, null, CancellationToken.None, resolver);
        double actual = measurer.MeasureText(planRun, "ACB", 9d);
        OpenTypeFont primary = OpenTypeFont.Load(primaryBytes);
        OpenTypeFont fallback = OpenTypeFont.Load(fallbackBytes);
        double expected =
            (primary.GetAdvanceWidth(primary.MapCodePoint(65)) +
            fallback.GetAdvanceWidth(fallback.MapCodePoint(67)) +
            primary.GetAdvanceWidth(primary.MapCodePoint(66))) * 9d / TestFontBuilder.UnitsPerEm;
        TestAssert.True(Math.Abs(actual - expected) < 0.000001d, "Expected per-character fallback to measure the missing glyph in the fallback face.");
    }

    private sealed class FamilyFontResolver(IReadOnlyDictionary<string, byte[]> fontsByFamily, string fallbackFamily) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            if (fontsByFamily.TryGetValue(request.FamilyName, out byte[]? bytes))
            {
                return new FontFaceResolution(
                    request.FamilyName,
                    request.FamilyName,
                    new FontStyleKey(request.Bold, request.Italic, 400, 0, false),
                    new MemoryFontProgramSource("test:" + request.FamilyName, bytes),
                    false);
            }

            return new FontFaceResolution(
                request.FamilyName,
                fallbackFamily,
                new FontStyleKey(request.Bold, request.Italic, 400, 0, false),
                new MemoryFontProgramSource("test:" + fallbackFamily, fontsByFamily[fallbackFamily]),
                true);
        }
    }

    public static void DocxMeasurerFallsBackPerCharacterForMissingGlyphs()
    {
        var resolver = new WindowsFontResolver();
        FontFaceResolution symbolResolution = resolver.Resolve(new FontRequest("Symbol"));
        if (symbolResolution.IsFallback)
        {
            TestAssert.Skip("Environmental precondition not met: (symbolResolution.IsFallback)");
        }

        OpenTypeFont? symbolFont = FontProgramLoader.Load(symbolResolution, CancellationToken.None);
        if (symbolFont is null || symbolFont.MapCodePoint(0xF0B7) == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (symbolFont is null || symbolFont.MapCodePoint(0xF0B7) == 0)");
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
        if (resolved.Resolution is not FontFaceResolution)
        {
            TestAssert.Skip("Environmental precondition not met: (resolved.Resolution is not FontFaceResolution primaryResolution)");
        }

        OpenTypeFont? primaryFont = FontProgramLoader.Load(resolved.Resolution as FontFaceResolution, CancellationToken.None);
        if (primaryFont is null || primaryFont.MapCodePoint(0xF0B7) != 0)
        {
            TestAssert.Skip("Environmental precondition not met: (primaryFont is null || primaryFont.MapCodePoint(0xF0B7) != 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory) || !Directory.EnumerateFiles(fontsDirectory, \"*.ttc\", SearchOption.TopDirectoryOnly).Any())");
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontFaceResolution? mathFace = resolver.GetDiscoveredFonts()
            .FirstOrDefault(f => f.HasMathTable &&
                f.Source is FileFontProgramSource source &&
                source.Path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase));
        if (mathFace is null)
        {
            TestAssert.Skip("Environmental precondition not met: (mathFace is null)");
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
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory) || !Directory.EnumerateFiles(fontsDirectory, \"*.ttc\", SearchOption.TopDirectoryOnly).Any())");
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
            TestAssert.Skip("Environmental precondition not met: (mathFace is null)");
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
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory))");
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
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(cloudFonts) || !Directory.EnumerateFiles(cloudFonts, \"*.ttf\", SearchOption.AllDirectories).Any())");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambriaCollection))");
        }

        byte[] bytes = File.ReadAllBytes(cambriaCollection);
        if (OpenTypeFont.GetCollectionFontCount(bytes) < 2)
        {
            TestAssert.Skip("Environmental precondition not met: (OpenTypeFont.GetCollectionFontCount(bytes) < 2)");
        }

        OpenTypeFont font = OpenTypeFont.Load(bytes, 1);
        ushort glyph = font.MapCodePoint('h');
        if (glyph == 0)
        {
            TestAssert.Skip("Environmental precondition not met: (glyph == 0)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort originalGlyph = font.MapCodePoint('é');
        if (originalGlyph == 0 ||
            !font.TryReadGlyphOutline(originalGlyph, out OpenTypeFont.OpenTypeGlyphOutline originalOutline) ||
            !originalOutline.IsCompound)
        {
            TestAssert.Skip("Environmental precondition not met: (originalGlyph == 0 || !font.TryReadGlyphOutline(originalGlyph, out OpenTypeFont.OpenTypeGlyphOutline originalOutline) || !originalOutline.IsCompound)");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambriaCollection))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(cambriaCollection));

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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambriaCollection))");
        }

        byte[] bytes = File.ReadAllBytes(cambriaCollection);
        if (OpenTypeFont.GetCollectionFontCount(bytes) < 2)
        {
            TestAssert.Skip("Environmental precondition not met: (OpenTypeFont.GetCollectionFontCount(bytes) < 2)");
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

    public static void FontPackResolverCoalescesConcurrentFontDownloads()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", fontBytes, "Aptos", "Aptos", "Aptos");
        OoxPdfFontPackResolver resolver = CreateFontPackResolver(manifest, fontBytes, out StubHttpMessageHandler handler, responseDelay: TimeSpan.FromMilliseconds(50));

        FontFaceResolution resolution = resolver.Resolve(new FontRequest("Aptos"));
        Task<ReadOnlyMemory<byte>>[] reads = Enumerable.Range(0, 16)
            .Select(_ => resolution.Source.GetBytesAsync(CancellationToken.None).AsTask())
            .ToArray();
        Task.WaitAll(reads);

        foreach (Task<ReadOnlyMemory<byte>> read in reads)
        {
            TestAssert.True(read.Result.Span.SequenceEqual(fontBytes), "Concurrent reads must all observe the downloaded bytes.");
        }

        TestAssert.Equal(2, handler.TotalRequests);
        TestAssert.Equal(1, handler.Requests.Count(path => path == "ooxpdf-fonts/test-pack/files/aptos.ttf"));
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

    private static OoxPdfFontPackResolver CreateFontPackResolver(byte[] manifest, byte[]? fontBytes, out StubHttpMessageHandler handler, TimeSpan responseDelay = default)
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

        handler = new StubHttpMessageHandler(responses, responseDelay: responseDelay);
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri))");
        }
        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(calibri));
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, new[] { 65, 32, 160, 66 }, CancellationToken.None);
        string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
        TestAssert.Contains("<0020>", cmap);
        TestAssert.DoesNotContain("<00A0>", cmap);
        TestAssert.True(embedded.EncodeGlyphHex(" \u00A0").Length > 0, "Both space variants must still encode (shared subset glyph).");
    }


    public static void PdfEmbeddedFontSubsetTagIsSixUppercaseLetters()
    {
        string first = PdfEmbeddedFont.CreateSubsetTag("ABCDEF12", "0123456789AB");
        string second = PdfEmbeddedFont.CreateSubsetTag("ABCDEF12", "0123456789AB");
        string other = PdfEmbeddedFont.CreateSubsetTag("ABCDEF12", "FEDCBA987654");

        TestAssert.Equal(6, first.Length);
        TestAssert.True(first.All(c => c >= 'A' && c <= 'Z'), "Expected subset tag to use six uppercase ASCII letters, got: " + first);
        TestAssert.Equal(first, second);
        TestAssert.True(first != other, "Differing codepoint sets must produce distinct subset tags.");
    }

    public static void PdfEmbeddedFontSanitizePreservesSubsetPlus()
    {
        TestAssert.Equal("ABCDEF+Arial", PdfEmbeddedFont.SanitizeName("ABCDEF+Arial"));
        TestAssert.Equal("A-B", PdfEmbeddedFont.SanitizeName("A B"));
        TestAssert.Equal("Font", PdfEmbeddedFont.SanitizeName(string.Empty));
    }

    public static void PdfEmbeddedFontSubsetNamesAreDistinctAndDeterministic()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, "ABC".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(font, "XYZ".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont repeat = PdfEmbeddedFont.Create(font, "ABC".Select(c => (int)c), CancellationToken.None);

        TestAssert.True(first.BaseFontName != second.BaseFontName, "Differing subsets of one face must have distinct BaseFont names.");
        TestAssert.Equal(first.BaseFontName, repeat.BaseFontName);
        foreach (string name in new[] { first.BaseFontName, second.BaseFontName })
        {
            TestAssert.True(name.Length >= 8, "Expected subset BaseFont name to carry a tag prefix, got: " + name);
            int plus = name.IndexOf('+');
            TestAssert.True(plus == 6, "Expected six-letter subset tag followed by plus, got: " + name);
            for (int i = 0; i < 6; i++)
            {
                TestAssert.True(name[i] >= 'A' && name[i] <= 'Z', "Expected subset tag to use uppercase ASCII letters, got: " + name);
            }

            TestAssert.True(name.Contains(PdfEmbeddedFont.SanitizeName(font.FamilyName), StringComparison.Ordinal), "Expected subset BaseFont name to retain the family name, got: " + name);
        }

        PdfEmbeddedFont merged = PdfEmbeddedFont.Merge([first, second], CancellationToken.None);
        TestAssert.True(merged.BaseFontName != first.BaseFontName && merged.BaseFontName != second.BaseFontName, "Merged differing subsets must not reuse either input BaseFont name.");
    }

    public static void PdfEmbeddedFontEmptySetUsesWholeFontName()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, [], CancellationToken.None);

        TestAssert.True(!embedded.UsesSubsetFontProgram, "Expected empty codepoint set to fall back to whole-font embedding.");
        TestAssert.Equal(PdfEmbeddedFont.SanitizeName(font.FamilyName), embedded.BaseFontName);
        TestAssert.True(!embedded.BaseFontName.Contains("+", StringComparison.Ordinal), "Whole-font BaseFont name must not carry a subset tag.");
    }

    public static void PdfEmbeddedFontToUnicodeBatchesWithinLimit()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        CheckCMapBatching(font, new[] { 65 });
        CheckCMapBatching(font, Enumerable.Range(0x20, 95).ToArray());
        CheckCMapBatching(font, Enumerable.Range(0x20, 300).ToArray());
        CheckCMapBatching(font, Enumerable.Range(0x20, 600).Concat(new[] { 0x1F600 }).ToArray());

        static void CheckCMapBatching(OpenTypeFont font, int[] codePoints)
        {
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, codePoints, CancellationToken.None);
            string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
            int total = embedded.UnicodeByCid.Count;
            if (total == 0)
            {
                TestAssert.DoesNotContain("beginbfchar", cmap);
                return;
            }

            List<int> blockSizes = new();
            int search = 0;
            while (true)
            {
                int begin = cmap.IndexOf("beginbfchar", search, StringComparison.Ordinal);
                if (begin < 0)
                {
                    break;
                }

                int lineStart = cmap.LastIndexOf('\n', begin);
                lineStart = lineStart < 0 ? 0 : lineStart + 1;
                string countText = cmap.Substring(lineStart, begin - lineStart).Trim();
                int count = int.Parse(countText, CultureInfo.InvariantCulture);
                TestAssert.True(count >= 1 && count <= 100, "Each bfchar block must hold 1..100 mappings, got: " + count);
                blockSizes.Add(count);
                search = begin + "beginbfchar".Length;
            }

            TestAssert.True(blockSizes.Count > 0, "Expected at least one bfchar block for non-empty mapping.");
            TestAssert.Equal(total, blockSizes.Sum());
            if (total > 100)
            {
                TestAssert.True(blockSizes.Count >= 2, "Mappings beyond 100 must span multiple bfchar blocks.");
            }
        }
    }

    public static void PdfEmbeddedFontToUnicodeEncodesSupplementaryAsSurrogatePair()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        int[] candidates = [0x1F600, 0x1D11E, 0x10428, 0x20000];
        foreach (int codePoint in candidates)
        {
            if (font.MapCodePoint(codePoint) == 0)
            {
                continue;
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, new[] { 65, codePoint }, CancellationToken.None);
            string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
            int scalar = codePoint - 0x10000;
            string expected = ((0xD800 + (scalar >> 10)).ToString("X4", CultureInfo.InvariantCulture) + (0xDC00 + (scalar & 0x3FF)).ToString("X4", CultureInfo.InvariantCulture));
            TestAssert.Contains(expected, cmap);
            return;
        }
    }

    public static void PdfEmbeddedFontPreservesZeroWidthMetrics()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        ushort combining = font.MapCodePoint(0x0301);
        if (combining == 0 || font.GetAdvanceWidth(combining) != 0)
        {
            TestAssert.Skip("Environmental precondition not met: (combining == 0 || font.GetAdvanceWidth(combining) != 0)");
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, new[] { 65, 0x0301 }, CancellationToken.None);
        TestAssert.True(embedded.TryGetEncodedCid(combining, out ushort cid), "Expected combining mark to have an encoded CID.");
        string widths = embedded.BuildWidthArray(CancellationToken.None);
        TestAssert.True(widths.Contains(cid.ToString(CultureInfo.InvariantCulture) + " [0]", StringComparison.Ordinal), "Zero-advance glyphs must emit an explicit zero width, got: " + widths);
        TestAssert.True(widths != "[]", "Width array must not be empty when glyphs are encoded.");
    }
    public static void FontPackRejectsOversizedManifest()
    {
        byte[] manifest = new byte[2 * 1024 * 1024];
        var handler = new StubHttpMessageHandler(new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["ooxpdf-fonts/test-pack/manifest.json"] = manifest,
        });
        var httpClient = new HttpClient(handler);
        OoxPdfFontPackException ex = TestAssert.Throws<OoxPdfFontPackException>(() => OoxPdfFontPackResolver.CreateHttpAsync("test-pack", new Uri("https://example.test/ooxpdf-fonts"), httpClient, CancellationToken.None).GetAwaiter().GetResult());
        TestAssert.Equal(OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed, ex.DiagnosticId);
    }

    public static void FontPackRejectsForgedContentLength()
    {
        byte[] fontBytes = [1, 2, 3, 4];
        byte[] manifest = BuildFontPackManifest("test-pack", "files/aptos.ttf", fontBytes, "Aptos", "Aptos", "Aptos");
        var forged = new ByteArrayContent(manifest);
        forged.Headers.ContentLength = 100_000_000;
        var handler = new StubHttpMessageHandler(
            new Dictionary<string, byte[]>(StringComparer.Ordinal),
            new Dictionary<string, HttpContent>(StringComparer.Ordinal)
            {
                ["ooxpdf-fonts/test-pack/manifest.json"] = forged,
            });
        var httpClient = new HttpClient(handler);
        OoxPdfFontPackException ex = TestAssert.Throws<OoxPdfFontPackException>(() => OoxPdfFontPackResolver.CreateHttpAsync("test-pack", new Uri("https://example.test/ooxpdf-fonts"), httpClient, CancellationToken.None).GetAwaiter().GetResult());
        TestAssert.Equal(OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed, ex.DiagnosticId);
    }

    public static void FontPackCopyEnforcesByteBudget()
    {
        using var exact = new MemoryStream(new byte[] { 1, 2, 3 });
        byte[] fitting = OoxPdfFontPackResolver.CopyCappedAsync(exact, 3, "probe", "https://example.test/x", CancellationToken.None).GetAwaiter().GetResult();
        TestAssert.Equal(3, fitting.Length);
        using var over = new MemoryStream(new byte[] { 1, 2, 3, 4 });
        TestAssert.Throws<OoxPdfFontPackException>(() => OoxPdfFontPackResolver.CopyCappedAsync(over, 3, "probe", "https://example.test/x", CancellationToken.None).GetAwaiter().GetResult());
    }

    private sealed class StubHttpMessageHandler(IReadOnlyDictionary<string, byte[]> responses, IReadOnlyDictionary<string, HttpContent>? rawContents = null, TimeSpan responseDelay = default) : HttpMessageHandler
    {
        public List<string> Requests { get; } = [];

        public int TotalRequests => totalRequests;

        private readonly object requestLock = new();
        private int totalRequests;

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (responseDelay > TimeSpan.Zero)
            {
                await Task.Delay(responseDelay, cancellationToken).ConfigureAwait(false);
            }

            string path = request.RequestUri?.AbsolutePath.TrimStart('/') ?? "";
            lock (requestLock)
            {
                Requests.Add(path);
            }

            Interlocked.Increment(ref totalRequests);
            if (rawContents is not null && rawContents.TryGetValue(path, out HttpContent? raw))
            {
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = raw
                };
            }


            if (!responses.TryGetValue(path, out byte[]? bytes))
            {
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                    Content = new ByteArrayContent(bytes)
                };
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
    public static void FileFontProgramSourceConcurrentReadsReturnSameBytes()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        string path = Path.Combine(Path.GetTempPath(), "ooxpdf-concurrent-" + Guid.NewGuid().ToString("N") + ".ttf");
        File.WriteAllBytes(path, bytes);
        try
        {
            var source = new FileFontProgramSource(path);
            Task<ReadOnlyMemory<byte>>[] reads = Enumerable.Range(0, 16)
                .Select(_ => source.GetBytesAsync(CancellationToken.None).AsTask())
                .ToArray();
            Task.WaitAll(reads);
            foreach (Task<ReadOnlyMemory<byte>> read in reads)
            {
                TestAssert.True(read.Result.Span.SequenceEqual(bytes), "Concurrent first-reads must all observe the same bytes.");
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static void FileFontProgramSourcePrecancelledTokenThrowsBeforeRead()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        string path = Path.Combine(Path.GetTempPath(), "ooxpdf-cancelled-" + Guid.NewGuid().ToString("N") + ".ttf");
        File.WriteAllBytes(path, bytes);
        try
        {
            var source = new FileFontProgramSource(path);
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => source.GetBytesAsync(cts.Token).AsTask().GetAwaiter().GetResult());
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static void DiscoveryHeadersMatchFullLoadOnSyntheticFonts()
    {
        CheckDiscoveryHeaders(TestFontBuilder.CreateTestFont(), 0);
        CheckDiscoveryHeaders(TestFontBuilder.CreateCffKindFont(), 0);
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        CheckDiscoveryHeaders(collection, 0);
        CheckDiscoveryHeaders(collection, 1);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.ReadDiscoveryHeaders(new byte[20], 0));
    }

    public static void DiscoveryHeadersMatchFullLoadAcrossWindowsFonts()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory))");
        }

        int checkedFaces = 0;
        string[] paths = Directory.EnumerateFiles(fontsDirectory)
            .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase)
                || p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase)
                || p.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        foreach (string path in paths)
        {
            byte[] bytes = File.ReadAllBytes(path);
            int faceCount;
            try
            {
                faceCount = OpenTypeFont.GetCollectionFontCount(bytes);
            }
            catch (InvalidDataException)
            {
                continue;
            }

            // Cap mirrors the collection header limit; real files stay far below it.
            for (int faceIndex = 0; faceIndex < faceCount && faceIndex < 256; faceIndex++)
            {
                CheckDiscoveryOutcome(bytes, faceIndex, path);
                checkedFaces++;
            }
        }

        TestAssert.True(checkedFaces > 0, "Expected to check at least one font face.");
    }

    private static void CheckDiscoveryHeaders(byte[] bytes, int fontIndex)
    {
        OpenTypeFont font = OpenTypeFont.Load(bytes, fontIndex);
        OpenTypeFont.FontDiscoveryHeaders headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, fontIndex);
        TestAssert.Equal(font.FamilyName, headers.FamilyName);
        TestAssert.Equal(font.Os2.WeightClass, headers.WeightClass);
        TestAssert.True(Math.Abs(font.Post.ItalicAngle - headers.ItalicAngle) < 1e-9, "Italic angle must match.");
        TestAssert.Equal(font.TableTags.Contains("MATH"), headers.HasMathTable);
    }

    private static void CheckDiscoveryOutcome(byte[] bytes, int fontIndex, string path)
    {
        OpenTypeFont? full = null;
        bool fullFailed = false;
        try
        {
            full = OpenTypeFont.Load(bytes, fontIndex);
        }
        catch (InvalidDataException)
        {
            fullFailed = true;
        }

        OpenTypeFont.FontDiscoveryHeaders? light = null;
        bool lightFailed = false;
        try
        {
            light = OpenTypeFont.ReadDiscoveryHeaders(bytes, fontIndex);
        }
        catch (InvalidDataException)
        {
            lightFailed = true;
        }

        TestAssert.Equal(fullFailed, lightFailed);
        if (fullFailed || full is null || light is null)
        {
            return;
        }

        TestAssert.Equal(full.FamilyName, light.Value.FamilyName);
        TestAssert.Equal(full.Os2.WeightClass, light.Value.WeightClass);
        TestAssert.True(Math.Abs(full.Post.ItalicAngle - light.Value.ItalicAngle) < 1e-9, "Italic angle must match.");
        TestAssert.Equal(full.TableTags.Contains("MATH"), light.Value.HasMathTable);
    }
}