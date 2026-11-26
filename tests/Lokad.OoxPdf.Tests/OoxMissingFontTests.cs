using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// RV01: text must not vanish with a successful undiagnosed result when no
// usable embeddable face exists. Empty/broken font sources are
// platform-independent, so these tests always run (never skipped).
internal static class OoxMissingFontTests
{
    public static void MissingFontsPreserveDocxPagesAndText()
    {
        // RV01: three forced pages with an empty font source keep their pages,
        // keep extractable text, and report the fallback (fails: 1 page, no BT).
        string input = WritePageBreakDocx("First", "Second", "Third");
        var resolver = new MapFontResolver([], "Fallback");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, resolver);
        AssertFallbackPages(pdf);
        TestAssert.DoesNotContain("/FontFile2", pdf);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE");
        TestAssert.Equal(OoxPdfSeverity.Warning, missing.Severity);
        string repeat = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, repeat, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver });
        TestAssert.Equal(pdf, File.ReadAllText(repeat, Encoding.Latin1));
    }

    public static void MissingFontsKeepFittedDocxCellGlyphsInsideTheirClip()
    {
        string[] lines = ["Original target words remain clear", "Original-target words remain clear"];
        string rows = string.Concat(lines.Select(line => $$"""
            <w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:tcFitText/></w:tcPr>
              <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{{line}}</w:t></w:r></w:p>
            </w:tc></w:tr>
            """));
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                  <w:tbl><w:tblPr><w:tblW w:w="2880" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
                    <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>{{rows}}</w:tbl>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr>
                </w:body></w:document>
                """
        });

        foreach ((var mode, var geometry) in new[]
        {
            (OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
            (OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
        })
        {
            input.Position = 0;
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Docx, FontResolver = new MissingTestFontResolver(),
                DocxMarkupMode = mode, DocxMarkupGeometryMode = geometry, DiagnosticSink = diagnostics.Add
            });
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "FONT_NO_USABLE_FACE"),
                "The fitted text must use the diagnosed fallback face.");
            string pdf = Encoding.Latin1.GetString(output.ToArray());
            TestAssert.Contains("/Type /Pages /Count 1", pdf);
            TestAssert.True(Regex.Matches(pdf, @"(?m)^72 [-\d.]+ 144 [-\d.]+ re W\*? n").Count >= lines.Length,
                "Each fitted row must retain its authored cell clip.");
            var rowsByBaseline = FallbackFontMetricsTests.ReadGlyphs(pdf).GroupBy(glyph => glyph.Y).ToArray();
            TestAssert.Equal(lines.Length, rowsByBaseline.Length);
            for (int i = 0; i < lines.Length; i++)
            {
                var glyphs = rowsByBaseline[i].ToArray();
                TestAssert.Equal(lines[i], Encoding.ASCII.GetString(glyphs.Select(glyph => glyph.Code).ToArray()));
                TestAssert.True(glyphs.All(glyph => glyph.Size == 12d && glyph.X >= 72d &&
                    glyph.X + PdfFallbackFont.Helvetica.MeasureAdvanceEm(glyph.Code) * .012d <= 216.001d),
                    $"{mode} must paint the entire fitted label '{lines[i]}' inside the 144pt clip.");
                // The final Helvetica r has a 333-unit advance at 12pt.
                TestAssert.True(glyphs[^1].X + 3.996d >= 215d, "The fitted run reaches the right inset.");
            }
        }
    }

    public static void InvalidFontBytesPreserveDocxPagesAndText()
    {
        // RV01: undecodable font programs take the same diagnosed fallback.
        string input = WritePageBreakDocx("First", "Second", "Third");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, new FixedBytesResolver(new byte[] { 1, 2, 3, 4 }));
        AssertFallbackPages(pdf);
        TestAssert.Equal(OoxPdfSeverity.Warning, diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE").Severity);
    }

    public static void UnsupportedOutlineFontsPreserveDocxPagesAndText()
    {
        // RV01: CFF outlines cannot embed, so their runs join the fallback.
        string input = WritePageBreakDocx("First", "Second", "Third");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, new FixedBytesResolver(TestFontBuilder.CreateCffKindFont()));
        AssertFallbackPages(pdf);
        TestAssert.Equal(OoxPdfSeverity.Warning, diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE").Severity);
    }

    private static void AssertFallbackPages(string pdf)
    {
        string[] expected = ["First", "Second", "Third"];
        var tree = Regex.Match(pdf, @"/Type /Pages /Count 3 /Kids \[([^\]]+)\]");
        var pages = Regex.Matches(tree.Groups[1].Value, @"(\d+) 0 R");
        TestAssert.Equal(expected.Length, pages.Count);
        for (int index = 0; index < pages.Count; index++)
        {
            string page = PdfWriterTests.ReadObject(pdf, pages[index].Groups[1].Value);
            TestAssert.Contains("/Font <<", page);
            string contentId = Regex.Match(page, @"/Contents (\d+) 0 R").Groups[1].Value;
            string content = PdfWriterTests.ReadObject(pdf, contentId);
            TestAssert.Equal(expected[index], ExtractWinAnsiText(content).Trim());
        }
    }

    public static void PartiallyCoveredFontsSubstituteMissingGlyphs()
    {
        // RV01: a usable face with an uncovered codepoint keeps the covered text,
        // substitutes a visible marker, and diagnoses the substitution.
        string input = WritePageBreakDocx("A中B", "Second", "Third");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, new FixedBytesResolver(TestFontBuilder.CreateTestFont()));
        TestAssert.Contains("BT", pdf);
        string extracted = ExtractCidMappedText(pdf);
        TestAssert.Contains("A?B", extracted);
        TestAssert.Contains("Second", extracted);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS");
        TestAssert.Equal(OoxPdfSeverity.Warning, missing.Severity);
        TestAssert.Contains("U+4E2D", missing.Message);
    }

    // Unresolvable runes retain a fallback advance and a diagnosed question mark
    // instead of disappearing or overlapping the following text.
    public static void UnresolvableRunesSubstituteQuestionMark()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="FaceA"/></a:rPr><a:t>AĀ中B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
        });
        var resolver = new SingleFamilyResolver("FaceA", TestFontBuilder.CreateTestFont());
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = resolver, DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        string extracted = ExtractCidMappedText(pdf);
        TestAssert.Contains("A??B", extracted);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS");
        TestAssert.Contains("U+0100", missing.Message);
        TestAssert.Contains("U+4E2D", missing.Message);
    }

    private sealed class SingleFamilyResolver(string family, byte[] bytes) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(request.FamilyName, family, new FontStyleKey(request.Bold, request.Italic, 400, 0, false), new MemoryFontProgramSource("test:" + family, bytes), IsFallback: false);
        }
    }

    // RV01 (slide spans path): a CFF run with estimator-unresolvable runes keeps its
    // covered runes through the per-glyph fallback (embedded and CID-extractable) while
    // the uncovered rune renders as a built-in question mark named by FONT_MISSING_GLYPHS.
    // The CFF family itself stays diagnosed (unsupported outlines plus no usable face).
    public static void CffUnresolvableRunesSubstituteQuestionMark()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="CffFamily"/></a:rPr><a:t>A中B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
        });
        var resolver = new CffSingleFallbackResolver(TestFontBuilder.CreateCffKindFont(), TestFontBuilder.CreateTestFont());
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = resolver, DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        string extracted = ExtractCidMappedText(pdf);
        TestAssert.Contains("AB", extracted);
        string fallbackExtracted = ExtractWinAnsiText(pdf);
        TestAssert.Contains("?", fallbackExtracted);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS");
        TestAssert.Contains("U+4E2D", missing.Message);
        TestAssert.Contains("CffFamily", missing.Message);
        TestAssert.True(diagnostics.Any(d => d.Id == "FONT_NO_USABLE_FACE" && d.Message.Contains("CffFamily", StringComparison.Ordinal)), "The CFF family must stay diagnosed.");
        TestAssert.True(diagnostics.Any(d => d.Id == "FONT_UNSUPPORTED_OUTLINES" && d.Message.Contains("CffFamily", StringComparison.Ordinal)), "The CFF substitution must stay diagnosed.");
    }

    private sealed class CffSingleFallbackResolver(byte[] cffBytes, byte[] fallbackBytes) : IFontResolver, IFontCatalog
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            bool cff = request.FamilyName.Equals("CffFamily", StringComparison.OrdinalIgnoreCase);
            return new FontFaceResolution(request.FamilyName, cff ? "CffFamily" : "FallbackFamily", new FontStyleKey(request.Bold, request.Italic, 400, 0, false), new MemoryFontProgramSource(cff ? "cff-test" : "tt-test", cff ? cffBytes : fallbackBytes), IsFallback: false);
        }

        public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
        {
            return [Resolve(new FontRequest("FallbackFamily"))];
        }
    }

    // RV01 (runs-path follow-up): chart runs with estimator-unresolvable runes keep
    // them as default-family runs so emission substitutes a diagnosed question mark.
    public static void CffChartUnresolvableRunesSubstituteQuestionMark()
    {
        using var source = PptxTypographyTests.CreateCffChartPackage("A中B");
        using var package = OoxPackage.Open(source, CancellationToken.None);
        byte[] fallback = File.ReadAllBytes(TestFixtures.PublicCasePath("PortableFonts/DejaVuSans.ttf"));
        var resolver = new CffSingleFallbackResolver(TestFontBuilder.CreateCffKindFont(), fallback);
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new PptxRenderer(resolver)
            .RenderPagesAsync(new PptxReader().Read(package, CancellationToken.None), package, diagnostics.Add, CancellationToken.None)
            .ToBlockingEnumerable().Single();
        // The 72pt slide insets and 432pt by 288pt chart frame bound the title.
        var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, "A?B", fontPrefix: "CT",
            requiredClip: new PdfRectangle(72, 180, 432, 288));
        TestAssert.Equal(3, glyphs.Count);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS");
        TestAssert.Contains("U+4E2D", missing.Message);
    }
    // RV01 (slide spans path): a split family that resolves to an embeddable face only
    // at emission gains an embedded subset instead of vanishing at the font lookup.
    public static void ResolvableSplitFamiliesEmbedInsteadOfVanishing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="RunFam"/></a:rPr><a:t>AB</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
        });
        var resolver = new UnresolvablePrimaryResolver(TestFontBuilder.CreateTestFont());
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = resolver, DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        string extracted = ExtractCidMappedText(pdf);
        TestAssert.Contains("AB", extracted);
    }

    private sealed class UnresolvablePrimaryResolver(byte[] fallbackBytes) : IFontResolver, IFontCatalog
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            bool missing = request.FamilyName.Equals("RunFam", StringComparison.OrdinalIgnoreCase);
            return new FontFaceResolution(request.FamilyName, missing ? "RunFam" : "DiscFam", new FontStyleKey(request.Bold, request.Italic, 400, 0, false), new MemoryFontProgramSource(missing ? "empty-test" : "tt-test", missing ? Array.Empty<byte>() : fallbackBytes), IsFallback: false);
        }

        public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
        {
            return [Resolve(new FontRequest("DiscFam"))];
        }
    }
    public static void MissingFontsPreservePptxText()
    {
        // RV01: the PPTX audit shows the same silent loss; text stays visible and
        // diagnosed with an empty font source (fails: no BT, no diagnostic).
        string input = TestFixtures.PublicCasePath("pptx-ladder-02-plain-text.pptx");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new MapFontResolver([], "Fallback"), DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        TestAssert.Contains("/Type /Pages /Count 1", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains("/Font <<", pdf);
        TestAssert.DoesNotContain("/FontFile2", pdf);
        TestAssert.True(ExtractWinAnsiText(pdf).Length > 0, "Fallback text must stay extractable.");
        TestAssert.Equal(OoxPdfSeverity.Warning, diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE").Severity);
    }

    private static string WritePageBreakDocx(string first, string second, string third)
    {
        // Three forced pages mirroring the review probe document.
        string Body(string text, bool breakBefore)
        {
            string open = breakBefore ? "<w:p><w:pPr><w:pageBreakBefore/></w:pPr>" : "<w:p>";
            return open + "<w:r><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><w:t>" + text + "</w:t></w:r></w:p>";
        }
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>()
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>"),
            ["_rels/.rels"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body>" + Body(first, false) + Body(second, true) + Body(third, true)
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>"),
        });
    }

    private static (string Pdf, List<OoxPdfDiagnostic> Diagnostics) ConvertDocx(string input, IFontResolver resolver)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, DiagnosticSink = diagnostics.Add });
        return (File.ReadAllText(output, Encoding.Latin1), diagnostics);
    }

    internal static string ExtractWinAnsiText(string pdf)
    {
        // The diagnosed fallback emits WinAnsi bytes in plain Tj operators, so
        // Latin-1 decoding recovers the preserved text deterministically.
        var builder = new StringBuilder();
        foreach (Match match in Regex.Matches(pdf, "<([0-9A-Fa-f]{2,})>\\s*Tj"))
        {
            int chars = match.Groups[1].Value.Length / 2;
            var bytes = new byte[chars];
            for (int i = 0; i < chars; i++)
            {
                bytes[i] = Convert.ToByte(match.Groups[1].Value.Substring(i * 2, 2), 16);
            }
            builder.Append(Encoding.Latin1.GetString(bytes));
        }
        return builder.ToString();
    }

    internal static string ExtractCidMappedText(string pdf)
    {
        var unicodes = ReadCidTextMap(pdf);
        return string.Concat(Regex.Matches(pdf, @"BT(.*?)ET", RegexOptions.Singleline)
            .Select(region => DecodeCidTextShows(region.Groups[1].Value, unicodes)));
    }

    internal static Dictionary<string, string> ReadCidTextMap(string pdf)
    {
        // Embedded subsets remap glyphs to dense CIDs, so extraction decodes
        // BT..ET text operators through the ToUnicode CMaps in the same file.
        var unicodes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (Match block in Regex.Matches(pdf, @"(\d+) beginbfchar(.*?)endbfchar", RegexOptions.Singleline))
        {
            foreach (Match entry in Regex.Matches(block.Groups[2].Value, @"<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4,8})>"))
            {
                unicodes[entry.Groups[1].Value] = char.ConvertFromUtf32(Convert.ToInt32(entry.Groups[2].Value, 16));
            }
        }
        return unicodes;
    }

    internal static string DecodeCidTextShows(string paint, IReadOnlyDictionary<string, string> unicodes)
    {
        var builder = new StringBuilder();
        foreach (Match op in Regex.Matches(paint, @"(?:<[0-9A-Fa-f]+>\s*)+Tj|\[(?:\s*<[0-9A-Fa-f]+>\s*|\s*-?\d+(?:\.\d+)?\s*)+\]\s*TJ"))
        {
            foreach (Match hex in Regex.Matches(op.Value, @"<([0-9A-Fa-f]+)>"))
            {
                string digits = hex.Groups[1].Value;
                for (int i = 0; i + 4 <= digits.Length; i += 4)
                {
                    if (unicodes.TryGetValue(digits.Substring(i, 4), out string? rune))
                    {
                        builder.Append(rune);
                    }
                }
            }
        }
        return builder.ToString();
    }

    private sealed class FixedBytesResolver(byte[] bytes) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(request.FamilyName, request.FamilyName, new FontStyleKey(request.Bold, request.Italic, 400, 0, false), new MemoryFontProgramSource("test:fixed", bytes), true);
        }
    }
    
    // RV06 (RV01 residual): comment balloon text with an empty font source keeps
    // extractable balloon words with a diagnosed fallback.
    public static void MissingFontsPreserveMarkupBalloonText()
    {
        string input = WriteCommentBalloonDocx();
        var resolver = new MapFontResolver([], "Fallback");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocxMarkup(input, resolver);
        string extracted = ExtractWinAnsiText(pdf);
        // Balloon bodies trim to margin-derived length with or without fonts
        // (TrimBalloonText), so the fallback preserves the trimmed body text.
        TestAssert.Contains("Balloon word", extracted);
        TestAssert.Contains("Reviewer", extracted);
        TestAssert.Contains("Anchored", extracted);
        TestAssert.True(diagnostics.Any(d => d.Id == "FONT_NO_USABLE_FACE"), "Balloon fallback must stay diagnosed.");
    }

    // RV06 (RV01 residual): word-compatible balloons keep their word-compatible
    // strings through the diagnosed fallback instead of substituting default text.
    public static void MissingFontsPreserveWordCompatibleBalloonText()
    {
        string input = WriteCommentBalloonDocx();
        var resolver = new MapFontResolver([], "Fallback");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocxWordCompatibleMarkup(input, resolver);
        string extracted = ExtractWinAnsiText(pdf);
        TestAssert.Contains("Commented [RV1]", extracted);
        TestAssert.Contains("Balloon word", extracted);
        TestAssert.Contains("Anchored", extracted);
        TestAssert.True(diagnostics.Any(d => d.Id == "FONT_NO_USABLE_FACE"), "Word-compatible fallback must stay diagnosed.");

        foreach (string wordSequence in new[] { new string('W', 84), string.Join(" ", Enumerable.Repeat("WWWWWWW", 12)) })
        {
            const string tail = ", Tail witness.";
            (string wrapped, List<OoxPdfDiagnostic> wrappedDiagnostics) =
                ConvertDocxWordCompatibleMarkup(WriteCommentBalloonDocx(wordSequence + tail), resolver);
            TestAssert.Contains("/Type /Pages /Count 1", wrapped);
            TestAssert.Contains(new string('W', 84) + ",Tailwitness.",
                string.Concat(ExtractWinAnsiText(wrapped).Where(character => !char.IsWhiteSpace(character))));
            TestAssert.True(wrappedDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_NO_USABLE_FACE"),
                "Wrapped fallback must stay diagnosed.");
            var glyphs = Regex.Matches(wrapped, @"1 0 0 1 (?<x>[0-9.]+) (?<y>[0-9.]+) Tm\s*<57> Tj");
            TestAssert.Equal(84, glyphs.Count);
            TestAssert.True(glyphs.Select(glyph => glyph.Groups["y"].Value).Distinct().Count() > 1,
                "Long fallback comments must use multiple baselines.");
            TestAssert.True(glyphs.All(glyph => double.Parse(glyph.Groups["x"].Value, CultureInfo.InvariantCulture) + 5d <= 612d),
                "Every fallback glyph must fit on the printed page.");
        }
    }
    public static void EmbeddedBalloonTextSubstitutesMissingGlyphs()
    {
        var resolver = new FixedBytesResolver(TestFontBuilder.CreateTestFont());
        const string missing = "\U0001F9EA";
        string input = WriteCommentBalloonDocx("Start" + missing + "End", "Agent" + missing + "Name");
        var (preserved, preserveDiagnostics) = RenderEmbeddedCommentBalloon(input, resolver, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var (review, reviewDiagnostics) = RenderEmbeddedCommentBalloon(input, resolver, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        TestAssert.Contains("Start?End", preserved);
        TestAssert.Contains("Start?End", review);
        TestAssert.Contains("Agent?Name", preserved);
        TestAssert.Contains("U+1F9EA", preserveDiagnostics.Single(d => d.Id == "DOCX_BALLOON_TITLE_MISSING_GLYPHS").Message);
        TestAssert.True(reviewDiagnostics.All(d => d.Id != "DOCX_BALLOON_TITLE_MISSING_GLYPHS"),
            "The hidden author name does not contribute a missing title glyph.");
        foreach (var diagnostics in new[] { preserveDiagnostics, reviewDiagnostics })
        {
            TestAssert.Contains("U+1F9EA", diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS").Message);
        }

    }

    public static void BalloonPreviewTruncatesAtTextElementBoundary()
    {
        var resolver = new FixedBytesResolver(TestFontBuilder.CreateTestFont());
        const string margins = """<w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="2880"/>""";
        const string suffix = "XYZ tail text";
        string boundary = new('i', 34);
        string fits = new('i', 33);
        foreach (string grapheme in new[] { "\U0001F600", "A\u0301" })
        foreach (var (prefix, expected) in new[] { (boundary, boundary), (fits, fits + grapheme) })
        {
            string input = WriteCommentBalloonDocx(prefix + grapheme + suffix, pageMargins: margins);
            var (text, diagnostics) = RenderEmbeddedCommentBalloon(input, resolver, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.True(text.EndsWith(expected + "...", StringComparison.Ordinal), "The preview keeps complete text elements within the existing cut.");
            TestAssert.DoesNotContain("\uFFFD", text);
            TestAssert.True(diagnostics.All(d => d.Id != "FONT_MISSING_GLYPHS"), "All authored characters have glyphs in the pinned face.");
        }
    }

    private static (string Text, List<OoxPdfDiagnostic> Diagnostics) RenderEmbeddedCommentBalloon(
        string input, IFontResolver resolver, OoxPdfDocxMarkupGeometryMode geometry)
    {
        using var stream = File.OpenRead(input);
        using var package = OoxPackage.Open(stream, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.AllMarkup;
        var document = new DocxReader().Read(package, diagnostics.Add, default, mode);
        PdfPage page = new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
        return (text, diagnostics);
    }

    private static string WriteCommentBalloonDocx(string body = "Balloon words here", string author = "Reviewer", string pageMargins = "")
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>()
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/>"
                + "</Types>"),
            ["_rels/.rels"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>"),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"comments.xml\"/>"
                + "</Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body><w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>Anchored</w:t></w:r><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" + pageMargins + "</w:sectPr>"
                + "</w:body></w:document>"),
            ["word/comments.xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:comment w:id=\"1\" w:author=\"" + System.Security.SecurityElement.Escape(author)
                + "\" w:initials=\"RV\" w:date=\"2026-06-01T00:00:00Z\">"
                + "<w:p><w:r><w:t>" + System.Security.SecurityElement.Escape(body) + "</w:t></w:r></w:p>"
                + "</w:comment></w:comments>"),
        });
    }
    private static (string Pdf, List<OoxPdfDiagnostic> Diagnostics) ConvertDocxMarkup(string input, IFontResolver resolver)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, DiagnosticSink = diagnostics.Add, DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup });
        return (File.ReadAllText(output, Encoding.Latin1), diagnostics);
    }

    private static (string Pdf, List<OoxPdfDiagnostic> Diagnostics) ConvertDocxWordCompatibleMarkup(string input, IFontResolver resolver)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, DiagnosticSink = diagnostics.Add, DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup, DocxMarkupGeometryMode = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup });
        return (File.ReadAllText(output, Encoding.Latin1), diagnostics);
    }
}
