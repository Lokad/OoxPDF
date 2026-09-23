using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

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
        TestAssert.Contains("/Type /Pages /Count 3", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.DoesNotContain("/FontFile2", pdf);
        string extracted = ExtractWinAnsiText(pdf);
        TestAssert.Contains("First", extracted);
        TestAssert.Contains("Second", extracted);
        TestAssert.Contains("Third", extracted);
        OoxPdfDiagnostic missing = diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE");
        TestAssert.Equal(OoxPdfSeverity.Warning, missing.Severity);
        string repeat = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, repeat, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver });
        string first = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, first, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver });
        TestAssert.True(File.ReadAllBytes(first).SequenceEqual(File.ReadAllBytes(repeat)), "Fallback output must be deterministic.");
    }

    public static void InvalidFontBytesPreserveDocxPagesAndText()
    {
        // RV01: undecodable font programs take the same diagnosed fallback.
        string input = WritePageBreakDocx("First", "Second", "Third");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, new FixedBytesResolver(new byte[] { 1, 2, 3, 4 }));
        TestAssert.Contains("/Type /Pages /Count 3", pdf);
        TestAssert.Contains("BT", pdf);
        string extracted = ExtractWinAnsiText(pdf);
        TestAssert.Contains("First", extracted);
        TestAssert.Contains("Third", extracted);
        TestAssert.Equal(OoxPdfSeverity.Warning, diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE").Severity);
    }

    public static void UnsupportedOutlineFontsPreserveDocxPagesAndText()
    {
        // RV01: CFF outlines cannot embed, so their runs join the fallback.
        string input = WritePageBreakDocx("First", "Second", "Third");
        (string pdf, List<OoxPdfDiagnostic> diagnostics) = ConvertDocx(input, new FixedBytesResolver(TestFontBuilder.CreateCffKindFont()));
        TestAssert.Contains("/Type /Pages /Count 3", pdf);
        TestAssert.Contains("BT", pdf);
        string extracted = ExtractWinAnsiText(pdf);
        TestAssert.Contains("First", extracted);
        TestAssert.Contains("Third", extracted);
        TestAssert.Equal(OoxPdfSeverity.Warning, diagnostics.Single(d => d.Id == "FONT_NO_USABLE_FACE").Severity);
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

    public static void MissingFontsPreservePptxText()
    {
        // RV01: the PPTX audit shows the same silent loss; text stays visible and
        // diagnosed with an empty font source (fails: no BT, no diagnostic).
        string input = FindCase("pptx-ladder-02-plain-text.pptx");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new MapFontResolver([], "Fallback"), DiagnosticSink = diagnostics.Add });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        TestAssert.Contains("/Type /Pages /Count 1", pdf);
        TestAssert.Contains("BT", pdf);
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

    private static string ExtractWinAnsiText(string pdf)
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

    private static string ExtractCidMappedText(string pdf)
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

        var builder = new StringBuilder();
        foreach (Match region in Regex.Matches(pdf, @"BT(.*?)ET", RegexOptions.Singleline))
        {
            foreach (Match op in Regex.Matches(region.Groups[1].Value, @"(?:<[0-9A-Fa-f]+>\s*)+Tj|\[(?:\s*<[0-9A-Fa-f]+>\s*|\s*-?\d+(?:\.\d+)?\s*)+\]\s*TJ"))
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
        }

        return builder.ToString();
    }

    private static string FindCase(string name)
    {
        string[] candidates = new[]
        {
            Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", name),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name),
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }
        throw new InvalidOperationException("Case file not found: " + name);
    }

    private sealed class FixedBytesResolver(byte[] bytes) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(request.FamilyName, request.FamilyName, new FontStyleKey(request.Bold, request.Italic, 400, 0, false), new MemoryFontProgramSource("test:fixed", bytes), true);
        }
    }
}
