using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class OoxComplexScriptTests
{
    // RV02: joining text reports one approximation per conversion, not per run.
    public static void ComplexScriptDocxReportsApproximationOnce()
    {
        string input = WriteJoiningDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        OoxPdfDiagnostic[] matches = diagnostics.Where(d => d.Id == "COMPLEX_SCRIPT_APPROXIMATION").ToArray();
        // Two paragraphs carry the same joining payload, so dedup keeps one warning.
        TestAssert.Equal(1, matches.Length);
        TestAssert.Equal(OoxPdfSeverity.Warning, matches[0].Severity);
        TestAssert.Equal("joining scripts", matches[0].Feature);
        TestAssert.Equal("Unshaped glyphs", matches[0].Fallback);
        TestAssert.True(matches[0].Message.Contains("cursive joining", StringComparison.Ordinal), "Joining diagnostic must name the missing behavior.");
    }

    // RV02: two slides share one resolver, so cross slide preparation still warns once.
    public static void ComplexScriptPptxReportsApproximationOnce()
    {
        string input = WriteJoiningPptx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        OoxPdfDiagnostic[] matches = diagnostics.Where(d => d.Id == "COMPLEX_SCRIPT_APPROXIMATION").ToArray();
        // Two slides carry the same joining payload, so resolver dedup keeps one warning.
        TestAssert.Equal(1, matches.Length);
        TestAssert.Equal(OoxPdfSeverity.Warning, matches[0].Severity);
        TestAssert.Equal("joining scripts", matches[0].Feature);
        TestAssert.Equal("Unshaped glyphs", matches[0].Fallback);
        TestAssert.True(matches[0].Message.Contains("cursive joining", StringComparison.Ordinal), "Joining diagnostic must name the missing behavior.");
    }

    // RV02: Latin and CJK need no shaping, so quiet input must stay quiet.
    public static void ComplexScriptLatinStaysQuiet()
    {
        string docxInput = WriteLatinDocx();
        string docxOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> docxDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(docxInput, docxOutput, new OoxPdfOptions { DiagnosticSink = docxDiagnostics.Add });
        TestAssert.True(docxDiagnostics.All(d => d.Id != "COMPLEX_SCRIPT_APPROXIMATION"), "Latin DOCX must not warn about complex scripts.");

        string pptxInput = WriteLatinPptx();
        string pptxOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> pptxDiagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(pptxInput, pptxOutput, new OoxPdfOptions { DiagnosticSink = pptxDiagnostics.Add });
        TestAssert.True(pptxDiagnostics.All(d => d.Id != "COMPLEX_SCRIPT_APPROXIMATION"), "Latin PPTX must not warn about complex scripts.");
    }

    private static string WriteJoiningDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>مرحبا</w:t></w:r></w:p>
                    <w:p><w:r><w:t>مرحبا</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string WriteLatinDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Hello world</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string WriteJoiningPptx()
    {
        string slideOne = SlideWithArabic();
        string slideTwo = SlideWithArabic();
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = slideOne,
            ["ppt/slides/slide2.xml"] = slideTwo
        });
    }

    private static string SlideWithArabic()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>مرحبا</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
            """;
    }

    private static string WriteLatinPptx()
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = SlideWithLatin()
        });
    }

    private static string SlideWithLatin()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Hello world</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
            """;
    }

    // RV02: remaining families each report once with the matching feature name.
    public static void ComplexScriptDocxReportsRemainingFamilies()
    {
        string input = WriteMultiFamilyDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        OoxPdfDiagnostic[] matches = diagnostics.Where(d => d.Id == "COMPLEX_SCRIPT_APPROXIMATION").ToArray();
        // Hebrew, Devanagari and combining-mark paragraphs each contribute one family warning.
        TestAssert.Equal(3, matches.Length);
        TestAssert.True(matches.Any(d => d.Feature == "bidirectional text"), "Hebrew must report the bidirectional family.");
        TestAssert.True(matches.Any(d => d.Feature == "reordering scripts"), "Devanagari must report the reordering family.");
        TestAssert.True(matches.Any(d => d.Feature == "combining marks"), "Combining marks must report the mark family.");
        TestAssert.True(matches.All(d => d.Severity == OoxPdfSeverity.Warning), "Approximation warnings stay warnings.");
        TestAssert.True(matches.All(d => d.Fallback == "Unshaped glyphs"), "All families share the unshaped-glyph fallback.");
    }

    // RV02: one slide with three shaped-as-source shapes still warns once per family.
    public static void ComplexScriptPptxReportsRemainingFamilies()
    {
        string input = WriteMultiFamilyPptx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        List<OoxPdfDiagnostic> diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        OoxPdfDiagnostic[] matches = diagnostics.Where(d => d.Id == "COMPLEX_SCRIPT_APPROXIMATION").ToArray();
        // Hebrew, Devanagari and combining-mark shapes each contribute one family warning.
        TestAssert.Equal(3, matches.Length);
        TestAssert.True(matches.Any(d => d.Feature == "bidirectional text"), "Hebrew must report the bidirectional family.");
        TestAssert.True(matches.Any(d => d.Feature == "reordering scripts"), "Devanagari must report the reordering family.");
        TestAssert.True(matches.Any(d => d.Feature == "combining marks"), "Combining marks must report the mark family.");
        TestAssert.True(matches.All(d => d.Severity == OoxPdfSeverity.Warning), "Approximation warnings stay warnings.");
        TestAssert.True(matches.All(d => d.Fallback == "Unshaped glyphs"), "All families share the unshaped-glyph fallback.");
    }

    private static string WriteMultiFamilyDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>שלום</w:t></w:r></w:p>
                    <w:p><w:r><w:t>कमल</w:t></w:r></w:p>
                    <w:p><w:r><w:t>é</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Hello world</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string WriteMultiFamilyPptx()
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = SlideWithMultiFamily()
        });
    }

    private static string SlideWithMultiFamily()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>שלום</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                  <p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="2286000"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>कमल</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                  <p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>é</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
            """;
    }
}
