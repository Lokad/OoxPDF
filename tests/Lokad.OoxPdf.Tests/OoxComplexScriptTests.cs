using System.Security;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class OoxComplexScriptTests
{
    public static void ComplexScriptDocxReportsApproximationOnce() =>
        AssertWarnings(ReadWarnings(CreateDocx("مرحبا", "مرحبا"), OoxPdfInputKind.Docx),
            ("joining scripts", "cursive joining"));

    public static void ComplexScriptPptxReportsApproximationOnce() =>
        AssertWarnings(ReadWarnings(CreatePptx(["مرحبا"], ["مرحبا"]), OoxPdfInputKind.Pptx),
            ("joining scripts", "cursive joining"));

    public static void ComplexScriptLatinStaysQuiet()
    {
        AssertWarnings(ReadWarnings(CreateDocx("Hello world", "日本語"), OoxPdfInputKind.Docx));
        AssertWarnings(ReadWarnings(CreatePptx(["Hello world", "日本語"]), OoxPdfInputKind.Pptx));
    }

    public static void ComplexScriptDocxReportsRemainingFamilies() =>
        AssertWarnings(ReadWarnings(CreateDocx("שלום", "कमल", "e\u0301", "Hello world"), OoxPdfInputKind.Docx),
            ("reordering scripts", "syllable reordering"), ("bidirectional text", "bidirectional reordering"),
            ("combining marks", "mark positioning"));

    public static void ComplexScriptPptxReportsRemainingFamilies() =>
        AssertWarnings(ReadWarnings(CreatePptx(["שלום", "कमल", "e\u0301", "Hello world"]), OoxPdfInputKind.Pptx),
            ("reordering scripts", "syllable reordering"), ("bidirectional text", "bidirectional reordering"),
            ("combining marks", "mark positioning"));

    private static OoxPdfDiagnostic[] ReadWarnings(MemoryStream input, OoxPdfInputKind kind)
    {
        using (input)
        using (var output = new MemoryStream())
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = kind, DiagnosticSink = diagnostics.Add });
            return diagnostics.Where(diagnostic => diagnostic.Id == "COMPLEX_SCRIPT_APPROXIMATION").ToArray();
        }
    }

    private static void AssertWarnings(OoxPdfDiagnostic[] actual, params (string Feature, string Behavior)[] expected)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < expected.Length; i++)
        {
            TestAssert.Equal(OoxPdfSeverity.Warning, actual[i].Severity);
            TestAssert.Equal(expected[i].Feature, actual[i].Feature);
            TestAssert.Equal("Unshaped glyphs", actual[i].Fallback);
            TestAssert.Contains(expected[i].Behavior, actual[i].Message);
        }
    }

    private static MemoryStream CreateDocx(params string[] paragraphs)
    {
        string body = string.Concat(paragraphs.Select(text => $"<w:p><w:r><w:t>{SecurityElement.Escape(text)}</w:t></w:r></w:p>"));
        return TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
    }

    private static MemoryStream CreatePptx(params string[][] slides)
    {
        var parts = new Dictionary<string, string> { ["_rels/.rels"] = PptxTests.PackageRelationship() };
        var types = new List<string>();
        var relationships = new List<string>();
        var slideIds = new List<string>();
        for (int i = 0; i < slides.Length; i++)
        {
            int slide = i + 1;
            string part = $"slides/slide{slide}.xml";
            types.Add($"""<Override PartName="/ppt/{part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>""");
            relationships.Add($"""<Relationship Id="rId{slide}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="{part}"/>""");
            slideIds.Add($"""<p:sldId id="{255 + slide}" r:id="rId{slide}"/>""");
            string shapes = string.Concat(slides[i].Select((text, index) => FormattableString.Invariant($"""
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="{index + 2}" name="Script sample"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr>
                    <a:xfrm><a:off x="914400" y="{914400 + index * 1371600}"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                    <a:prstGeom prst="rect"/><a:noFill/>
                  </p:spPr>
                  <p:txBody>
                    <a:bodyPr/><a:lstStyle/>
                    <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>{SecurityElement.Escape(text)}</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>
                """)));
            parts[$"ppt/{part}"] = $"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>{shapes}</p:spTree></p:cSld>
                </p:sld>
                """;
        }

        parts["[Content_Types].xml"] = $"""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              {string.Concat(types)}
            </Types>
            """;
        parts["ppt/_rels/presentation.xml.rels"] = $"""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{string.Concat(relationships)}</Relationships>
            """;
        parts["ppt/presentation.xml"] = $"""
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldIdLst>{string.Concat(slideIds)}</p:sldIdLst>
              <p:sldSz cx="9144000" cy="6858000"/>
            </p:presentation>
            """;
        return TestFixtures.CreateZipPackage(parts);
    }
}
