using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlTests
{
    public static void ParsesContentTypesAndRelationships()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = "<p:presentation xmlns:p=\"p\"/>",
            ["ppt/slides/slide1.xml"] = "<p:sld xmlns:p=\"p\"/>"
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);

        OoxPart presentation = TestAssert.NotNull(package.GetPart("/ppt/presentation.xml"));
        TestAssert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", presentation.ContentType);

        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/", CancellationToken.None);
        TestAssert.Equal(1, relationships.Count);
        TestAssert.Equal("/ppt/presentation.xml", relationships[0].ResolvedTarget);
    }

    public static void RejectsPackagePartPathTraversal()
    {
        TestAssert.Throws<InvalidDataException>(() => OoxPath.NormalizePartName("../evil.xml"));
    }

    public static void ResolvesRelationshipTargets()
    {
        string resolved = OoxPath.ResolveRelationshipTarget("/ppt/slides/slide1.xml", "../media/image1.png");

        TestAssert.Equal("/ppt/media/image1.png", resolved);
    }

    public static void StrictOoxmlDialectWarnsForDocx()
    {
        string input = WriteDialectProbeDocx("http://purl.oclc.org/ooxml/wordprocessingml/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_STRICT_DIALECT"), "Strict DOCX content should fail visibly instead of converting silently empty.");
    }

    public static void TransitionalOoxmlDialectStaysQuietForDocx()
    {
        string input = WriteDialectProbeDocx("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional DOCX content should not warn about dialects.");
    }

    public static void StrictOoxmlDialectWarnsForPptx()
    {
        string input = WriteDialectProbePptx("http://purl.oclc.org/ooxml/presentationml/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_STRICT_DIALECT"), "Strict PPTX content should fail visibly instead of converting silently empty.");
    }

    public static void TransitionalOoxmlDialectStaysQuietForPptx()
    {
        string input = WriteDialectProbePptx("http://schemas.openxmlformats.org/presentationml/2006/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional PPTX content should not warn about dialects.");
    }

    private static string WriteDialectProbeDocx(string wordNamespace)
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
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="{{wordNamespace}}">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string WriteDialectProbePptx(string presentationNamespace)
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
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
                </Relationships>
                """,
            ["ppt/presentation.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="{{presentationNamespace}}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
    }

    public static void AlternateContentUnderstoodVmlChoiceWins()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
              <mc:AlternateContent>
                <mc:Choice Requires="v"><v:shape>VectorShape</v:shape></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace v = "urn:schemas-microsoft-com:vml";
        TestAssert.Equal("VectorShape", string.Concat(document.Descendants(v + "shape").Select(e => e.Value)));
    }

    public static void AlternateContentUnderstoodChoiceWins()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:AlternateContent>
                <mc:Choice Requires="w"><w:a>ChoiceText</w:a></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("ChoiceText", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent" || e.Name.LocalName == "Choice" || e.Name.LocalName == "Fallback"), "Exactly one representation should survive.");
    }

    public static void AlternateContentUnknownChoiceFallsBack()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <mc:AlternateContent>
                <mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("FallbackText", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
    }

    public static void AlternateContentNestedBlocksResolveInnermostFirst()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <mc:AlternateContent>
                <mc:Choice Requires="wps"><wpc:canvas><mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><w:a>InnerFallback</w:a></mc:Fallback></mc:AlternateContent></wpc:canvas></mc:Choice>
                <mc:Fallback><w:a>OuterFallback</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("OuterFallback", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent"), "Nested blocks should resolve fully.");
    }

    public static void AlternateContentWithoutFallbackIsRemoved()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <w:keep>Keep</w:keep>
              <mc:AlternateContent>
                <mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("Keep", string.Concat(document.Descendants(w + "keep").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent"), "Fallback-less unknown blocks should vanish.");
    }

    public static void AlternateContentRendersExactlyOneRepresentation()
    {
        string input = WriteAlternateContentProbeDocx("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", "ChoiceText", "FallbackText");

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        Lokad.OoxPdf.Docx.DocxDocument document = new Lokad.OoxPdf.Docx.DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        string text = string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
        TestAssert.Equal("FallbackText", text);
    }

    public static void AlternateContentUnderstoodChoiceRenders()
    {
        string input = WriteAlternateContentProbeDocx("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "ChoiceText", "FallbackText");

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        Lokad.OoxPdf.Docx.DocxDocument document = new Lokad.OoxPdf.Docx.DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        string text = string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
        TestAssert.Equal("ChoiceText", text);
    }

    private static string WriteAlternateContentProbeDocx(string choicePrefix, string choiceNamespace, string choiceText, string fallbackText)
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
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                  <w:body>
                    <w:p>
                      <mc:AlternateContent>
                        <mc:Choice Requires="{{choicePrefix}}" xmlns:{{choicePrefix}}="{{choiceNamespace}}"><w:r><w:t>{{choiceText}}</w:t></w:r></mc:Choice>
                        <mc:Fallback><w:r><w:t>{{fallbackText}}</w:t></w:r></mc:Fallback>
                      </mc:AlternateContent>
                    </w:p>
                  </w:body>
                </w:document>
                """
        });
    }

    public static void StrictRelationshipTypesNormalizeToTransitional()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slide" Target="ppt/slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="ppt/charts/chart1.xml"/>
                </Relationships>
                """
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);

        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/", CancellationToken.None);
        TestAssert.Equal(2, relationships.Count);
        TestAssert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", relationships[0].Type);
        TestAssert.Equal("/ppt/slides/slide1.xml", relationships[0].ResolvedTarget);
        TestAssert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relationships[1].Type);
    }

    public static void StrictRelationshipTypesStillResolveSlides()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
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
                  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional parts with Strict relationship arcs should resolve without dialect warnings.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForDocx()
    {
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <wpc:canvas mc:MustUnderstand="1"/>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_MUST_UNDERSTAND"), "Must-understand content outside the understood set should fail visibly.");
    }

    public static void MustUnderstandKnownNamespaceStaysQuietForDocx()
    {
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <v:shape mc:MustUnderstand="1"/>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_MUST_UNDERSTAND"), "Must-understand content inside the understood set should stay quiet.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForPptx()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
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
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                  <p14:creationId mc:MustUnderstand="1"/>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_MUST_UNDERSTAND"), "Must-understand content outside the understood set should fail visibly.");
    }

    public static void AlternateContentFallbackOnlyPayloadSurvives()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:AlternateContent>
                <mc:Fallback><w:a>FallbackOnly</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.ResolveAlternateContent(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("FallbackOnly", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
    }

    public static void ConvertsCommonOfficeUnits()
    {
        TestAssert.Equal(72d, OoxUnits.EmuToPoints(914400));
        TestAssert.Equal(12d, OoxUnits.TwipsToPoints(240));
        TestAssert.Equal(9d, OoxUnits.HalfPointsToPoints(18));
    }
}
