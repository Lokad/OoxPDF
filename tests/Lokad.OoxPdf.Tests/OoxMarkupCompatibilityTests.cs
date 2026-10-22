using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxMarkupCompatibilityTests
{
    private const string Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    public static void ProcessContentUsesDeclarationNamespaceScope()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:first"
                        mc:Ignorable="e" mc:ProcessContent="e:box">
              <w:body xmlns:e="urn:second" xmlns:original="urn:first">
                <original:box><w:p>Inherited</w:p></original:box>
                <e:box><w:p>Unrecognized wrapper</w:p></e:box>
                <e:box mc:Ignorable="e" mc:ProcessContent="e:box"><w:p>Local</w:p></e:box>
              </w:body>
            </w:document>
            """);
        XElement body = TestAssert.NotNull(document.Root?.Element(XName.Get("body", Word)));
        TestAssert.Equal("InheritedLocal", string.Concat(body.Elements(XName.Get("p", Word)).Select(p => p.Value)));
        TestAssert.Equal(1, body.Elements(XName.Get("box", "urn:second")).Count());
    }

    public static void ProcessContentRequiresIgnorableNamespaceAndPreservesKnownElements()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension"
                        mc:Ignorable="w" mc:ProcessContent="w:p e:box">
              <w:body><w:p>Known</w:p><e:box><w:p>Not ignorable</w:p></e:box></w:body>
            </w:document>
            """);
        XElement body = TestAssert.NotNull(document.Root?.Element(XName.Get("body", Word)));
        TestAssert.Equal("Known", TestAssert.NotNull(body.Element(XName.Get("p", Word))).Value);
        TestAssert.NotNull(body.Element(XName.Get("box", "urn:extension")));
    }

    public static void ProcessContentSupportsNestedAndWildcardDeclarations()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension"
                        mc:Ignorable="e" mc:ProcessContent="e:*">
              <w:body><e:outer><e:inner><w:p>Kept</w:p></e:inner></e:outer></w:body>
            </w:document>
            """);
        TestAssert.Equal("Kept", document.Root?.Element(XName.Get("body", Word))?.Element(XName.Get("p", Word))?.Value);
        TestAssert.Equal(0, document.Descendants().Count(e => e.Name.NamespaceName == "urn:extension"));
    }

    public static void ProcessContentResolvesOnlySelectedAlternateBranch()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}">
              <mc:AlternateContent>
                <mc:Choice Requires="w" xmlns:e="urn:extension" mc:Ignorable="e" mc:ProcessContent="e:box">
                  <e:box><w:p>Selected</w:p></e:box>
                </mc:Choice>
                <mc:Fallback xmlns:e="urn:extension" mc:MustUnderstand="e"><w:p>Unselected</w:p></mc:Fallback>
              </mc:AlternateContent>
            </w:document>
            """);
        TestAssert.Equal("Selected", document.Root?.Element(XName.Get("p", Word))?.Value);
        TestAssert.True(!OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document), "Unselected demands must not warn.");
    }

    public static void MustUnderstandUsesPrefixListsAndDeclarationScope()
    {
        foreach (string prefixes in new[] { "w", " w\t w\n" })
        {
            XDocument known = Load($$"""<w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" mc:MustUnderstand="{{prefixes}}"/>""");
            TestAssert.True(!OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(known), "Known prefixes should stay quiet.");
        }
        foreach (string prefixes in new[] { "e", "w e", "unresolved" })
        {
            XDocument unknown = Load($$"""<w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension" mc:MustUnderstand="{{prefixes}}"/>""");
            TestAssert.True(OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(unknown), "Unknown or unresolved required prefixes should warn.");
        }
        XDocument rebound = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="{{Word}}" mc:MustUnderstand="e">
              <w:body xmlns:e="urn:extension" mc:MustUnderstand="e"/>
            </w:document>
            """);
        TestAssert.True(OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(rebound), "Child declarations use their own prefix binding.");
    }

    public static void MustUnderstandSurvivesRemovedDeclarationCarriers()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}">
              <mc:AlternateContent>
                <mc:Choice Requires="w" xmlns:e="urn:extension" mc:Ignorable="e" mc:ProcessContent="e:box">
                  <e:box mc:MustUnderstand="e"><w:p>Kept with warning</w:p></e:box>
                </mc:Choice>
              </mc:AlternateContent>
            </w:document>
            """);
        TestAssert.Equal("Kept with warning", document.Root?.Element(XName.Get("p", Word))?.Value);
        TestAssert.True(OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document), "A removed carrier must not erase its active requirement.");
    }

    public static void MustUnderstandInsideIgnoredWrapperDoesNotBecomeActive()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension" mc:Ignorable="e">
              <e:box><w:p mc:MustUnderstand="e">Ignored</w:p></e:box>
            </w:document>
            """);
        TestAssert.NotNull(document.Root?.Element(XName.Get("box", "urn:extension")));
        TestAssert.True(!OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document), "Ignored extension descendants are not selected content.");
    }

    public static void ProcessContentRetainsPreferredChartGalleryAnnotation()
    {
        XDocument document = Load($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension"
              xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
              mc:Ignorable="e" mc:ProcessContent="e:box">
              <e:box><c:chartSpace><mc:AlternateContent>
                <mc:Choice Requires="c14"><c14:style val="108"/></mc:Choice>
                <mc:Fallback><c:style val="2"/></mc:Fallback>
              </mc:AlternateContent></c:chartSpace></e:box>
            </w:document>
            """);
        TestAssert.Equal<int?>(108, OoxMarkupCompatibility.ReadPreferredChartGalleryStyle(document.Root?.Elements().Single()));
    }

    public static void CompatibilityNormalizationHonorsCancellationDuringUnwrap()
    {
        XDocument document = XDocument.Parse($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension"
                        mc:Ignorable="e" mc:ProcessContent="e:box">
              <e:box><w:p>First</w:p></e:box><w:p>Next</w:p>
            </w:document>
            """);
        using var cancellation = new CancellationTokenSource();
        document.Changed += (_, _) => cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxMarkupCompatibility.Normalize(document, cancellation.Token));
    }

    public static void ProcessContentDocxBodyCellAndHeaderRenderIdentically()
    {
        byte[] Convert(bool wrapped)
        {
            string Run(string text) => Wrap($"<w:r><w:t>{text}</w:t></w:r>", wrapped);
            string Paragraph(string text) => Wrap($"<w:p>{Run(text)}</w:p>", wrapped);
            string declarations = $"xmlns:w=\"{Word}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"{Mc}\" xmlns:e=\"urn:extension\" mc:Ignorable=\"e\" mc:ProcessContent=\"e:box\"";
            return Render(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/></Types>",
                ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
                ["word/document.xml"] = $"<w:document {declarations}><w:body>{Paragraph("Body payload")}<w:tbl><w:tblPr><w:tblW w:w=\"4000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>{Paragraph("Cell payload")}</w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type=\"default\" r:id=\"header\"/></w:sectPr></w:body></w:document>",
                ["word/_rels/document.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"header\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/></Relationships>",
                ["word/header1.xml"] = $"<w:hdr {declarations}>{Paragraph("Header payload")}</w:hdr>"
            }, OoxPdfInputKind.Docx);
        }
        byte[] baseline = Convert(false);
        TestAssert.True(baseline.Length > 1000 && baseline.SequenceEqual(Convert(true)), "Declared wrappers must preserve the entire DOCX PDF, including body, cell and header content.");
    }

    public static void ProcessContentPptxTextboxRendersIdentically()
    {
        byte[] Convert(bool wrapped) => Render(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:mc="{{Mc}}" xmlns:e="urn:extension" mc:Ignorable="e" mc:ProcessContent="e:box">
                  <p:cSld><p:spTree>{{Wrap("""
                    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Payload"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2000"/><a:t>Visible payload</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    """, wrapped)}}</p:spTree></p:cSld>
                </p:sld>
                """
        }, OoxPdfInputKind.Pptx);
        byte[] baseline = Convert(false);
        TestAssert.True(baseline.Length > 1000 && baseline.SequenceEqual(Convert(true)), "A declared wrapper must preserve the complete textbox PDF.");
    }

    private static XDocument Load(string xml)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        return SafeXml.Load(stream, CancellationToken.None);
    }

    private static string Wrap(string content, bool wrapped) => wrapped ? $"<e:box>{content}</e:box>" : content;

    private static byte[] Render(Dictionary<string, string> parts, OoxPdfInputKind kind)
    {
        using MemoryStream input = TestFixtures.CreateZipPackage(parts);
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = kind, FontResolver = new TestFaceFontResolver() });
        return output.ToArray();
    }
}
