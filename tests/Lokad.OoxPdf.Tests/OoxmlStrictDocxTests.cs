using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlStrictDocxTests
{
    public static void StrictDocxCommonVocabularyPreservesNamesPrefixesAndMeasures()
    {
        const string xml = """
            <w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
              xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
              xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:x="urn:unhandled"
              xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:MustUnderstand="w">
              <w:body><w:p><w:pPr><w:spacing w:after="6pt"/><w:ind w:start="0.5in"/></w:pPr>
              <w:r><w:rPr><w:sz w:val="11pt"/></w:rPr><w:t>Literal 50% and 6pt</w:t></w:r>
              <w:hyperlink r:id="link"/><x:measure x:val="6pt"/></w:p>
              <w:tbl><w:tblPr><w:tblW w:w="50%" w:type="pct"/></w:tblPr></w:tbl>
              <a:alpha val="63%"/><a:latin typeface="Literal 50%"/>
              <mc:AlternateContent><mc:Choice Requires="w"><w:p>Chosen</w:p></mc:Choice><mc:Fallback><w:p>Rejected</w:p></mc:Fallback></mc:AlternateContent>
              </w:body></w:document>
            """;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        XDocument document = SafeXml.Load(stream, default);
        XNamespace w = OoxNamespaces.WordprocessingNamespace;
        TestAssert.Equal(w + "document", document.Root!.Name);
        TestAssert.Equal(w, document.Root.GetNamespaceOfPrefix("w"));
        TestAssert.Equal(OoxNamespaces.RelationshipsNamespace, document.Root.GetNamespaceOfPrefix("r"));
        TestAssert.Equal("120", (string?)document.Descendants(w + "spacing").Single().Attribute(w + "after"));
        TestAssert.Equal("720", (string?)document.Descendants(w + "ind").Single().Attribute(w + "start"));
        TestAssert.Equal("22", (string?)document.Descendants(w + "sz").Single().Attribute(w + "val"));
        TestAssert.Equal("2500", (string?)document.Descendants(w + "tblW").Single().Attribute(w + "w"));
        TestAssert.Equal("link", (string?)document.Descendants(w + "hyperlink").Single().Attribute(OoxNamespaces.RelationshipsNamespace + "id"));
        TestAssert.Equal("63000", (string?)document.Descendants(OoxNamespaces.DrawingNamespace + "alpha").Single().Attribute("val"));
        TestAssert.Equal("Literal 50%", (string?)document.Descendants(OoxNamespaces.DrawingNamespace + "latin").Single().Attribute("typeface"));
        TestAssert.Equal("Literal 50% and 6pt", document.Descendants(w + "t").Single().Value);
        TestAssert.Equal("6pt", (string?)document.Descendants(XName.Get("measure", "urn:unhandled")).Single().Attribute(XName.Get("val", "urn:unhandled")));
        TestAssert.True(OoxNamespaces.HasStrictOoxmlOrigin(document), "Dialect provenance must survive normalization.");
        TestAssert.True(!OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document), "Known Strict prefix aliases must remain understood.");
        TestAssert.True(document.Descendants(w + "p").Any(p => p.Value == "Chosen") &&
            document.Descendants(w + "p").All(p => p.Value != "Rejected"), "Strict aliases must resolve before alternate-content selection.");
    }

    public static void StrictDocxMeasuresRetainBoundsCancellationAndTransitionalBytes()
    {
        foreach (string value in new[] { "72pt", "1in", "2.54cm", "25.4mm", "6pc", "6pi" })
        {
            XDocument document = XDocument.Parse($"""<w:pgSz xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" w:w="{value}"/>""");
            OoxStrictDocx.Normalize(document, default);
            TestAssert.Equal("1440", (string?)document.Root!.Attribute(OoxNamespaces.WordprocessingNamespace + "w"));
        }
        XDocument unchanged = XDocument.Parse("""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>""");
        string before = unchanged.ToString();
        OoxStrictDocx.Normalize(unchanged, default);
        TestAssert.Equal(before, unchanged.ToString());
        TestAssert.True(!OoxNamespaces.HasStrictOoxmlOrigin(unchanged), "Transitional input must retain its origin.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxStrictDocx.Normalize(unchanged, canceled.Token));
        foreach (string value in new[] { "NaNpt", "99999999999999999999999999999pt" })
        {
            XDocument document = XDocument.Parse($"""<w:pgSz xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" w:w="{value}"/>""");
            TestAssert.Throws<InvalidDataException>(() => OoxStrictDocx.Normalize(document, default));
        }
        XDocument unsupported = XDocument.Parse("""<p:presentation xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"/>""");
        OoxStrictDocx.Normalize(unsupported, default);
        TestAssert.Equal("http://purl.oclc.org/ooxml/presentationml/main", unsupported.Root!.Name.NamespaceName);
    }

    public static void StrictDocxOfficeCompositionMatchesTransitional()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Strict"));
        using var metadata = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "report-office.json")));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var outputs = new List<byte[]>();
            foreach (string name in new[] { "report-transitional.docx", "report-strict.docx" })
            {
                string input = Path.Combine(directory, name);
                TestAssert.Equal(metadata.RootElement.GetProperty("FixtureHashes").GetProperty(name).GetString(),
                    Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
                using var output = new MemoryStream();
                using var stream = File.OpenRead(input);
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(stream, output, new OoxPdfOptions
                {
                    InputKind = OoxPdfInputKind.Docx,
                    FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true),
                    DocxMarkupMode = mode,
                    DocxMarkupGeometryMode = mode == OoxPdfDocxMarkupMode.Final
                        ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                        : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,
                    DiagnosticSink = diagnostics.Add
                });
                TestAssert.Equal(name.Contains("strict.docx"), diagnostics.Any(d => d.Id == "OOXML_STRICT_DIALECT"));
                TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Supported composition math must survive the dialect change.");
                outputs.Add(output.ToArray());
            }
            TestAssert.True(outputs[0].SequenceEqual(outputs[1]), "The Office-authored dialect pair must retain identical supported rendering.");
        }
    }
}
