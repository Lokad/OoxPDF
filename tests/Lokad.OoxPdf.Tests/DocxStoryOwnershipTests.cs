using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxStoryOwnershipTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];

    public static void DocxStoryOwnershipMatchesIndependentWordControls()
    {
        foreach (var mode in Modes)
        {
            var baseline = Render(Case("baseline"), mode);
            foreach (string id in new[] { "unselected-header", "selected-header", "uncached-unselected-header", "uncached-selected-header" })
            {
                var actual = Render(Case(id), mode);
                TestAssert.Equal(id == "uncached-selected-header" ? 1 : 0,
                    actual.Diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_COMPLEX_FIELD"));
                if (id.Contains("unselected", StringComparison.Ordinal))
                {
                    TestAssert.Equal(0, actual.Diagnostics.Count);
                    TestAssert.Equal(string.Join("\n", baseline.Pages.Select(page => page.Content)),
                        string.Join("\n", actual.Pages.Select(page => page.Content)));
                }
                string text = DocxPageFootnoteNumberingTests.DecodeText(actual.Pages.Single());
                TestAssert.Contains("Visible body", text);
                TestAssert.Equal(id == "selected-header", text.Contains("HEADER CONTENT", StringComparison.Ordinal));
            }
        }
    }

    public static void DocxStoryOwnershipSharesTypedInternalHeaderFooterSelection()
    {
        foreach (var mode in Modes)
        foreach (string story in new[] { "header", "footer" })
        foreach (string relationshipCase in new[] { "selected", "unselected", "external", "dangling", "wrong-type" })
            Mutate("uncached-selected-header", (parts, xml) => {
                var reference = xml.Descendants(W+"headerReference").Single();
                reference.Name = W+(story+"Reference");
                if (relationshipCase == "unselected") reference.Remove();
                else if (relationshipCase == "selected")
                {
                    var duplicate = new XElement(reference); duplicate.SetAttributeValue(W+"type", "first");
                    reference.AddAfterSelf(duplicate);
                }
                var rels = Xml(parts, "word/_rels/document.xml.rels");
                var rel = rels.Root!.Elements().Single(e => (string?)e.Attribute("Id") == "reviewHeader");
                rel.SetAttributeValue("Type", R.NamespaceName+"/"+(relationshipCase == "wrong-type" ? (story == "header" ? "footer" : "header") : story));
                if (relationshipCase == "external") rel.SetAttributeValue("TargetMode", "External");
                if (relationshipCase == "dangling") rel.SetAttributeValue("Target", "absent.xml");
                parts["word/_rels/document.xml.rels"] = Bytes(rels);
                var payload = Xml(parts, "word/unused-header.xml"); payload.Root!.Name = W+(story == "header" ? "hdr" : "ftr");
                parts["word/unused-header.xml"] = Bytes(payload);
                var types = Xml(parts, "[Content_Types].xml");
                types.Root!.Elements().Single(e => (string?)e.Attribute("PartName") == "/word/unused-header.xml")
                    .SetAttributeValue("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml."+story+"+xml");
                parts["[Content_Types].xml"] = Bytes(types);
            }, path => {
                var actual = Render(path, mode);
                TestAssert.Equal(relationshipCase == "selected" ? 1 : 0,
                    actual.Diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_COMPLEX_FIELD"));
                TestAssert.Contains("Visible body", DocxPageFootnoteNumberingTests.DecodeText(actual.Pages.Single()));
            });
    }

    public static void DocxStoryOwnershipPreservesCachedNoteXmlAcrossReads()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-section-footnote-roman-4.docx"));
        using var input = File.OpenRead(path);
        var package = OoxPackage.Open(input, default);
        var relation = package.GetRelationships("/word/document.xml", default).Single(r => r.Type == R.NamespaceName+"/footnotes");
        var part = package.GetPart(relation.ResolvedTarget!)!;
        var cached = package.LoadXml(part, default);
        string before = cached.ToString(SaveOptions.DisableFormatting);
        TestAssert.True(cached.Descendants(W+"footnoteRef").Any(), "Control must contain an automatic note mark.");
        foreach (var mode in Modes)
        {
            var document = new DocxReader().Read(package, _ => { }, default, mode);
            TestAssert.True(document.RelatedStories.Any(s => s.Kind == DocxRelatedStoryKind.Footnote), "Note model must remain available.");
            TestAssert.Equal(before, cached.ToString(SaveOptions.DisableFormatting));
            TestAssert.True(ReferenceEquals(cached, package.LoadXml(part, default)), "One package owns one immutable source DOM.");
        }
        using var secondInput = File.OpenRead(path);
        var secondPackage = OoxPackage.Open(secondInput, default);
        TestAssert.True(!ReferenceEquals(cached, secondPackage.LoadXml(secondPackage.GetPart(part.Name)!, default)),
            "Different package owners must not share mutable DOM instances.");
    }

    public static void DocxStoryOwnershipRetainsMalformedRecoveryAndCancellation()
    {
        foreach (bool selected in new[] { false, true })
            Mutate(selected ? "selected-header" : "unselected-header", (parts, _) => {
                parts["word/unused-header.xml"] = Encoding.UTF8.GetBytes("<broken");
            }, path => {
                if (selected) TestAssert.Throws<InvalidDataException>(() => Render(path, OoxPdfDocxMarkupMode.Final));
                else TestAssert.Equal(0, Render(path, OoxPdfDocxMarkupMode.Final).Diagnostics.Count);
            });
        using var input = File.OpenRead(Case("selected-header"));
        var package = OoxPackage.Open(input, default);
        var part = package.GetPart("/word/unused-header.xml")!;
        _ = package.LoadXml(part, default);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => package.LoadXml(part, cancellation.Token));
        TestAssert.Throws<OperationCanceledException>(() => new DocxReader().Read(package, _ => { }, cancellation.Token, OoxPdfDocxMarkupMode.Final));
    }

    public static void DocxStoryOwnershipRetainsActivePartAndPackageWarnings()
    {
        foreach (bool selected in new[] { false, true })
            Mutate(selected ? "selected-header" : "unselected-header", (parts, xml) => {
                var header = Xml(parts, "word/unused-header.xml");
                XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                header.Root!.Add(new XAttribute(XNamespace.Xmlns+"unknown", "urn:story-owner-control"), new XAttribute(mc+"MustUnderstand", "unknown"));
                parts["word/unused-header.xml"] = Bytes(header);
                if (selected)
                {
                    var reference = new XElement(xml.Descendants(W+"headerReference").Single());
                    reference.SetAttributeValue(W+"type", "first"); xml.Descendants(W+"sectPr").Last().AddFirst(reference);
                }
            }, path => TestAssert.Equal(selected ? 1 : 0, Render(path, OoxPdfDocxMarkupMode.Final).Diagnostics.Count(d => d.Id == "OOXML_MUST_UNDERSTAND")));
        Mutate("baseline", (parts, _) => {
            parts["word/vbaProject.bin"] = [1, 2, 3];
            var types = Xml(parts, "[Content_Types].xml");
            types.Root!.Add(new XElement(types.Root.Name.Namespace+"Override", new XAttribute("PartName", "/word/vbaProject.bin"), new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")));
            parts["[Content_Types].xml"] = Bytes(types);
        }, path => TestAssert.Equal(1, Render(path, OoxPdfDocxMarkupMode.Final).Diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_MACRO")));
    }

    public static void DocxStoryOwnershipDiagnosesDrawingFallbacksWithLocalProvenance()
    {
        foreach (var mode in Modes)
        foreach (string story in new[] { "header", "footer" })
        foreach (bool selected in new[] { false, true })
            Mutate("uncached-selected-header", (parts, xml) => {
                SetDrawingStory(parts, xml, story, selected, external: true, unsupported: true);
            }, path => {
                var actual = Render(path, mode);
                foreach (string id in new[] { "DOCX_UNSUPPORTED_EXTERNAL_IMAGE", "DOCX_UNSUPPORTED_FLOATING_DRAWING", "DOCX_UNSUPPORTED_VML" })
                {
                    var warnings = actual.Diagnostics.Where(d => d.Id == id).ToArray();
                    TestAssert.Equal(selected ? 1 : 0, warnings.Length);
                    if (selected)
                    {
                        TestAssert.Equal("/word/unused-header.xml", warnings[0].PartName);
                        TestAssert.Equal("Ignored", warnings[0].Fallback);
                        TestAssert.Equal(OoxPdfSeverity.Warning, warnings[0].Severity);
                    }
                }
                string text = DocxPageFootnoteNumberingTests.DecodeText(actual.Pages.Single());
                TestAssert.Contains("Visible body", text);
                TestAssert.Equal(selected, text.Contains("Story neighbor", StringComparison.Ordinal));
                TestAssert.Equal(0, actual.Pages.Sum(page => page.Images.Count));
            });
    }

    public static void DocxStoryOwnershipKeepsDrawingRelationshipIdsLocal()
    {
        foreach (var mode in Modes)
        foreach (string story in new[] { "header", "footer" })
            Mutate("uncached-selected-header", (parts, xml) => {
                SetDrawingStory(parts, xml, story, selected: true, external: false, unsupported: false);
            }, path => {
                var actual = Render(path, mode);
                TestAssert.Equal(0, actual.Diagnostics.Count);
                TestAssert.Equal(2, actual.Pages.Sum(page => page.Images.Count));
                TestAssert.Contains("Story neighbor", DocxPageFootnoteNumberingTests.DecodeText(actual.Pages.Single()));
            });
    }

    public static void DocxStoryOwnershipDrawingDiagnosticsObserveCancellation()
    {
        Mutate("uncached-selected-header", (parts, xml) => {
            SetDrawingStory(parts, xml, "header", selected: true, external: true, unsupported: true);
        }, path => {
            using var input = File.OpenRead(path);
            var package = OoxPackage.Open(input, default);
            using var cancellation = new CancellationTokenSource();
            int warnings = 0;
            var error = TestAssert.Throws<OperationCanceledException>(() => new DocxReader().Read(package, diagnostic => {
                if (diagnostic.Id == "DOCX_UNSUPPORTED_FLOATING_DRAWING")
                {
                    warnings++;
                    cancellation.Cancel();
                }
            }, cancellation.Token, OoxPdfDocxMarkupMode.Final));
            TestAssert.Equal(1, warnings);
            TestAssert.Equal(cancellation.Token, error.CancellationToken);
        });
    }

    private static void SetDrawingStory(Dictionary<string, byte[]> parts, XDocument document,
        string story, bool selected, bool external, bool unsupported)
    {
        var reference = document.Descendants(W+"headerReference").Single();
        reference.Name = W+(story+"Reference");
        if (!selected) reference.Remove();
        var relationships = Xml(parts, "word/_rels/document.xml.rels");
        relationships.Root!.Elements().Single(e => (string?)e.Attribute("Id") == "reviewHeader")
            .SetAttributeValue("Type", R.NamespaceName+"/"+story);
        // The same ID in the body has the opposite target kind. A story must
        // use its own relationship table for both image reading and diagnostics.
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        relationships.Root.Add(new XElement(rel+"Relationship", new XAttribute("Id", "localImage"),
            new XAttribute("Type", R.NamespaceName+"/image"),
            new XAttribute("Target", external ? "media/pixel.png" : "https://example.invalid/body.png"),
            external ? null : new XAttribute("TargetMode", "External")));
        parts["word/_rels/document.xml.rels"] = Bytes(relationships);
        parts["word/_rels/unused-header.xml.rels"] = Bytes(new XDocument(new XElement(rel+"Relationships",
            new XElement(rel+"Relationship", new XAttribute("Id", "localImage"),
                new XAttribute("Type", R.NamespaceName+"/image"),
                new XAttribute("Target", external ? "https://example.invalid/story.png" : "media/pixel.png"),
                external ? new XAttribute("TargetMode", "External") : null))));
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var paragraph = new XElement(W+"p", new XElement(W+"r", new XElement(W+"t", "Story neighbor")),
            new XElement(W+"r", new XElement(W+"drawing", new XElement(wp+"inline",
                new XElement(wp+"extent", new XAttribute("cx", 127000), new XAttribute("cy", 127000)),
                new XElement(a+"graphic", new XElement(a+"graphicData", new XElement(a+"blip",
                    new XAttribute(R+(external ? "link" : "embed"), "localImage"))))))));
        if (unsupported) paragraph.Add(new XElement(W+"r",
            new XElement(W+"drawing", new XElement(wp+"anchor")),
            new XElement(W+"pict", new XElement(v+"rect"))));
        parts["word/unused-header.xml"] = Bytes(new XDocument(new XElement(W+(story == "header" ? "hdr" : "ftr"),
            paragraph, new XElement(paragraph))));
        parts["word/media/pixel.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0]);
        var types = Xml(parts, "[Content_Types].xml");
        types.Root!.Elements().Single(e => (string?)e.Attribute("PartName") == "/word/unused-header.xml")
            .SetAttributeValue("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml."+story+"+xml");
        types.Root.Add(new XElement(types.Root.Name.Namespace+"Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")));
        parts["[Content_Types].xml"] = Bytes(types);
    }

    private static (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(string path, OoxPdfDocxMarkupMode mode)
    {
        using var input = File.OpenRead(path);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var document = new DocxReader().Read(OoxPackage.Open(input, default), diagnostics.Add, default, mode);
        var pages = new DocxRenderer(new TestFaceFontResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, diagnostics.Add, default).ToArray();
        return (pages, diagnostics);
    }
    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "StoryOwnership", id+".docx"));
    private static XDocument Xml(Dictionary<string, byte[]> parts, string name) => XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
    private static byte[] Bytes(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString());
    private static void Mutate(string id, Action<Dictionary<string, byte[]>, XDocument> change, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(Case(id));
        var parts = archive.Entries.ToDictionary(e => e.FullName, e => { using var stream=e.Open(); using var bytes=new MemoryStream(); stream.CopyTo(bytes); return bytes.ToArray(); });
        var xml = Xml(parts, "word/document.xml"); change(parts, xml); parts["word/document.xml"] = Bytes(xml);
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(path); } finally { File.Delete(path); }
    }
}
