using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPartOwnershipTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];

    public static void DocxOptionalPartsMatchIndependentWordOwnershipControls()
    {
        foreach (var mode in Modes)
        foreach (string id in new[] { "no-styles", "linked-styles", "orphan-styles", "no-settings", "linked-settings", "orphan-settings" })
        {
            var document = DocxTests.ReadDocx(Case(id), mode);
            if (id.EndsWith("styles", StringComparison.Ordinal))
            {
                var run = document.Paragraphs.Single().Runs.Single();
                TestAssert.Equal(id == "linked-styles" ? 40d : 12d, run.FontSize);
                TestAssert.Equal(id == "linked-styles", run.ColorHex == "FF0000");
            }
            var pages = new DocxRenderer(new TestFaceFontResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(id.EndsWith("settings", StringComparison.Ordinal) ? 2 : 1, pages.Length);
            if (pages.Length == 2)
            {
                string text = DocxPageFootnoteNumberingTests.DecodeText(pages[1]);
                TestAssert.Contains("SECTION TWO", text);
                TestAssert.Equal(id != "linked-settings", text.Contains("SECOND SECTION DEFAULT"));
            }
        }
    }

    public static void DocxOptionalPartsIgnoreExternalAndDanglingOwners()
    {
        foreach (var mode in Modes)
        foreach (bool external in new[] { false, true })
            WithMutation("orphan-styles", (parts, _) => {
                var rel = new XElement(Rels(parts).Root!.Name.Namespace+"Relationship", new XAttribute("Id", "unusableStyles"),
                    new XAttribute("Type", R.NamespaceName+"/styles"),
                    new XAttribute("Target", external ? "https://example.invalid/styles.xml" : "missing-styles.xml"));
                if (external) rel.Add(new XAttribute("TargetMode", "External"));
                AddRelationship(parts, rel);
            }, path => TestAssert.Equal(12d, DocxTests.ReadDocx(path, mode).Paragraphs.Single().Runs.Single().FontSize));
    }

    public static void DocxOptionalPartsShareNumberingAndNoteOwnership()
    {
        foreach (var mode in Modes)
        foreach (bool linked in new[] { false, true })
            WithMutation("orphan-styles", (parts, xml) => {
                xml.Descendants(W+"pPr").First().Add(new XElement(W+"numPr",
                    new XElement(W+"ilvl", new XAttribute(W+"val", "0")), new XElement(W+"numId", new XAttribute(W+"val", "1"))));
                xml.Descendants(W+"p").First().Add(new XElement(W+"r", new XElement(W+"footnoteReference", new XAttribute(W+"id", "1"))));
                AddPart(parts, "numbering", "<w:numbering xmlns:w=\""+W+"\"><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>", linked);
                AddPart(parts, "footnotes", "<w:footnotes xmlns:w=\""+W+"\"><w:footnote w:id=\"1\"><w:p><w:r><w:t>Owned note</w:t></w:r></w:p></w:footnote></w:footnotes>", linked);
            }, path => {
                var document = DocxTests.ReadDocx(path, mode);
                TestAssert.Equal(linked, document.Paragraphs.Single().ListLabel is not null);
                TestAssert.Equal(linked, document.RelatedStories.Any(story => story.Kind == DocxRelatedStoryKind.Footnote && story.Id == "1"));
            });
    }

    public static void DocxOptionalPartsWarnOnlyForSelectedMarkupRequirements()
    {
        foreach (bool linked in new[] { false, true })
            WithMutation("orphan-styles", (parts, _) => {
                var style = XDocument.Parse(Encoding.UTF8.GetString(parts["word/detached-styles.xml"]));
                style.Root!.Add(new XAttribute(XNamespace.Xmlns+"unknown", "urn:optional-part-control"),
                    new XAttribute(XNamespace.Get("http://schemas.openxmlformats.org/markup-compatibility/2006")+"MustUnderstand", "unknown"));
                parts["word/detached-styles.xml"] = Encoding.UTF8.GetBytes(style.ToString());
                if (linked) AddRelationship(parts, new XElement(Rels(parts).Root!.Name.Namespace+"Relationship",
                    new XAttribute("Id", "selectedStyles"), new XAttribute("Type", R.NamespaceName+"/styles"), new XAttribute("Target", "detached-styles.xml")));
            }, path => {
                using var input = File.OpenRead(path);
                var diagnostics = new List<OoxPdfDiagnostic>();
                _ = new DocxReader().Read(Lokad.OoxPdf.Ooxml.OoxPackage.Open(input, default), diagnostics.Add, default, OoxPdfDocxMarkupMode.Final);
                TestAssert.Equal(linked ? 1 : 0, diagnostics.Count(diagnostic => diagnostic.Id == "OOXML_MUST_UNDERSTAND"));
            });
    }

    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", "optional-part-"+id+".docx"));
    private static XDocument Rels(Dictionary<string, byte[]> parts) => XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
    private static void AddRelationship(Dictionary<string, byte[]> parts, XElement relationship)
    {
        var rels = Rels(parts); rels.Root!.Add(relationship);
        parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
    }
    private static void AddPart(Dictionary<string, byte[]> parts, string kind, string xml, bool linked)
    {
        string path = "word/owned-"+kind+".xml";
        parts[path] = Encoding.UTF8.GetBytes(xml);
        var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        types.Root!.Add(new XElement(types.Root.Name.Namespace+"Override", new XAttribute("PartName", "/"+path),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml."+kind+"+xml")));
        parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString());
        if (linked) AddRelationship(parts, new XElement(Rels(parts).Root!.Name.Namespace+"Relationship", new XAttribute("Id", "owned"+kind),
            new XAttribute("Type", R.NamespaceName+"/"+kind), new XAttribute("Target", "owned-"+kind+".xml")));
    }
    private static void WithMutation(string id, Action<Dictionary<string, byte[]>, XDocument> mutate, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(Case(id));
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var bytes = new MemoryStream(); source.CopyTo(bytes); return bytes.ToArray();
        });
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        mutate(parts, xml); parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(path); } finally { File.Delete(path); }
    }
}
