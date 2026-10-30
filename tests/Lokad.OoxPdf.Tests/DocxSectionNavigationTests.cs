using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxSectionNavigationTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void DocxSectionNavigationKeepsIndependentWordControlContent()
    {
        foreach (var mode in Modes)
        foreach (string id in new[] { "single-title", "two-title", "two-default", "bookmark-before", "bookmark-after", "bookmark-empty" })
        {
            var document = DocxTests.ReadDocx(Case(id), mode);
            PdfPage[] pages = Renderer(mode).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(id == "single-title" ? 1 : 2, pages.Length);
            if (id.StartsWith("bookmark-", StringComparison.Ordinal))
            {
                var link = pages[0].Annotations.Single();
                TestAssert.True(link.Destination is { PageIndex: 1 }, "All three independently exported bookmark controls must retain their incoming GoTo.");
                TestAssert.Contains("TARGET PARAGRAPH", Text(pages[1]));
            }
            else
            {
                string header = id == "two-default" ? "SECOND SECTION DEFAULT" : "SECOND SECTION FIRST";
                TestAssert.Contains(header, Text(pages[^1]));
                if (pages.Length == 2) TestAssert.Contains("FIRST SECTION DEFAULT", Text(pages[0]));
                TestAssert.Contains("SECTION TWO", Text(pages[^1]));
            }
        }
    }

    public static void DocxSectionNavigationSelectsFirstThenDefaultAndEvenStories()
    {
        foreach (var mode in Modes)
        WithMutation("two-title", (parts, xml) => {
            XElement body = xml.Root!.Element(W+"body")!;
            XElement final = body.Element(W+"sectPr")!;
            final.AddBeforeSelf(Break(), Paragraph("CONTINUATION"), Break(), Paragraph("EVEN CONTINUATION"));
            AddFooter(parts, final, "default", "FOOTER DEFAULT");
            AddFooter(parts, final, "first", "FOOTER FIRST");
            var settings = XDocument.Parse(Encoding.UTF8.GetString(parts["word/settings.xml"]));
            settings.Root!.Elements(W+"evenAndOddHeaders").Remove();
            settings.Root.Add(new XElement(W+"evenAndOddHeaders"));
            parts["word/settings.xml"] = Encoding.UTF8.GetBytes(settings.ToString());
        }, path => {
            var document = DocxTests.ReadDocx(path, mode);
            var renderer = Renderer(mode);
            PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(4, pages.Length);
            TestAssert.Contains("SECOND SECTION FIRST", Text(pages[1]));
            TestAssert.Contains("FOOTER FIRST", Text(pages[1]));
            TestAssert.Contains("SECOND SECTION DEFAULT", Text(pages[2]));
            TestAssert.Contains("FOOTER DEFAULT", Text(pages[2]));
            TestAssert.True(!Text(pages[3]).Contains("SECOND SECTION") && !Text(pages[3]).Contains("FOOTER"),
                "An undefined even variant remains empty; the first variant applies only at the owning section's first page.");
        });
    }

    public static void DocxSectionNavigationDoesNotBackfillAnUndefinedFirstHeader()
    {
        foreach (var mode in Modes)
        WithMutation("two-title", (_, xml) => {
            xml.Root!.Element(W+"body")!.Element(W+"sectPr")!.Elements(W+"headerReference")
                .Where(e => (string?)e.Attribute(W+"type") == "first").Remove();
        }, path => {
            PdfPage[] pages = Renderer(mode).RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("FIRST SECTION DEFAULT", Text(pages[0]));
            TestAssert.True(!Text(pages[1]).Contains("SECOND SECTION"), "An undefined first variant must not select the default header.");
            TestAssert.Contains("SECTION TWO", Text(pages[1]));
        });
    }

    public static void DocxParagraphEndBookmarksFollowTheFinalWrappedPage()
    {
        foreach (var mode in Modes)
        WithMutation("bookmark-after", (_, xml) => {
            XElement body = xml.Root!.Element(W+"body")!;
            body.Elements(W+"p").Last().Descendants(W+"t").Single().Value = string.Join(" ", Enumerable.Repeat("TARGET PARAGRAPH", 90));
            XElement section = body.Element(W+"sectPr")!;
            section.Element(W+"pgSz")!.SetAttributeValue(W+"h", "6000");
            section.Element(W+"pgSz")!.SetAttributeValue(W+"w", "6000");
        }, path => CheckFinalTextDestination(path, mode, expectMultiplePages: true));
    }

    public static void DocxParagraphEndBookmarksIgnoreHiddenTrailingText()
    {
        foreach (var mode in Modes)
        WithMutation("bookmark-after", (_, xml) => {
            var anchor = xml.Descendants(W+"bookmarkStart").Single();
            anchor.AddBeforeSelf(new XElement(W+"r", new XElement(W+"rPr", new XElement(W+"vanish")), new XElement(W+"t", "HIDDEN TRAILING TEXT")));
        }, path => CheckFinalTextDestination(path, mode, expectMultiplePages: false));
    }

    public static void DocxParagraphEndBookmarksPreserveMissingTargetFallback()
    {
        WithMutation("bookmark-after", (_, xml) => xml.Descendants(W+"bookmarkStart").Remove(), path => {
            PdfPage[] pages = Renderer(OoxPdfDocxMarkupMode.Final).RenderPagesAsync(DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(0, pages.Sum(page => page.Annotations.Count));
            TestAssert.Contains("TARGET PARAGRAPH", Text(pages[1]));
        });
    }

    private static void CheckFinalTextDestination(string path, OoxPdfDocxMarkupMode mode, bool expectMultiplePages)
    {
        var document = DocxTests.ReadDocx(path, mode);
        var renderer = Renderer(mode);
        PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
        if (expectMultiplePages) TestAssert.True(pages.Length > 2, "The target must wrap across pages.");
        PdfLinkDestination destination = pages[0].Annotations.Single().Destination!.Value;
        TestAssert.Equal(pages.Length-1, destination.PageIndex);
        var lastLine = renderer.InspectTextEmission(document).Lines.Last(line => line.ContainerStoryKind == "Body" &&
            line.Segments.Any(segment => !segment.IsTerminalLineSpace && segment.TextLength > 0));
        var lastSegment = lastLine.Segments.Last(segment => !segment.IsTerminalLineSpace && segment.TextLength > 0);
        TestAssert.True(Math.Abs(destination.Left!.Value - (lastSegment.X+lastSegment.Width-3d*lastSegment.FontSize/14d)) < .02d,
            "The end bookmark must use the final emitted endpoint and existing viewport context.");
        TestAssert.True(destination.Top > lastSegment.BaselineY && destination.Top < lastSegment.BaselineY+lastSegment.PdfFontSize*2d,
            "The destination must belong to the final line, not an earlier wrapped fragment.");
    }

    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];
    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", "section-nav-"+id+".docx"));
    private static string Text(PdfPage page) => DocxPageFootnoteNumberingTests.DecodeText(page);
    private static DocxRenderer Renderer(OoxPdfDocxMarkupMode mode) => new(new TestFaceFontResolver(), mode,
        mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
    private static XElement Paragraph(string text) => new(W+"p", new XElement(W+"r", new XElement(W+"t", text)));
    private static XElement Break() => new(W+"p", new XElement(W+"r", new XElement(W+"br", new XAttribute(W+"type", "page"))));

    private static void AddFooter(Dictionary<string,byte[]> parts, XElement section, string type, string text)
    {
        string name = "footer-"+type;
        section.AddFirst(new XElement(W+"footerReference", new XAttribute(W+"type", type), new XAttribute(R+"id", name)));
        parts["word/"+name+".xml"] = Encoding.UTF8.GetBytes(new XElement(W+"ftr", Paragraph(text)).ToString());
        var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        rels.Root!.Add(new XElement(rels.Root.Name.Namespace+"Relationship", new XAttribute("Id", name), new XAttribute("Target", name+".xml"), new XAttribute("Type", R.NamespaceName+"/footer")));
        parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
        var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        types.Root!.Add(new XElement(types.Root.Name.Namespace+"Override", new XAttribute("PartName", "/word/"+name+".xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));
        parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString());
    }

    private static void WithMutation(string id, Action<Dictionary<string,byte[]>,XDocument> mutate, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(Case(id));
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var bytes = new MemoryStream(); source.CopyTo(bytes); return bytes.ToArray();
        });
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        mutate(parts, xml);
        parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(path); } finally { File.Delete(path); }
    }
}
