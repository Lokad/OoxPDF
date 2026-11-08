using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMinimumCommentTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static void DocxMinimumCommentsKeepFinalBodyGlyphFramesAndSource()
    {
        foreach (bool arial12 in new[] { false, true })
        {
            var original = Read(false, false, arial12);
            var expected = Render(original);
            foreach (bool empty in new[] { false, true })
            {
                var source = Read(true, empty, arial12);
                TestAssert.Equal(1, source.Paragraphs[0].CommentRanges.Count);
                TestAssert.Equal(1, source.Paragraphs[0].InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment));
                var pages = Render(source);
                EqualTextFrames(expected, pages);
                TestAssert.DoesNotContain("Unused public comment sentinel", string.Concat(Shows(pages).Select(show => show.Text)));
                TestAssert.Equal(1, source.Paragraphs[0].CommentRanges.Count);
            }
        }
    }

    public static void DocxMinimumCommentsKeepParagraphPaintAndHyperlinkFrames()
    {
        var expected = Render(Read(false, false, decorate: true));
        var actual = Render(Read(true, false, decorate: true));
        EqualTextFrames(expected, actual);
        TestAssert.Equal(1, expected.Sum(page => page.Annotations.Count));
        TestAssert.True(expected.SelectMany(page => page.Annotations).SequenceEqual(actual.SelectMany(page => page.Annotations)),
            "Comment metadata must not replace the hyperlink's placed text frame.");
        // The filled paragraph rectangle is a semantic paint result, separate
        // from source run grouping and PDF font/object identities.
        var rectangles = new System.Text.RegularExpressions.Regex(@"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re");
        var a = expected.SelectMany(page => rectangles.Matches(page.Content).Select(match => match.Value)).ToArray();
        var b = actual.SelectMany(page => rectangles.Matches(page.Content).Select(match => match.Value)).ToArray();
        TestAssert.True(a.Length > 0 && a.SequenceEqual(b), "Natural-line shading and occupied paragraph slots retain their geometry.");
    }

    public static void DocxMinimumCommentsKeepRealNoteMarkerSlotsAndCompactFallback()
    {
        foreach (int minimum in new[] { 300, 400 })
        foreach (bool annotate in new[] { false, true })
        {
            var source = Read(annotate, false, note: true, minimum: minimum);
            var resolver = new PortableRenderingTests.PinnedResolver();
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(DocxFontPlan.Create(source, resolver, default),
                resolver.Resolve(new FontRequest("Calibri", false, false)), default, resolver).GetAwaiter().GetResult();
            var before = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, measurer, default);
            DocxMinimumBodyLineLeading.TryCreateLayout(source, before, out var after);
            var a = before.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().First();
            var b = after.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().First();
            TestAssert.Equal(1, source.Paragraphs[0].InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Footnote));
            if (minimum == 300)
                TestAssert.True(ReferenceEquals(a, b), "A compact actual note-marker slot retains its established fallback, with or without comments.");
            else
            {
                TestAssert.True(a.BaselineY > b.BaselineY, "A supported minimum slot still lowers its body baseline beside comment metadata.");
                TestAssert.Equal(a.LineHeight, b.BodyLineBoxHeightPoints);
                TestAssert.Equal(b.FirstLineInsetPoints, b.BodyLineBoxBaselineInsetPoints);
                TestAssert.True(ReferenceEquals(before.RelatedStories, after.RelatedStories), "The minimum correction does not replace note layouts.");
            }
        }
    }

    public static void DocxMinimumCommentsMatchIndependentWordFinalOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent comment/minimum origins require installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("minimum-comment-office.json")));
        TestAssert.Equal(gold.RootElement.GetProperty("OriginalInputSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Case("minimum-navigation-ordinary11-control.docx")))));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = Read(row.GetProperty("Annotated").GetBoolean(), row.GetProperty("Empty").GetBoolean(),
                row.GetProperty("Arial12").GetBoolean());
            var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var first = Shows(pages).First(show => !string.IsNullOrWhiteSpace(show.Text));
            TestAssert.True(first.Text.StartsWith("Public", StringComparison.Ordinal), "The first painted glyph belongs to the first source paragraph.");
            TestAssert.True(Math.Abs(first.X - row.GetProperty("X").GetDouble()) <= 1d &&
                Math.Abs(pages[0].Height - first.Y - row.GetProperty("Y").GetDouble()) <= 1d &&
                Math.Abs(first.Size - row.GetProperty("Size").GetDouble()) <= .1d,
                "Comment metadata retains the independent one-point Word origin and .1-point glyph-size gates.");
        }
    }

    private static void EqualTextFrames(PdfPage[] expected, PdfPage[] actual)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        static (string Text, double X, double Y, string Sizes)[] Lines(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).GroupBy(show => show.Y)
                .Select(group => (string.Concat(group.Select(show => show.Text)), group.Min(show => show.X), group.Key,
                    string.Join(",", group.Select(show => show.Size).Distinct().Order()))).ToArray();
        for (int page = 0; page < expected.Length; page++)
        {
            var a = Lines(expected[page]); var b = Lines(actual[page]);
            TestAssert.Equal(a.Length, b.Length);
            for (int i = 0; i < a.Length; i++)
                TestAssert.True(a[i].Text == b[i].Text && Math.Abs(a[i].X - b[i].X) < .000001d &&
                    Math.Abs(a[i].Y - b[i].Y) < .000001d && a[i].Sizes == b[i].Sizes,
                    "Source annotations preserve measured lines independently of PDF operation grouping.");
        }
    }

    private static IEnumerable<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)> Shows(IEnumerable<PdfPage> pages) =>
        pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true));

    private static PdfPage[] Render(DocxDocument document) => new DocxRenderer(new PortableRenderingTests.PinnedResolver(),
        OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
        .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();

    private static DocxDocument Read(bool annotate, bool empty, bool arial12 = false, bool decorate = false, bool note = false, int minimum = 400)
    {
        var parts = new Dictionary<string, string>();
        using (var archive = ZipFile.OpenRead(Case("minimum-navigation-ordinary11-control.docx")))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]);
        var paragraph = xml.Descendants(W + "body").Elements(W + "p").First();
        paragraph.Element(W + "pPr")!.Element(W + "spacing")!.SetAttributeValue(W + "line", minimum);
        if (arial12)
            foreach (var run in xml.Descendants(W + "r"))
            {
                var properties = run.Element(W + "rPr");
                if (properties is null) { properties = new XElement(W + "rPr"); run.AddFirst(properties); }
                properties.SetElementValue(W + "rFonts", null);
                properties.Add(new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial")));
                properties.SetElementValue(W + "sz", null);
                properties.Add(new XElement(W + "sz", new XAttribute(W + "val", "24")));
            }
        var types = XDocument.Parse(parts["[Content_Types].xml"]);
        var relationships = XDocument.Parse(parts["word/_rels/document.xml.rels"]);
        void AddPart(string name, string kind, XElement content)
        {
            parts["word/" + name + ".xml"] = content.ToString();
            types.Root!.Add(new XElement(Ct + "Override", new XAttribute("PartName", "/word/" + name + ".xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + name + "+xml")));
            relationships.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", name),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind), new XAttribute("Target", name + ".xml")));
        }
        if (note)
        {
            paragraph.Add(new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Calibri")),
                new XElement(W + "sz", new XAttribute(W + "val", "22")), new XElement(W + "vertAlign", new XAttribute(W + "val", "superscript"))),
                new XElement(W + "footnoteReference", new XAttribute(W + "id", "5"))));
            AddPart("footnotes", "footnotes", new XElement(W + "footnotes", new XElement(W + "footnote", new XAttribute(W + "id", "5"),
                new XElement(W + "p", new XElement(W + "r", new XElement(W + "footnoteRef"), new XElement(W + "t", "Public note."))))));
        }
        if (decorate)
        {
            paragraph.Element(W + "pPr")!.Add(new XElement(W + "shd", new XAttribute(W + "fill", "D8E4F0")));
            var run = paragraph.Elements(W + "r").First();
            run.ReplaceWith(new XElement(W + "hyperlink", new XAttribute(R + "id", "publicLink"), new XElement(run)));
            relationships.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "publicLink"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                new XAttribute("Target", "https://example.invalid/minimum-comment"), new XAttribute("TargetMode", "External")));
        }
        if (annotate)
        {
            var start = new XElement(W + "commentRangeStart", new XAttribute(W + "id", "42"));
            paragraph.Elements().First(child => child.Name != W + "pPr").AddBeforeSelf(start);
            var end = new XElement(W + "commentRangeEnd", new XAttribute(W + "id", "42"));
            var marker = new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "42")));
            if (empty) { start.AddAfterSelf(end); end.AddAfterSelf(marker); }
            else paragraph.Add(end, marker);
            AddPart("comments", "comments", new XElement(W + "comments", new XElement(W + "comment", new XAttribute(W + "id", "42"),
                new XAttribute(W + "author", "Public synthetic"), new XElement(W + "p", new XElement(W + "r",
                    new XElement(W + "rPr", new XElement(W + "b"), new XElement(W + "sz", new XAttribute(W + "val", "72"))),
                    new XElement(W + "t", "Unused public comment sentinel"))))));
        }
        parts["word/document.xml"] = xml.ToString(); parts["[Content_Types].xml"] = types.ToString();
        parts["word/_rels/document.xml.rels"] = relationships.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final); }
        finally { File.Delete(input); }
    }

    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes", name));
}
