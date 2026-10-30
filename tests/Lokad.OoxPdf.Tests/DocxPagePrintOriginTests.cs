using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPagePrintOriginTests
{
    public static void DocxContinuationPagePaintIgnoresOtherPageFontSizes()
    {
        foreach (bool final in new[] { true, false })
        foreach (string family in new[] { "composed", "bare", "header" })
        {
            double? continuationY = null;
            double? noteY = null;
            foreach (int firstSize in family == "composed" ? new[] { 11, 18, 24 } : new[] { 11, 24 })
            {
                string id = family == "composed" ? $"page-origin-first{firstSize}-continuation11" : $"page-origin-{family}-first{firstSize}";
                var source = DocxTests.ReadDocx(Case(id + ".docx"), final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
                var pages = Renderer(new TestFaceFontResolver(), final).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, pages.Length);
                var body = Anchor(pages, "Public continuation body.");
                var note = Anchor(pages, "All quantities");
                TestAssert.Equal(2, body.Page);
                TestAssert.Equal(2, note.Page);
                if (continuationY is { } expected) Near(expected, body.Y, .0015d, "A forced continuation cannot inherit another page's font metrics.");
                if (noteY is { } expectedNote) Near(expectedNote, note.Y, .0015d, "Bottom-anchored notes retain their own printed frame.");
                continuationY = body.Y;
                noteY = note.Y;
            }
        }
    }

    public static void DocxPagePrintOriginsMatchIndependentWordViews()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent Word origins require the source-declared installed Windows fonts.");
        using var golden = JsonDocument.Parse(File.ReadAllText(Case("page-print-origins-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Case(row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var source = DocxTests.ReadDocx(input, final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
            var pages = Renderer(null, final).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            foreach (var expected in row.GetProperty("WordAnchors").EnumerateArray())
            {
                string text = expected.GetProperty("Text").GetString()!;
                var actual = Anchor(pages, text);
                TestAssert.Equal(expected.GetProperty("Page").GetInt32(), actual.Page);
                Near(expected.GetProperty("X").GetDouble(), actual.X, 1d, "Independent Word origin X: " + text);
                Near(expected.GetProperty("Y").GetDouble(), actual.Y, 1d, "Independent Word origin Y: " + text);
                Near(expected.GetProperty("Size").GetDouble(), actual.Size, .1d, "Independent Word font size: " + text);
            }
        }
    }

    public static void DocxPagePrintOriginsKeepInspectionLinksAndNoteTargetsTogether()
    {
        double? targetTop = null;
        double? noteTop = null;
        PdfLinkAnnotation? continuationLink = null;
        foreach (int firstSize in new[] { 11, 18, 24 })
        {
            var source = WithContinuationBookmark(Case($"page-origin-first{firstSize}-continuation11.docx"));
            var renderer = Renderer(new TestFaceFontResolver(), final: false);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var snapshot = renderer.InspectTextEmission(source);
            TestAssert.Equal(2, pages.Length);
            var body = Anchor(pages, "Public continuation body.");
            var note = Anchor(pages, "All quantities");
            var emitted = snapshot.Lines.Single(line => line.PageIndex == 1 && !line.IsStaticStory && line.StoryKind == "Body")
                .Segments.First(segment => segment.Role == "Text" && !segment.IsTerminalLineSpace);
            Near(pages[1].Height - body.Y, emitted.BaselineY, .0015d, "Inspection and painted continuation must share the owning page's origin.");

            var toContinuation = pages[0].Annotations.Single(annotation => annotation.Destination?.PageIndex == 1);
            double top = toContinuation.Destination!.Value.Top!.Value;
            Near(pages[1].Height - body.Y, top, 2d * body.Size + 3d, "Bookmark target must surround actual continuation paint.");
            if (targetTop is { } previousTop) Near(previousTop, top, .0015d, "Continuation bookmark targets cannot inherit page-one font size.");
            targetTop = top;

            var bodyLink = pages[1].Annotations.Single(annotation => annotation.Destination?.PageIndex == 0);
            var noteLink = pages[1].Annotations.Single(annotation => annotation.Destination?.PageIndex == 1);
            double actualNoteTop = noteLink.Destination!.Value.Top!.Value;
            Near(pages[1].Height - note.Y, actualNoteTop, 2d * note.Size + 3d, "Note links must target the painted note.");
            if (noteTop is { } previousNote) Near(previousNote, actualNoteTop, .0015d, "Notes retain their independent frame.");
            noteTop = actualNoteTop;
            if (continuationLink is { } prior)
            {
                Near(prior.X, bodyLink.X, .0015d, "Hyperlink X remains source-owned.");
                Near(prior.Y, bodyLink.Y, .0015d, "Hyperlink Y follows continuation paint.");
                Near(prior.Width, bodyLink.Width, .0015d, "Hyperlink width retains the same source text.");
                Near(prior.Height, bodyLink.Height, .0015d, "Hyperlink height follows the same line.");
            }
            continuationLink = bodyLink;
        }
    }

    public static void DocxExactContinuationPrintOriginKeepsItsAuthoredSlot()
    {
        double? previous = null;
        foreach (int firstSize in new[] { 11, 18, 24 })
        {
            var source = WithContinuationBookmark(Case($"page-origin-first{firstSize}-continuation11.docx"), exact: true);
            var fonts = new TestFaceFontResolver();
            var printed = Renderer(fonts, final: false).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var nominal = new DocxRenderer(fonts, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, printed.Length);
            TestAssert.Equal(2, nominal.Length);
            var actual = Anchor(printed, "Public continuation body.");
            var design = Anchor(nominal, "Public continuation body.");
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source,
                DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            double expected = new DocxPrintedPageFrame(nominal[1].Height, scale).MapY(nominal[1].Height - design.Y);
            Near(expected, printed[1].Height - actual.Y, .0015d, "An exact slot retains the independently scaled source origin.");
            if (previous is { } prior) Near(prior, actual.Y, .0015d, "A different first-page font cannot translate the exact continuation slot.");
            previous = actual.Y;
        }
    }

    private static DocxDocument WithContinuationBookmark(string input, bool exact = false)
    {
        string temporary = Path.Combine(Path.GetTempPath(), "ooxpdf-page-origin-" + Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(input, temporary);
        try
        {
            File.SetAttributes(temporary, FileAttributes.Normal);
            using (var archive = ZipFile.Open(temporary, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("word/document.xml")!;
                XDocument xml;
                using (var stream = entry.Open()) xml = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var paragraphs = xml.Root!.Element(w + "body")!.Elements(w + "p").ToArray();
                if (exact)
                {
                    var spacing = paragraphs[1].Element(w + "pPr")!.Element(w + "spacing")!;
                    spacing.SetAttributeValue(w + "line", 1600);
                    spacing.SetAttributeValue(w + "lineRule", "exact");
                }
                var heading = paragraphs[0].Elements(w + "r").First(run => run.Element(w + "t") is not null);
                heading.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicContinuation"), new XElement(heading)));
                paragraphs[1].Element(w + "pPr")!.AddAfterSelf(new XElement(w + "bookmarkStart",
                    new XAttribute(w + "id", 43), new XAttribute(w + "name", "PublicContinuation")));
                paragraphs[1].Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 43)));
                entry.Delete();
                using var output = archive.CreateEntry("word/document.xml").Open();
                xml.Save(output);
            }
            return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup);
        }
        finally { File.Delete(temporary); }
    }

    private static DocxRenderer Renderer(IFontResolver? fonts, bool final) => new(fonts,
        final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup,
        final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    private static (int Page, double X, double Y, double Size) Anchor(IReadOnlyList<PdfPage> pages, string text)
    {
        var matches = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
            .GroupBy(show => show.Y).Select(group => group.OrderBy(show => show.X).ToArray())
            .Where(line => string.Concat(line.Select(show => show.Text)).StartsWith(text, StringComparison.Ordinal))
            .Select(line => (Page: index + 1, line[0].X, Y: page.Height - line[0].Y, line[0].Size))).ToArray();
        TestAssert.True(matches.Length == 1, $"Expected one painted source line starting with '{text}', found {matches.Length}.");
        return matches.Single();
    }

    private static void Near(double expected, double actual, double tolerance, string message) =>
        TestAssert.True(double.IsFinite(actual) && Math.Abs(expected - actual) < tolerance,
            message + $" Expected {expected}; actual {actual}; tolerance {tolerance}.");

    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", name));
}
