using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNoteContinuationPaintTests
{
    public static void DocxBalloonlessUnitScaleOriginsDoNotRequireBodyText()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (double? baseline in new double?[] { null, 700d })
        {
            TestAssert.Equal(0d, DocxRenderer.ResolveWordCompatibleTextYOffset(context, baseline, 792d,
                hasBalloonContent: false));
            TestAssert.Equal(0d, DocxRenderer.ResolveWordCompatibleTextYOffset(context, baseline, 792d,
                hasBalloonContent: true));
        }
        TestAssert.Equal(69.58d, DocxRenderer.ResolveWordCompatibleTextYOffset(
            context with { WordCompatiblePrintScale = .75d }, null, 792d, hasBalloonContent: true));
    }

    public static void DocxEndnotePropertyHistoryKeepsContinuationPaintInsideUnscaledClips()
    {
        foreach (bool cell in new[] { false, true })
        {
            PdfPage[] Render(OoxPdfDocxMarkupMode mode)
            {
                var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table",
                    mode, false, false, (xml, w, _) =>
                    {
                        var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                        var row = note.Element(w + "tbl")!.Element(w + "tr")!;
                        row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                            new XAttribute(w + "val", "8400"), new XAttribute(w + "hRule", "atLeast"))));
                        row.Descendants(w + "p").Single().Add(new XElement(w + "r", new XElement(w + "t",
                            " " + string.Join(" ", Enumerable.Range(0, 160).Select(i => $"FlowToken{i:000}")))));
                        var second = new XElement(row);
                        second.Descendants().Where(element => element.Name == w + "bookmarkStart" || element.Name == w + "bookmarkEnd").Remove();
                        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                        second.Descendants(wp + "docPr").Single().SetAttributeValue("id", "2701");
                        row.AddAfterSelf(second);
                        string name = cell ? "tcPr" : "trPr";
                        var owner = cell ? row.Element(w + "tc")! : row;
                        var properties = owner.Element(w + name);
                        if (properties is null) { properties = new XElement(w + name); owner.AddFirst(properties); }
                        properties.Add(new XElement(w + name + "Change", new XAttribute(w + "id", "6901"),
                            new XAttribute(w + "author", "Property Reviewer"), new XElement(w + name)));
                    }, partOverride: "word/endnotes.xml");
                var settings = source.PageSettings with { HeightValue = "12000" };
                source = source with { PageHeightPoints = 600d, PageSettings = settings,
                    FinalSectionBreak = source.FinalSectionBreak is { } final ? final with { PageSettings = settings } : null,
                    BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
                        ? section with { PageSettings = section.PageSettings with { HeightValue = "12000" } } : element).ToArray() };
                return new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                        : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            }
            var final = Render(OoxPdfDocxMarkupMode.Final);
            var review = Render(OoxPdfDocxMarkupMode.AllMarkup);
            TestAssert.True(final.Length >= 3 && final.Length == review.Length,
                "Property history must retain complete rows and dedicated note continuation pages.");
            foreach (var pages in new[] { final, review })
                TestAssert.Equal(320, Regex.Matches(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)), @"FlowToken\d{3}").Count);
            foreach (string anchor in new[] { "Public destination Bravo.", "Public destination Charlie.", "Public destination Delta." })
            {
                var before = Anchor(final, anchor);
                var after = Anchor(review, anchor);
                TestAssert.True(before.Page > 1 && before.Page == after.Page &&
                    Math.Abs(before.X - after.X) < .002d && Math.Abs(before.Y - after.Y) < .002d,
                    "A property-history balloon cannot displace the following note outside its clip: " + anchor);
            }
            for (int page = 0; page < final.Length; page++)
                TestAssert.True(MathBaselines(final[page]).SequenceEqual(MathBaselines(review[page])),
                    "Resolved row/cell history preserves source math placement.");
            TestAssert.True(final.SelectMany(page => page.Annotations).Select(link => link.Destination)
                .SequenceEqual(review.SelectMany(page => page.Annotations).Select(link => link.Destination)),
                "Property history retains source-owned note destinations.");
        }
    }

    public static void DocxEndnoteOnlyContinuationsKeepUnscaledTextMathPictureAndClipOrigins()
    {
        foreach (int count in new[] { 300, 420, 560 })
        {
            PdfPage[] Render(OoxPdfDocxMarkupMode mode)
            {
                var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-paragraph",
                    mode, false, false, (xml, w, _) =>
                    {
                        var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                        note.Descendants(w + "t").First(text => text.Value.Contains("Public destination Alpha.", StringComparison.Ordinal))
                            .Value += " " + string.Join(" ", Enumerable.Range(0, count).Select(i => $"FlowToken{i:000}")) + " TrailBoundarySentinel.";
                    }, partOverride: "word/endnotes.xml");
                return new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                        : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            }
            var final = Render(OoxPdfDocxMarkupMode.Final);
            var review = Render(OoxPdfDocxMarkupMode.AllMarkup);
            TestAssert.True(final.Length >= 3 && final.Length == review.Length,
                "The source must continue onto a note-only page in both unscaled modes.");
            foreach (var pages in new[] { final, review })
            {
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(count, Regex.Matches(text, @"FlowToken\d{3}").Count);
                TestAssert.Contains("TrailBoundarySentinel.", text);
                TestAssert.Equal(4, pages.Sum(page => page.Annotations.Count));
            }
            // Encoded text alone also passes on the broken renderer. These named
            // anchors, mathematical baselines and painted picture must stay in
            // their independently visible Final frames; clips must not expand.
            foreach (string anchor in new[] { "TrailBoundarySentinel.", "Public destination Bravo.",
                "Public destination Charlie.", "Public destination Delta." })
            {
                var before = Anchor(final, anchor);
                var after = Anchor(review, anchor);
                TestAssert.True(before.Page > 1 && before.Page == after.Page &&
                    Math.Abs(before.X - after.X) < .002d && Math.Abs(before.Y - after.Y) < .002d,
                    "Review must preserve the visible source anchor on its continuation page: " + anchor);
            }
            for (int page = 0; page < final.Length; page++)
            {
                TestAssert.True(MathBaselines(final[page]).SequenceEqual(MathBaselines(review[page])),
                    "Both fractions retain their painted baselines and source page.");
                TestAssert.True(ClipFrames(final[page]).SequenceEqual(ClipFrames(review[page])),
                    "Fixing paint origins must not enlarge or remove the source clips.");
            }
            TestAssert.True(final.Any(page => MathBaselines(page).Length != 0), "The source fractions must be painted.");
            TestAssert.Equal(Picture(final), Picture(review));
            TestAssert.True(final.SelectMany(page => page.Annotations).Select(link => link.Destination)
                .SequenceEqual(review.SelectMany(page => page.Annotations).Select(link => link.Destination)),
                "Note destinations retain their page and visible source frame.");
        }
    }

    private static (int Page, double X, double Y) Anchor(PdfPage[] pages, string text) => pages
        .SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
            .Where(show => show.Text.Contains(text, StringComparison.Ordinal)).Select(show => (index + 1, show.X, show.Y)))
        .Single();

    private static double[] MathBaselines(PdfPage page)
    {
        var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
            .Select(font => font.Font).ToHashSet();
        return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
            .Where(show => fonts.Contains(show.Font)).Select(show => show.Y).Distinct().Order().ToArray();
    }

    private static string[] ClipFrames(PdfPage page) => Regex.Matches(page.Content,
        @"[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ re\s+W\s+n").Select(match => match.Value).ToArray();

    private static (int Page, double Y) Picture(PdfPage[] pages) => pages.SelectMany((page, index) =>
        Regex.Matches(page.Content, @"[-\d.]+ ([-\d.]+) cm\s*/Im\d+ Do")
            .Select(match => (index + 1, double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)))).Single();
}
