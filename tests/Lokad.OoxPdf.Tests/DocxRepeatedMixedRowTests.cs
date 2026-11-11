using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRepeatedMixedRowTests
{
    public static void DocxMixedTableRowsKeepCaptionInkAcrossPages()
    {
        const string caption = "Repeat-header-letter";
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool repeated in new[] { false, true })
        {
            double? originalGap = null;
            foreach (bool longTable in new[] { false, true })
            {
                var source = ReadMixedTable("Composition/letter", mode, caption,
                    (table, w) =>
                    {
                        if (repeated) table.Element(w + "tr")!.Element(w + "trPr")!.AddFirst(new XElement(w + "tblHeader"));
                        if (longTable)
                            for (int index = 1; index <= 54; index++)
                                table.Add(new XElement(w + "tr", new XElement(w + "trPr",
                                    new XElement(w + "trHeight", new XAttribute(w + "val", "480"), new XAttribute(w + "hRule", "atLeast")),
                                    new XElement(w + "cantSplit")), new XElement(w + "tc", new XElement(w + "p",
                                        new XElement(w + "r", new XElement(w + "t", "Body-row-" + index.ToString("D3", CultureInfo.InvariantCulture)))))));
                    });
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                if (longTable) TestAssert.True(pages.Length > 1, "The source table must cross a physical page.");
                int headers = 0;
                foreach (var page in pages)
                {
                    var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
                    if (!shows.Any(show => show.Text.Contains("Repeat", StringComparison.Ordinal))) continue;
                    headers++;
                    DocxInlineCompositionTests.AssertSourceLabelsPaint(page, mode, [caption],
                        $"mixed cell/{mathFont}/{repeated}/{longTable}");
                    var picture = DocxInlineCompositionTests.ReadPictureBounds(page).Single().Bounds;
                    double baseline = shows.Single(show => show.Text.Contains("Repeat", StringComparison.Ordinal)).Y;
                    double gap = baseline - picture.Y - picture.Height;
                    originalGap ??= gap;
                    TestAssert.True(Math.Abs(gap - originalGap.Value) < .01d,
                        "A complete cell's relative paragraph/picture frame must not depend on later table rows or page placement.");
                }
                TestAssert.True(headers >= (longTable && repeated ? 2 : 1), "The source header survives every required placement.");
            }
        }
    }

    public static void DocxKeptMixedCellRetainsInkAndFollowingParagraphAfterPageAdmission()
    {
        foreach (string fixture in new[] { "letter", "specification" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string caption = "Kept-row-caption-" + fixture;
            string following = "After-kept-row-" + fixture;
            double? originalGap = null;
            foreach (bool deferred in new[] { false, true })
            {
                var source = ReadMixedTable("Composition/" + fixture, mode, caption, (table, w) =>
                {
                    table.Element(w + "tr")!.Element(w + "trPr")!.Add(new XElement(w + "trHeight",
                        new XAttribute(w + "val", "3600"), new XAttribute(w + "hRule", "atLeast")));
                    table.AddAfterSelf(Paragraph(w, following));
                    if (!deferred) return;
                    var filler = new XElement(w + "tbl", new XElement(table.Element(w + "tblPr")!),
                        new XElement(table.Element(w + "tblGrid")!), new XElement(w + "tr",
                            new XElement(w + "trPr", new XElement(w + "cantSplit"), new XElement(w + "trHeight",
                                new XAttribute(w + "val", "8200"), new XAttribute(w + "hRule", "exact"))),
                            new XElement(w + "tc", new XElement(w + "p", new XElement(w + "r",
                                new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "20"))),
                                new XElement(w + "t", "Page-filler-" + fixture))))));
                    table.AddBeforeSelf(filler, Paragraph(w, "Before-kept-row-" + fixture));
                });
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                var painted = pages.Select(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray()).ToArray();
                int[] owners = Enumerable.Range(0, pages.Length).Where(index =>
                    string.Concat(painted[index].Select(show => show.Text)).Contains(caption, StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(1, owners.Length);
                int owner = owners[0];
                TestAssert.Equal(deferred ? 1 : 0, owner);
                if (deferred) TestAssert.Contains("Page-filler-" + fixture,
                    string.Concat(painted[0].Select(show => show.Text)));
                DocxInlineCompositionTests.AssertSourceLabelsPaint(pages[owner], mode, [caption, following],
                    $"kept row/{fixture}/{mathFont}/{deferred}");
                double captionY = LabelBaseline(caption);
                double followingY = LabelBaseline(following);
                string cellText = string.Concat(painted[owner].Where(show => show.Y <= captionY && show.Y > followingY)
                    .SelectMany(show => show.Text).Where(character => !char.IsWhiteSpace(character)));
                TestAssert.Contains("a+b=c", cellText);
                var picture = DocxInlineCompositionTests.ReadPictureBounds(pages[owner]).Single().Bounds;
                TestAssert.True(picture.Y > followingY, "The following paragraph clears the kept picture.");
                double gap = captionY - picture.Y - picture.Height;
                originalGap ??= gap;
                TestAssert.True(Math.Abs(gap - originalGap.Value) < .01d,
                    "Page admission preserves the complete cell's relative caption/picture frame.");

                double LabelBaseline(string label)
                {
                    string text = string.Concat(painted[owner].Select(show => show.Text));
                    int start = text.IndexOf(label, StringComparison.Ordinal);
                    TestAssert.True(start >= 0, "The complete source label must survive split PDF text operations.");
                    int offset = 0;
                    foreach (var show in painted[owner])
                    {
                        if (start >= offset && start < offset + show.Text.Length) return show.Y;
                        offset += show.Text.Length;
                    }
                    throw new InvalidOperationException("The source label has no painted baseline.");
                }
            }
        }
    }

    private static DocxDocument ReadMixedTable(string fixture, OoxPdfDocxMarkupMode mode, string caption,
        Action<XElement, XNamespace> configure) =>
        DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
        {
            var paragraph = xml.Root!.Element(w + "body")!.Elements(w + "p").Single(p =>
                p.Descendants(m + "oMath").Any() && p.Descendants(w + "drawing").Any());
            // Both portable faces own ordinary literal glyphs. Mathematical
            // shaping and Office geometry are qualified independently.
            foreach (var run in paragraph.Descendants(m + "r"))
            {
                var properties = run.Element(m + "rPr");
                if (properties is null) { properties = new XElement(m + "rPr"); run.AddFirst(properties); }
                properties.Elements(m + "nor").Remove();
                properties.Add(new XElement(m + "nor", new XAttribute(m + "val", "1")));
            }
            var cell = new XElement(w + "tc", new XElement(w + "tcPr"), Paragraph(w, caption));
            var row = new XElement(w + "tr", new XElement(w + "trPr", new XElement(w + "cantSplit")), cell);
            var table = new XElement(w + "tbl", new XElement(w + "tblPr",
                new XElement(w + "tblLayout", new XAttribute(w + "type", "fixed")),
                new XElement(w + "tblCellMar", new[] { "top", "left", "bottom", "right" }.Select(edge =>
                    new XElement(w + edge, new XAttribute(w + "w", "0"), new XAttribute(w + "type", "dxa"))))),
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))), row);
            paragraph.AddBeforeSelf(table); paragraph.Remove(); cell.Add(paragraph);
            configure(table, w);
        });

    private static XElement Paragraph(XNamespace w, string text) => new(w + "p",
        new XElement(w + "pPr", new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"))),
        new XElement(w + "r", new XElement(w + "t", text)));
}
