using System.Text.RegularExpressions;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNoteRevisionLaneTests
{
    public static void DocxNotePropertyBalloonsShareThePrintedPageLane()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var body = DocxInlineCompositionTests.ReadMutation("CompositionProbes/table-math-revision-cell-format",
            OoxPdfDocxMarkupMode.AllMarkup, false, false);
        double expected = DocxRenderer.ResolveWordCompatiblePrintScale(body, context);
        TestAssert.True(expected < 1d, "The independently qualified body formatting control needs a printed lane.");
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool cell in new[] { false, true })
        {
            var source = Read(kind, cell, twoRows: false);
            TestAssert.Equal(expected, DocxRenderer.ResolveWordCompatiblePrintScale(source, context));
            TestAssert.Equal(1, Renderer().InspectMarkupBalloons(source).Count);
            var unreferenced = source with { BodyElements = source.BodyElements.Select(element =>
                element is DocxParagraphElement paragraph ? paragraph with
                { Paragraph = paragraph.Paragraph with { InlineReferences = [] } } : element).ToArray() };
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(unreferenced, context));
            TestAssert.Equal(0, Renderer().InspectMarkupBalloons(unreferenced).Count);
        }
        var hidden = context with { RendersRevisionBalloons = false };
        TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(Read("endnote", false, false), hidden));
    }

    public static void DocxScaledNotePropertyBalloonsPreserveCompleteRowsAndReadableContent()
    {
        foreach (bool cell in new[] { false, true })
        {
            var source = Read("endnote", cell, twoRows: true);
            var renderer = Renderer();
            var balloon = renderer.InspectMarkupBalloons(source).Single();
            var layout = renderer.InspectLayout(source);
            var owners = layout.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.Id == "37" && note.TableRowCount != 0).ToArray();
            TestAssert.Equal(2, owners.Length);
            TestAssert.True(owners.All(note => note.TableRowCount == 1 && note.Height >= 420d),
                "Scaled note painting must admit each complete measured row in its own design frame.");
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.Equal(320, Regex.Matches(text, @"FlowToken\d{3}").Count);
            foreach (string name in new[] { "Alpha", "Bravo", "Charlie", "Delta" })
                TestAssert.Contains("Public destination " + name + ".", text);
            TestAssert.Equal(2, pages.Sum(page => Regex.Matches(page.Content, @"/Im\d+ Do").Count));
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[balloon.PageIndex], true)
                .Where(show => show.Text.Contains("FlowToken", StringComparison.Ordinal)).ToArray();
            TestAssert.True(shows.Length != 0, "The caption page must contain the source-token witness.");
            foreach (var show in shows)
            {
                double width = show.Text.EnumerateRunes().Sum(rune => (double)show.Font.Font.GetAdvanceWidth(
                    show.Font.Font.MapCodePoint(rune.Value))) * show.Size / show.Font.Font.UnitsPerEm;
                TestAssert.True(show.X + width <= balloon.X + .01d,
                    "A source token line cannot enter the printed property-caption lane.");
            }
            var final = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(final.Sum(page => page.Annotations.Count), pages.Sum(page => page.Annotations.Count));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxRenderer Renderer() => new(new PortableRenderingTests.PinnedResolver(mathFont: true),
        OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    public static void DocxScaledNoteRevisionAnchorsShareTheirPaintedRowFrame()
    {
        foreach (bool cell in new[] { false, true })
        {
            var source = Read("endnote", cell, twoRows: true);
            var renderer = Renderer();
            var balloon = renderer.InspectMarkupBalloons(source).Single();
            var page = renderer.InspectLayout(source).Pages[balloon.PageIndex];
            var row = page.PlacedRelatedStories.Single(note => note.Id == "37" && note.TableRowCount == 1);
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source,
                DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            var frame = new DocxPrintedPageFrame(page.Height, scale);
            TestAssert.True(Math.Abs(balloon.AnchorConnectorX - (row.LeftX + row.Width * .5d) * scale) < .01d,
                "The single full-width source cell/row owns the connector's printed horizontal center.");
            // The established Word profile places its anchor2.44pt below the
            // source center. This coarse frame check catches a second coordinate
            // convention; independent Office checks retain their1pt thresholds.
            TestAssert.True(Math.Abs(balloon.AnchorY - frame.MapY(row.TopY - row.Height * .5d)) < 3d,
                "The revision anchor follows its painted row, independently of the body page origin.");
        }
    }

    private static DocxDocument Read(string kind, bool cell, bool twoRows) =>
        DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-" + kind + "-table",
            OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, _) =>
            {
                var note = xml.Root!.Elements(w + kind).Single(note => note.Element(w + "tbl") is not null);
                var row = note.Element(w + "tbl")!.Element(w + "tr")!;
                if (twoRows)
                {
                    row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                        new XAttribute(w + "val", "8400"), new XAttribute(w + "hRule", "atLeast"))));
                    row.Descendants(w + "p").Single().Add(new XElement(w + "r", new XElement(w + "t",
                        " " + string.Join(" ", Enumerable.Range(0, 160).Select(i => $"FlowToken{i:000}")))));
                    var second = new XElement(row);
                    second.Descendants().Where(element => element.Name == w + "bookmarkStart" || element.Name == w + "bookmarkEnd").Remove();
                    XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                    second.Descendants(wp + "docPr").Single().SetAttributeValue("id", "7201");
                    row.AddAfterSelf(second);
                }
                string name = cell ? "tcPr" : "trPr";
                var owner = cell ? row.Element(w + "tc")! : row;
                var properties = owner.Element(w + name);
                if (properties is null) { properties = new XElement(w + name); owner.AddFirst(properties); }
                properties.Add(new XElement(w + name + "Change", new XAttribute(w + "id", "7200"),
                    new XAttribute(w + "author", "Property Reviewer"), new XElement(w + name,
                        cell ? new XElement(w + "shd", new XAttribute(w + "fill", "FFFF00")) :
                            new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "atLeast")))));
            }, partOverride: "word/" + kind + "s.xml");
}
