using System.Xml.Linq;
using System.Reflection;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEndnoteBlockTests
{
    public static void DocxCompleteEndnoteRowDefersWithoutClippingItsContent() => Check(false);

    public static void DocxCompleteEndnoteRowsAdvanceIndependentlyOfDirectTextLines() => Check(true);

    public static void DocxOversizedEndnoteRowsDeferBeforeTheirBoundedClip() => Check(true, oversized: true);

    public static void DocxOversizedEndnoteRowsRetainHeightOnlyProgress() => Check(false, oversized: true, emptyOnly: true);

    private static void Check(bool multipleRows, bool oversized = false, bool emptyOnly = false)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table",
                mode, false, false, (xml, w, _) =>
                {
                    var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                    var row = note.Element(w + "tbl")!.Element(w + "tr")!;
                    row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                        new XAttribute(w + "val", oversized ? "12000" : "8400"), new XAttribute(w + "hRule", "atLeast"))));
                    row.Descendants(w + "p").Single().Add(new XElement(w + "r", new XElement(w + "t",
                        " " + string.Join(" ", Enumerable.Range(0, 160).Select(i => $"FlowToken{i:000}")) + " TrailBoundarySentinel.")));
                    if (multipleRows)
                    {
                        var second = new XElement(row);
                        second.Descendants().Where(element => element.Name == w + "bookmarkStart" || element.Name == w + "bookmarkEnd").Remove();
                        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                        second.Descendants(wp + "docPr").Single().SetAttributeValue("id", "2701");
                        row.AddAfterSelf(second);
                    }
                }, partOverride: "word/endnotes.xml");
            var settings = source.PageSettings with { HeightValue = "12000" };
            source = source with { PageHeightPoints = 600d, PageSettings = settings,
                FinalSectionBreak = source.FinalSectionBreak is { } final ? final with { PageSettings = settings } : null,
                BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
                    ? section with { PageSettings = section.PageSettings with { HeightValue = "12000" } } : element).ToArray() };
            if (emptyOnly)
                source = source with { RelatedStories = source.RelatedStories.Select(item => item.Id == "37"
                    ? item with { BodyElements = item.BodyElements.OfType<DocxTableElement>().Cast<DocxBodyElement>().ToArray() }
                    : item).ToArray() };
            var story = source.RelatedStories.Single(story => story.Id == "37");
            var math = DocxBlockTraversal.EnumerateBodyParagraphs(story).SelectMany(paragraph => paragraph.Equations)
                .ToDictionary(equation => equation, equation => new DocxPreparedMathParagraph(equation,
                    new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(layout.Pages.All(page => page.Height == 600d),
                "The raw section height and resolved dimensions must describe the intended test frame.");
            var canonical = layout.RelatedStories.Single(item => ReferenceEquals(item.Story, story));
            int expectedRows = multipleRows ? 2 : 1;
            TestAssert.Equal(expectedRows, canonical.TableRows.Count);
            TestAssert.True(canonical.TableRows.All(row => oversized ? row.Height > 456d : row.Height < 456d),
                "The source must exercise the declared fitting or oversized row condition.");
            if (emptyOnly)
            {
                CheckEmptyFrame(canonical, layout.Pages[0]);
                continue;
            }
            var owners = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(placed => ReferenceEquals(placed.StoryLayout.Story, story))
                .Select(placed => (Page: index, Placement: placed))).ToArray();
            var prefix = owners.Single(owner => owner.Placement.TextLines.Any(line => line.Text.Contains("Public destination Alpha.", StringComparison.Ordinal)));
            var rows = owners.Where(owner => owner.Placement.TableRows.Count != 0).ToArray();
            TestAssert.Equal(expectedRows, rows.Length);
            TestAssert.True(rows[0].Page > prefix.Page && rows.Select(owner => owner.Page).Distinct().Count() == expectedRows,
                $"A complete row defers from an occupied frame; the next row advances independently. Prefix={prefix.Page}/{prefix.Placement.TopY}/{prefix.Placement.Height}; rows={string.Join(";", rows.Select(owner => $"{owner.Page}/{owner.Placement.TopY}/{owner.Placement.Height}"))}; source={canonical.ContentHeight}/{canonical.TableRows[0].Height}; cellLines={canonical.TableRows[0].Cells.Sum(cell => cell.TextLines.Count)}/{DocxTableCellClipGeometry.EnumerateCellLines(canonical.TableRows[0], null, default).Count()}.");
            foreach (var owner in rows)
            {
                var placed = owner.Placement;
                var row = placed.TableRows.Single();
                if (oversized)
                    TestAssert.True(placed.Height > 0d && placed.Height <= layout.Pages[owner.Page].Height -
                        layout.Pages[owner.Page].MarginTop - layout.Pages[owner.Page].MarginBottom && row.Height > placed.Height,
                        "An oversized row retains its measured geometry and bounded page clip.");
                else
                    TestAssert.True(row.Y >= placed.TopY - placed.Height - .001d && row.Y + row.Height <= placed.TopY + .001d,
                        "The entire measured row lies inside its admitted story clip.");
                var visible = DocxTableCellClipGeometry.EnumerateCellLines(row, null, default).ToArray();
                TestAssert.Equal(row.Cells.Sum(cell => cell.TextLines.Count), visible.Length);
                TestAssert.Equal(2, row.Cells.SelectMany(cell => cell.TextLines).Sum(line => line.InlineMath.Count));
                TestAssert.Contains("TrailBoundarySentinel.", string.Concat(row.Cells.SelectMany(cell => cell.TextLines).Select(line => line.Text)));
            }
            foreach (var page in layout.Pages)
            foreach (var note in page.PlacedRelatedStories.Where(note => note.StoryLayout.Story.IsNormalStoryType))
            foreach (var line in note.TextLines)
                TestAssert.True(line.BaselineY >= note.TopY - note.Height - .001d && line.BaselineY <= note.TopY + .001d,
                    "A following note also needs a complete admitted line instead of a clipped forced take.");
        }
    }

    private static void CheckEmptyFrame(DocxRelatedStoryLayout story, DocxLayoutPage sourcePage)
    {
        var page = sourcePage with { Items = [], PlacedRelatedStories = [] };
        var pages = new List<DocxLayoutPage> { page };
        var method = typeof(DocxLayoutEngine).GetMethod("PlaceRelatedStorySlices", BindingFlags.Static | BindingFlags.NonPublic)!;
        object?[] arguments = Enumerable.Repeat<object?>(Type.Missing, method.GetParameters().Length).ToArray();
        arguments[0] = pages;
        arguments[1] = 0;
        arguments[2] = page;
        arguments[3] = new List<DocxPlacedRelatedStoryLayout>();
        arguments[4] = page.Height - page.MarginTop;
        arguments[5] = story;
        arguments[6] = 0;
        arguments[7] = true;
        method.Invoke(null, arguments);
        var placements = pages.SelectMany(frame => frame.PlacedRelatedStories).ToArray();
        double fullHeight = page.Height - page.MarginTop - page.MarginBottom;
        TestAssert.True(pages.Count > 0 && pages.Count <= Math.Ceiling(story.ContentHeight / fullHeight) + 1,
            "An empty frame makes bounded source progress without indefinite continuation pages.");
        TestAssert.True(placements.All(placed => placed.Height > 0d && placed.Height <= fullHeight),
            "Every height-only continuation retains its existing page clip.");
        TestAssert.True(Math.Abs(placements.Sum(placed => placed.Height) - story.ContentHeight) < .001d,
            "The existing height-only path consumes the entire measured source instead of dropping its tail.");
    }
}
