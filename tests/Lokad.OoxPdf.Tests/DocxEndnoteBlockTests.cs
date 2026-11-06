using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxEndnoteBlockTests
{
    public static void DocxCompleteEndnoteRowDefersWithoutClippingItsContent() => Check(false);

    public static void DocxCompleteEndnoteRowsAdvanceIndependentlyOfDirectTextLines() => Check(true);

    private static void Check(bool multipleRows)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table",
                mode, false, false, (xml, w, _) =>
                {
                    var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                    var row = note.Element(w + "tbl")!.Element(w + "tr")!;
                    row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                        new XAttribute(w + "val", "8400"), new XAttribute(w + "hRule", "atLeast"))));
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
            TestAssert.True(canonical.TableRows.All(row => row.Height < 456d),
                "Each source row must fit the full page; multiple rows need not fit together.");
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
}
