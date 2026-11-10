using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEndnoteCounterPrefixTests
{
    public static void DocxEndnoteCounterPrefixSectionStartAndFollowingHeadIgnoreLaterOverflow()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (double after in new[] { 0d, 8d })
        {
            DocxDocument Source(int words)
            {
                var source = Read(mode, "sectEnd", 1, null, "1", words);
                return source with { BodyElements = source.BodyElements.Select(element =>
                    element is DocxParagraphElement { Paragraph: var p } paragraph && p.InlineReferences.Any(r => r.Id == "4")
                        ? paragraph with { Paragraph = p with { SpacingAfterPoints = after,
                            Spacing = p.Spacing with { AfterValue = ((int)(after * 20d)).ToString(System.Globalization.CultureInfo.InvariantCulture) } } }
                        : element).ToArray() };
            }
            var before = SectionLayout(Source(0), mode);
            var afterLayout = SectionLayout(Source(400), mode);
            foreach (string id in new[] { "37", "4" })
            {
                var a = FirstSectionNote(before, id);
                var b = FirstSectionNote(afterLayout, id);
                TestAssert.Equal(a.Page, b.Page);
                TestAssert.Equal(a.Story.TopY, b.Story.TopY);
                TestAssert.Equal(a.Story.TextLines[0].BaselineY, b.Story.TextLines[0].BaselineY);
                if (id == "37")
                    TestAssert.Equal(a.Story.TableRows.Single().Cells.Single().InlineImages.Single().Y,
                        b.Story.TableRows.Single().Cells.Single().InlineImages.Single().Y);
            }
            string text = string.Join(" ", afterLayout.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(story => story.StoryLayout.Story.Id == "4").SelectMany(story => story.TextLines).Select(line => line.Text));
            foreach (int index in Enumerable.Range(0, 400))
                TestAssert.True(text.Contains($"LaterNote{index:000}", StringComparison.Ordinal), "Retain every continued source token.");
        }
    }

    public static void DocxEndnoteCounterPrefixSectionStartKeepsSourceAndScaleFallbacks()
    {
        foreach (string variant in new[] { "keep", "source-clip", "scaled" })
        {
            DocxDocument Source(int words) => Read(OoxPdfDocxMarkupMode.Final, "sectEnd", 1, null, "1", words,
                (xml, w) =>
                {
                    var note = xml.Root!.Elements(w + "endnote").Single(n => (string?)n.Attribute(w + "id") == "37");
                    if (variant == "keep") note.Element(w + "p")!.Element(w + "pPr")!.Add(new XElement(w + "keepLines"));
                    if (variant == "source-clip") note.Element(w + "tbl")!.Element(w + "tr")!.AddFirst(new XElement(w + "trPr",
                        new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact"))));
                });
            var before = FirstSectionNote(SectionLayout(Source(0), OoxPdfDocxMarkupMode.Final, variant == "scaled"), "37");
            var after = FirstSectionNote(SectionLayout(Source(400), OoxPdfDocxMarkupMode.Final, variant == "scaled"), "37");
            TestAssert.Equal(before.Page, after.Page);
            TestAssert.Equal(before.Story.TopY, after.Story.TopY);
            TestAssert.Equal(before.Story.TextLines[0].BaselineY, after.Story.TextLines[0].BaselineY);
        }
    }

    private static DocxLayout SectionLayout(DocxDocument source, OoxPdfDocxMarkupMode mode, bool scaled = false)
    {
        var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
            .SelectMany(p => p.Equations).ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
        return new DocxLayoutEngine(scaled || mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
                wordCompatiblePrintScale: scaled ? .8d : 1d)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
    }

    private static (int Page, DocxPlacedRelatedStoryLayout Story) FirstSectionNote(DocxLayout layout, string id) => layout.Pages
        .SelectMany((page, index) => page.PlacedRelatedStories.Where(story => story.StoryLayout.Story.Id == id)
            .Select(story => (index, story))).First();
}
