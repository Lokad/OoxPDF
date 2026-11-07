using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxEndnoteMeasurementsSkipUnrelatedBodyWidths()
    {
        var source = ReadMutation("CompositionProbes/mixed-note-endnote-paragraph", OoxPdfDocxMarkupMode.Final,
            false, false, (xml, w, _) =>
            {
                var final = xml.Root!.Element(w + "body")!.Element(w + "sectPr")!;
                for (int index = 0; index < 16; index++)
                {
                    var section = new XElement(w + "sectPr",
                        new XElement(w + "type", new XAttribute(w + "val", "nextPage")),
                        new XElement(w + "pgSz", new XAttribute(w + "w", (12240 + index * 120).ToString(CultureInfo.InvariantCulture)),
                            new XAttribute(w + "h", "15840")),
                        new XElement(w + "pgMar", new XAttribute(w + "top", "1440"), new XAttribute(w + "bottom", "1440"),
                            new XAttribute(w + "left", "1440"), new XAttribute(w + "right", "1440")));
                    final.AddBeforeSelf(new XElement(w + "p", new XElement(w + "pPr", section),
                        new XElement(w + "r", new XElement(w + "t", "Public width section " + index))));
                }
            });
        source = source with { RelatedStories = source.RelatedStories.Select(story =>
        {
            if (!story.IsNormalStoryType) return story;
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Public measured note " + story.Id, 12d, 18d);
            return story with { BodyElements = [new DocxParagraphElement(paragraph)], SourceEquationCount = 0 };
        }).ToArray() };
        var measurements = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            { PlacementRelatedStoryMeasurements = measurements }.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        int widths = layout.Pages.Select(page => page.Width - page.MarginLeft - page.MarginRight).Distinct().Count();
        TestAssert.True(widths >= 16, "The source actually has many different body widths.");
        TestAssert.Equal(4, layout.Pages.SelectMany(page => page.PlacedRelatedStories)
            .Count(story => story.StoryLayout.Story.IsNormalStoryType));
        TestAssert.True(measurements.Count <= 2,
            "Endnotes need initial/final measurement, not every unrelated body width; retained " + measurements.Count + " widths for " + widths + " body widths.");
    }
}
