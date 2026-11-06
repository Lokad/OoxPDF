using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxEndnoteParagraphBoundaryTests
{
    public static void DocxLaterEndnotePitchDoesNotDependOnEarlierTable() => Check(0, 0);

    public static void DocxLaterEndnotesCollapseSourceBeforeAndAfterSpacing()
    {
        Check(200, 120);
        Check(120, 200);
    }

    private static void Check(int before, int after)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            double[] Measure(bool table)
            {
                var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table",
                    mode, false, false, (xml, w, _) =>
                    {
                        var first = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                        if (table)
                        {
                            first.Element(w + "tbl")!.Element(w + "tr")!.AddFirst(new XElement(w + "trPr",
                                new XElement(w + "trHeight", new XAttribute(w + "val", "8400"), new XAttribute(w + "hRule", "atLeast"))));
                            first.Elements(w + "p").Last().Add(new XElement(w + "r", new XElement(w + "t",
                                string.Concat(Enumerable.Repeat(" Wrapped trailing paragraph.", 180)))));
                        }
                        else first.Elements().Skip(1).Remove();
                        foreach (var note in xml.Root.Elements(w + "endnote").Where(note =>
                            (string?)note.Attribute(w + "id") is "4" or "73" or "8"))
                        {
                            var spacing = note.Element(w + "p")!.Element(w + "pPr")!.Element(w + "spacing")!;
                            spacing.SetAttributeValue(w + "before", before);
                            spacing.SetAttributeValue(w + "after", after);
                        }
                    }, partOverride: "word/endnotes.xml");
                var settings = source.PageSettings with { HeightValue = "12000" };
                source = source with { PageHeightPoints = 600d, PageSettings = settings,
                    FinalSectionBreak = source.FinalSectionBreak is { } final ? final with { PageSettings = settings } : null,
                    BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
                        ? section with { PageSettings = section.PageSettings with { HeightValue = "12000" } } : element).ToArray() };
                var math = source.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs)
                    .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                        equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
                var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
                TestAssert.True(layout.Pages.All(page => page.Height == 600d), "Use the same authored and resolved page frame.");
                var tail = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(note => note.StoryLayout.Story.Id is "4" or "73" or "8")
                    .Select(note => (Page: index, Line: note.TextLines.Single()))).ToArray();
                TestAssert.Equal(3, tail.Length);
                double[] pitches = tail.Zip(tail.Skip(1), (left, right) => (Left: left, Right: right))
                    .Where(pair => pair.Left.Page == pair.Right.Page)
                    .Select(pair => pair.Left.Line.BaselineY - pair.Right.Line.BaselineY).ToArray();
                TestAssert.True(pitches.Length > 0, "Compare an actual shared paragraph boundary on one page.");
                return pitches;
            }
            double[] ordinary = Measure(false), composed = Measure(true);
            TestAssert.True(ordinary.All(pitch => Math.Abs(pitch - ordinary[0]) < .001d),
                "Uniform source spacing has the same pitch at every ordinary boundary.");
            foreach (double pitch in composed)
                TestAssert.True(Math.Abs(ordinary[0] - pitch) < .001d,
                    $"An earlier supported table must not change later paragraph-boundary spacing ({mode}, {before}/{after}): {ordinary[0]} vs {pitch}.");
        }
    }
}
