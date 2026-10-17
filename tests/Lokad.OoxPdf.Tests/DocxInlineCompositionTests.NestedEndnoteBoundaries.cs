using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFittingNestedEndnotesMatchIndependentWordBoundaries()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent nested endnote boundaries use installed Windows fonts.");
        CheckFittingEndnoteBoundaries("endnote-fitting-nested-office.json");
    }

    public static void DocxInlineCompositionFittingNestedEndnotesRetainAncestorChildAndPartialClipFallbacks()
    {
        foreach (string variant in new[] { "ancestor-clip", "child-clip", "partial-clip", "merge", "fit", "vertical", "revision" })
        {
            var source = ReadMutation("CompositionProbes/mixed-nested-math-endnote", OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, _) =>
                {
                    var story = xml.Root!.Elements(w + "endnote").Single(item => (string?)item.Attribute(w + "id") == "37");
                    var outer = story.Element(w + "tbl")!;
                    var outerCell = outer.Element(w + "tr")!.Element(w + "tc")!;
                    var inner = outerCell.Element(w + "tbl")!;
                    var cell = inner.Element(w + "tr")!.Element(w + "tc")!;
                    if (variant is "ancestor-clip" or "child-clip" or "partial-clip")
                    {
                        var table = variant == "child-clip" ? inner : outer;
                        table.Element(w + "tr")!.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                            new XAttribute(w + "val", variant == "partial-clip" ? "800" : "20"), new XAttribute(w + "hRule", "exact"))));
                        if (variant == "partial-clip")
                        {
                            outerCell.Elements(w + "p").Remove();
                            cell.Element(w + "p")!.Add(new XElement(w + "r", new XElement(w + "t",
                                string.Concat(Enumerable.Repeat(" More nested note words.", 80)))));
                        }
                    }
                    else if (variant == "revision")
                        inner.Element(w + "tr")!.AddFirst(new XElement(w + "trPr", new XElement(w + "ins",
                            new XAttribute(w + "id", "71"), new XAttribute(w + "author", "Public author"))));
                    else cell.Element(w + "tcPr")!.Add(variant switch
                    {
                        "merge" => new XElement(w + "vMerge", new XAttribute(w + "val", "restart")),
                        "fit" => new XElement(w + "tcFitText"),
                        _ => new XElement(w + "textDirection", new XAttribute(w + "val", "tbRl"))
                    });
                }, partOverride: "word/endnotes.xml");
            var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
                .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            if (variant == "partial-clip")
            {
                var story = layout.RelatedStories.Single(item => item.Story.Id == "37");
                int measured = NestedMathRows(story.TableRows).SelectMany(row => row.Cells).Sum(cell => cell.TextLines.Count);
                int visible = story.TableRows.Sum(row => DocxTableCellClipGeometry.EnumerateCellLines(row, null, default).Count());
                TestAssert.True(visible > 0 && visible < measured,
                    "The control retains visible lines of one paragraph while later lines cross the ancestor clip.");
            }
            TestAssert.True(!layout.Pages.SelectMany(page => page.PlacedRelatedStories).Any(story => story.UsesCompleteDocumentEndnoteFlow),
                "Incomplete or unqualified nested owners retain the complete boundary fallback: " + variant);
        }
    }
}
