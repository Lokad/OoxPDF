using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedContainerAdmissionKeepsActualNoteOwners()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (bool twoCellParagraphs in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/spanning-note-math-cell-tail-" + typography, mode, false, false,
                (xml, w, _) =>
                {
                    if (!twoCellParagraphs) return;
                    var body = xml.Root!.Element(w + "body")!;
                    var marker = body.Element(w + "p")!.Descendants(w + "footnoteReference").Last().Parent!;
                    marker.Remove();
                    var cell = body.Element(w + "tbl")!.Descendants(w + "tc").Single();
                    var paragraph = new XElement(cell.Element(w + "p")!);
                    paragraph.Elements(w + "r").First(run => run.Element(w + "footnoteReference") is not null).ReplaceWith(marker);
                    cell.Element(w + "p")!.AddBeforeSelf(paragraph);
                });
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
            var equations = paragraphs.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(layout.UsesSourceOwnedFootnoteReserve, "A cell reference must not disable measured body-note admission.");
            var notes = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(note => note.StoryLayout.Story.IsNormalStoryType).Select(note => (Page: index, Note: note))).ToArray();
            TestAssert.Equal(4, notes.Length);
            TestAssert.Equal(4, notes.Select(item => item.Note.StoryLayout.Story.Id).Distinct().Count());
            foreach (var (page, note) in notes)
            {
                var paragraph = paragraphs.Single(item => item.InlineReferences.Any(reference => reference.Id == note.StoryLayout.Story.Id));
                var reference = paragraph.InlineReferences.Single(item => item.Id == note.StoryLayout.Story.Id);
                TestAssert.True(DocxTableCellClipGeometry.EnumerateBodyLines(layout.Pages[page], default).Any(item =>
                    ReferenceEquals(item.Line.SourceParagraph, paragraph) && item.Line.Segments.Any(segment =>
                        segment.SourceTextRunIndex == reference.SourceRunIndex && segment.SourceTextOffsetInRun <= reference.TextOffsetInRun &&
                        reference.TextOffsetInRun < (long)segment.SourceTextOffsetInRun + segment.Text.Length)),
                    "Every admitted body/cell story shares the page of its exact paragraph/run owner.");
            }
            TestAssert.True(layout.Pages.All(page => page.PlacedRelatedStories.Count(note =>
                note.StoryLayout.Story.Type == DocxRelatedStoryType.Separator) <= 1), "A physical page has one separator.");
        }
    }

    public static void DocxInlineCompositionMixedContainerAdmissionRetainsFallbacksAndCancellation()
    {
        foreach (string variant in new[] { "keep", "nested", "split-note", "hidden-marker" })
        {
            var source = ReadMutation("CompositionProbes/spanning-note-math-cell-tail-calibri-12-zero",
                OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
                {
                    if (variant == "split-note")
                        xml.Descendants(w + "footnote").First(note => note.Attribute(w + "type") is null)
                            .Descendants(w + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped note text.", 50));
                    else if (variant == "keep") xml.Descendants(w + "tc").Single().Element(w + "p")!.Element(w + "pPr")!
                        .Add(new XElement(w + "keepLines"));
                    else if (variant == "nested")
                    {
                        var cell = xml.Descendants(w + "tc").Single();
                        var paragraph = cell.Element(w + "p")!;
                        paragraph.ReplaceWith(new XElement(w + "tbl", new XElement(w + "tr", new XElement(w + "tc", paragraph))));
                    }
                    else
                    {
                        var run = xml.Descendants(w + "tc").Single().Descendants(w + "footnoteReference").Single().Parent!;
                        var properties = run.Element(w + "rPr");
                        if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
                        properties.Add(new XElement(w + "vanish"));
                    }
                }, partOverride: variant == "split-note" ? "word/footnotes.xml" : null);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(!layout.UsesSourceOwnedFootnoteReserve, "Existing complex/container fallback remains: " + variant);
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var document = ReadMutation("CompositionProbes/spanning-note-math-cell-tail-calibri-12-zero", OoxPdfDocxMarkupMode.Final, false, false);
        TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), cancellation.Token));
    }

    public static void DocxInlineCompositionMixedContainerAdmissionMatchesIndependentPagesAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-container geometry uses installed Windows fonts.");
        CheckSpanningPageNotes(null, requireIndependentGeometry: true, referenceFile: "spanning-note-math-cell-tail-office.json");
    }
}
