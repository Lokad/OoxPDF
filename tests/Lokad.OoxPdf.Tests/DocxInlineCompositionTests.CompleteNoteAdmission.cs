using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCompleteNoteAdmissionUsesSourceLinesAndSharedSeparator()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (string fixture in new[] { "spanning-page-notes-", "math-page-notes-compound-", "math-page-notes-mixed-custom-" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/" + fixture + typography, mode, false, false);
            var equations = source.Paragraphs.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.True(layout.UsesSourceOwnedFootnoteReserve,
                "Complete source-owned notes share admission across ordinary and mixed paragraphs without font/size rules.");
            var destinations = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(note => note.StoryLayout.Story.IsNormalStoryType).Select(note => (Page: index, Note: note))).ToArray();
            TestAssert.Equal(4, destinations.Length);
            TestAssert.Equal(4, destinations.Select(item => item.Note.StoryLayout.Story.Id).Distinct().Count());
            foreach (var (page, note) in destinations)
            {
                var reference = source.Paragraphs.SelectMany(paragraph => paragraph.InlineReferences).Single(reference => reference.Id == note.StoryLayout.Story.Id);
                var sourceParagraph = source.Paragraphs.Single(paragraph => paragraph.InlineReferences.Contains(reference));
                TestAssert.True(layout.Pages[page].Items.OfType<DocxTextLineLayout>().Any(line =>
                    ReferenceEquals(line.SourceParagraph, sourceParagraph) && line.Segments.Any(segment =>
                        segment.SourceTextRunIndex == reference.SourceRunIndex && segment.SourceTextOffsetInRun <= reference.TextOffsetInRun &&
                        reference.TextOffsetInRun < (long)segment.SourceTextOffsetInRun + segment.Text.Length)),
                    "Each admitted note's page contains its actual source glyph, including custom markers.");
            }
            foreach (var page in layout.Pages)
                TestAssert.True(page.PlacedRelatedStories.Count(note => note.StoryLayout.Story.Type == DocxRelatedStoryType.Separator) <= 1,
                    "Complete notes paint one shared separator on their physical page.");
        }
    }

    public static void DocxInlineCompositionCompleteNoteAdmissionRetainsComplexFrameFallbacks()
    {
        foreach (string variant in new[] { "keep", "columns", "split-note", "hidden-marker", "ambiguous-marker" })
        {
            var source = ReadMutation("CompositionProbes/spanning-page-notes-calibri-12-zero", OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, _) =>
                {
                    if (variant == "split-note")
                        xml.Descendants(w + "footnote").First(note => note.Attribute(w + "type") is null)
                            .Descendants(w + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped note text.", 50));
                    else if (variant == "columns") xml.Descendants(w + "sectPr").Single()
                        .Add(new XElement(w + "cols", new XAttribute(w + "num", "2")));
                    else if (variant == "keep") xml.Descendants(w + "pPr").First().Add(new XElement(w + "keepLines"));
                    else if (variant == "hidden-marker")
                    {
                        var run = xml.Descendants(w + "footnoteReference").First().Parent!;
                        var properties = run.Element(w + "rPr");
                        if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
                        properties.Add(new XElement(w + "vanish"));
                    }
                }, partOverride: variant == "split-note" ? "word/footnotes.xml" : null);
            if (variant == "ambiguous-marker")
            {
                var paragraph = source.Paragraphs.Single();
                var reference = paragraph.InlineReferences[0];
                var changed = paragraph with { InlineReferences = [.. paragraph.InlineReferences, reference] };
                source = source with { BodyElements = [new DocxParagraphElement(changed)], FallbackParagraphs = [changed] };
            }
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(!layout.UsesSourceOwnedFootnoteReserve,
                "Complex frame or unowned marker retains the existing reserve path: " + variant);
        }
    }
}
