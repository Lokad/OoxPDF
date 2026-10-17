using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedPageNotesCountOnlyAutomaticOwnedMarkers()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (bool sameParagraph in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string input = "CompositionProbes/math-page-notes-mixed-custom-" + typography;
            var source = ReadMutation(input, mode, true, true, (xml, w, _) =>
            {
                if (!sameParagraph) return;
                var references = xml.Descendants(w + "footnoteReference").ToArray();
                var run = references[1].Parent!;
                run.Remove();
                references[0].Ancestors(w + "p").First().Add(run);
            });
            var custom = source.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.CustomMarkText is not null);
            var customParagraph = source.Paragraphs.Single(p => p.InlineReferences.Contains(custom));
            var customStory = source.RelatedStories.Single(story => story.IsNormalStoryType && story.Id == custom.Id);
            var math = source.Paragraphs.SelectMany(p => p.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var numbered, out var pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Complete custom glyph ownership cannot disable automatic counters.");
            TestAssert.Equal("*,09,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText ?? r.CustomMarkText)));
            TestAssert.Equal("*,105,106,107", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText ?? r.CustomMarkText)));
            var retainedCustom = numbered.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.CustomMarkText is not null);
            TestAssert.True(ReferenceEquals(custom, retainedCustom) && retainedCustom.DisplayText is null && retainedCustom.DisplayRun is null,
                "An authored mark does not acquire generated text or automatic marker identity.");
            var retainedStory = numbered.RelatedStories.Single(story => story.IsNormalStoryType && story.Id == custom.Id);
            TestAssert.True(ReferenceEquals(customStory.Paragraphs[0], retainedStory.Paragraphs[0]), "Manual target glyphs keep their original paragraph owner.");
            if (!sameParagraph) TestAssert.True(numbered.Paragraphs.Contains(customParagraph), "A paragraph with only a custom reference needs no replacement.");
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Numbering validates both manual and automatic source/target page maps.");
            TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final, pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Typography purposes retain their existing stricter marker scope.");
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(numbered, final, out var unchanged, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(numbered, unchanged),
                "Already correct labels do not trigger another coordinator retry.");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out _, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellation.Token));
        }
    }

    public static void DocxInlineCompositionMixedPageNotesRejectUnownedCustomTargets()
    {
        foreach (string variant in new[] { "mismatch", "hidden" })
        {
            var source = ReadMutation("CompositionProbes/math-page-notes-mixed-custom-calibri-12-zero", OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, _) =>
                {
                    var mark = xml.Descendants(w + "footnote").First(note => note.Attribute(w + "type") is null)
                        .Descendants(w + "t").First();
                    if (variant == "mismatch") mark.Value = "#";
                    else
                    {
                        var properties = mark.Parent!.Element(w + "rPr");
                        if (properties is null) { properties = new XElement(w + "rPr"); mark.Parent.AddFirst(properties); }
                        properties.Add(new XElement(w + "vanish"));
                    }
                }, partOverride: "word/footnotes.xml");
            var math = source.Paragraphs.SelectMany(p => p.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var unchanged, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(source, unchanged),
                "Mismatched or invisible custom targets retain the original document without partial renumbering.");
        }
    }
}
