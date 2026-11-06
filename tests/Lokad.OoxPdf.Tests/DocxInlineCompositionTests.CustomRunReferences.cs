using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSharedCustomRunReferencesPreserveSupportedMathAndMarkerOwners()
    {
        foreach (string fixture in new[] { "math-custom-note-footnote-1", "math-custom-note-endnote-1", "math-page-notes-mixed-custom-calibri-12-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool merge in new[] { false, true })
        {
            var source = ReadCustomRunReferences(fixture, mode, merge);
            var paragraph = source.Paragraphs.Single(p => p.InlineReferences.Any(r => r.CustomMarkFollowsValue == "1"));
            var custom = paragraph.InlineReferences[0];
            var automatic = paragraph.InlineReferences[1];
            TestAssert.Equal("*", custom.CustomMarkText!);
            TestAssert.Equal(merge, custom.SourceRunIndex == automatic.SourceRunIndex);
            TestAssert.True(DocxCustomNoteMarkerOwners.TryGetSourceRun(custom,
                DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, default), out var glyph) &&
                glyph.Text.StartsWith('*'), "The immediately following source character owns the custom reference.");
            TestAssert.True(DocxInlineContent.TryGetAutomaticNoteRun(automatic, out var marker) &&
                !ReferenceEquals(glyph, marker), "Custom and generated markers retain distinct owners even inside one XML run.");
            var equations = source.Paragraphs.SelectMany(p => p.Equations).ToArray();
            TestAssert.Equal(equations.Length, source.Paragraphs.Sum(p => DocxMathParagraphPolicy.GetPreparationEquations(p).Count));
            var prepared = equations.ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared);
            var placed = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>()
                .SelectMany(line => line.InlineMath.Select(math => math.Math.Equation));
            TestAssert.Equal(equations.Length, placed.Distinct(ReferenceEqualityComparer.Instance).Count());
        }
    }

    public static void DocxSharedCustomRunReferencesKeepAutomaticPageCountersAndStableOwners()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? controlLabels = null;
            foreach (bool merge in new[] { false, true })
            {
                var source = ReadCustomRunReferences("math-page-notes-mixed-custom-calibri-12-zero", mode, merge);
                var prepared = source.Paragraphs.SelectMany(p => p.Equations).ToDictionary(e => e,
                    e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
                var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared);
                TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var numbered, out var pages,
                    allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                    "Complete custom and generated source owners can share page-restarted numbering.");
                string labels = string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.CustomMarkText ?? r.DisplayText));
                TestAssert.True(labels.StartsWith("*,09", StringComparison.Ordinal), "A custom glyph does not consume an automatic number.");
                if (controlLabels is not null) TestAssert.Equal(controlLabels, labels);
                controlLabels = labels;
                TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered,
                    engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared), pages,
                    allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                    "The retry preserves every complete custom/generated marker's source page.");
            }
        }
    }

    public static void DocxSharedCustomRunReferencesKeepImmediateCharacterFallbacks()
    {
        foreach (string text in new[] { "", " ", "\ud83d\ude00" })
        {
            var source = ReadCustomRunReferences("math-custom-note-footnote-1", OoxPdfDocxMarkupMode.Final, merge: true,
                (xml, w) => xml.Descendants(w + "footnoteReference").First().ElementsAfterSelf().First(e => e.Name == w + "t").Value = text);
            TestAssert.True(source.Paragraphs.SelectMany(p => p.InlineReferences).First().CustomMarkText is null,
                "Empty, whitespace and surrogate custom marks retain their unsupported-source fallback.");
        }
        var acrossRuns = ReadCustomRunReferences("math-custom-note-footnote-1", OoxPdfDocxMarkupMode.Final, merge: true,
            (xml, w) => {
                var reference = xml.Descendants(w + "footnoteReference").First();
                reference.ElementsAfterSelf().First(e => e.Name == w + "t").Remove();
                reference.Parent!.AddAfterSelf(new XElement(w + "r", new XElement(w + "t", "*")));
            });
        TestAssert.True(acrossRuns.Paragraphs.SelectMany(p => p.InlineReferences).First().CustomMarkText is null,
            "A custom marker cannot borrow a character from another run or across another reference.");
    }

    private static DocxDocument ReadCustomRunReferences(string fixture, OoxPdfDocxMarkupMode mode, bool merge,
        Action<XDocument, XNamespace>? modify = null) => ReadMutation("CompositionProbes/" + fixture, mode, false, false, (xml, w, _) =>
        {
            string kind = fixture.Contains("endnote", StringComparison.Ordinal) ? "endnoteReference" : "footnoteReference";
            var pair = xml.Descendants(w + kind).Take(2).ToArray();
            var first = pair[0].Parent!; var second = pair[1].Parent!;
            TestAssert.True(XNode.DeepEquals(first.Element(w + "rPr"), second.Element(w + "rPr")), "Merged runs need identical formatting.");
            second.Remove(); first.AddAfterSelf(second);
            if (merge)
            {
                first.Add(second.Elements().Where(e => e.Name != w + "rPr").Select(e => new XElement(e)));
                second.Remove();
            }
            modify?.Invoke(xml, w);
        });
}
