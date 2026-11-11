using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInsertedInlineBoxesRequireTheirActualRevisionOwners()
    {
        foreach (bool shared in new[] { false, true })
        {
            var document = ReadInsertedBoxSource("inline-box-body-equation-inline-box-valid", OoxPdfDocxMarkupMode.Final,
                sharePrecedingText: shared);
            var paragraph = document.Paragraphs.Single(item => item.InlineTextBoxes.Count != 0);
            var box = paragraph.InlineTextBoxes.Single(); var revision = box.Revisions.Single();
            TestAssert.Equal(DocxRevisionKind.Insertion, revision.Kind);
            TestAssert.True(paragraph.Revisions.Any(item => ReferenceEquals(item, revision)), "The parsed drawing owns an actual declared wrapper.");
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph), "An actual drawing insertion cannot erase its supported siblings.");
            TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with { Revisions = [] }), "Missing aggregate metadata cannot invent ownership.");
            foreach (var invalid in new[]
            {
                box with { SourceRunIndex = -1 },
                box with { Revisions = [revision with { }] },
                box with { Revisions = [revision with { Kind = DocxRevisionKind.Deletion }] }
            })
            {
                var changed = paragraph with { InlineTextBoxes = [invalid], InlineContent = paragraph.InlineContent!.Select(item =>
                    item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(invalid) : item).ToArray() };
                TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(changed), "Foreign indices, copied metadata and other box revisions retain their fallback.");
            }
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedSourceOwner(paragraph, token: cancellation.Token));
        }
    }

    internal static DocxDocument ReadInsertedBoxSource(string fixture, OoxPdfDocxMarkupMode mode, string? part = null,
        bool sharePrecedingText = false) => ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, _) =>
        {
            var content = xml.Descendants(w + "txbxContent").Single();
            var run = content.Ancestors(w + "drawing").First().Parent!;
            TestAssert.True(run.Name == w + "r", "The insertion wraps the existing drawing's source run.");
            XElement? preceding = sharePrecedingText ? run.PreviousNode as XElement : null;
            if (sharePrecedingText) TestAssert.True(preceding?.Name == w + "r", "The shared wrapper keeps adjacent source text in order.");
            var wrapper = new XElement(w + "ins", new XAttribute(w + "id", "920"),
                new XAttribute(w + "author", "Public Reviewer"), new XAttribute(w + "date", "2026-11-11T10:00:00Z"));
            run.AddBeforeSelf(wrapper);
            if (preceding is not null) { preceding.Remove(); wrapper.Add(preceding); }
            run.Remove(); wrapper.Add(run);
        }, partOverride: part);
}
