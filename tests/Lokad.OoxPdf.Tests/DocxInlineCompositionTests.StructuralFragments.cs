using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStructuralFragmentsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent fragment revisions use installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-fragment-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionStructuralFragmentsKeepActualOwnersAndDeletionFrames()
    {
        foreach (string role in new[] { "root", "delimiter", "operator", "fraction" })
        foreach (string kind in new[] { "ins", "del" })
        {
            string fixture = "CompositionProbes/math-fragment-" + role + "-" + kind;
            var source = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(equation.Revision is null && equation.HasRunRevisions &&
                DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) && owners.Count == 1 &&
                paragraph.Revisions.Count != 0 && paragraph.Revisions.All(owner => ReferenceEquals(owner, owners[0])) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "The actual fragment, rather than an invented whole-equation wrapper, owns its revision.");
            foreach (var invalid in new[] { equation with { SourceRunIndex = -1 }, equation with { HasRunRevisions = false },
                equation with { Content = new([]) } })
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { Equations = [invalid] }, default),
                    "Unknown positions, missing source and hints cannot replace actual fragment ownership.");
            if (kind == "del")
            {
                var final = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false);
                var retained = final.Paragraphs.First().Equations.Single();
                TestAssert.True(final.Paragraphs.First().Revisions.Count == 0 && !retained.HasRunRevisions &&
                    retained.Revision is null && DocxMathLayout.TryGetRuns(retained, default, out var runs) && runs.Count == 0,
                    "Final preserves the supported frame without deleted literal ink or author metadata.");
                TestAssert.True(role == "fraction" ? DocxMathLayout.IsEmptyFraction(retained) :
                    DocxMathLayout.GetStandaloneControl(retained, default) is not null,
                    "Deletion retains actual supported structural paint, including a glyphless fraction rule.");
                TestAssert.True(!DocxRenderer.BuildWordReviewAuthorPalette(final, [], default).AuthorSlots.ContainsKey("Whole Math Reviewer"),
                    "Excluded fragment owners cannot reserve Final author ink.");
            }
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
        }
        foreach (string role in new[] { "root", "delimiter", "operator" })
        {
            var empty = ReadMutation("CompositionProbes/math-fragment-" + role + "-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, m) => xml.Descendants(m + "oMath").First().Descendants(m + "r").Remove());
            var paragraph = empty.Paragraphs.First();
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph),
                "Structural-only fragments retain actual ownership and shared revision neutrality.");
        }
    }

    public static void DocxInlineCompositionStructuralFragmentsKeepUnknownDeeperAndNestedFallbacks()
    {
        foreach (string role in new[] { "root", "delimiter", "operator", "fraction" })
        {
            string fixture = "CompositionProbes/math-fragment-" + role + "-ins";
            foreach (int variant in new[] { 0, 1, 2 })
            {
                var source = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
                {
                    var wrapper = xml.Descendants(m + "oMath").First().Elements().Single();
                    var control = wrapper.Elements().Single();
                    if (variant == 0) control.ReplaceWith(new XElement(m + "future"));
                    else if (variant == 1) control.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "999"), new XElement(control)));
                    else wrapper.Add(new XElement(control));
                });
                TestAssert.True(!DocxMathLayout.TryGetRuns(source.Paragraphs.First().Equations.Single(), default, out _),
                    "Unknown, nested and multiple structural wrapper children keep explicit whole-equation fallback.");
            }
        }
        var deeper = ReadMutation("CompositionProbes/math-fragment-root-del", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
            xml.Descendants(m + "oMath").First().Descendants(m + "e").First().ReplaceNodes(new XElement(m + "f",
                new XElement(m + "num", new XElement(m + "r", new XElement(m + "t", "1"))),
                new XElement(m + "den", new XElement(m + "r", new XElement(m + "t", "2"))))));
        TestAssert.True(!DocxMathLayout.TryGetRuns(deeper.Paragraphs.First().Equations.Single(), default, out _),
            "Deleting a richer argument cannot silently erase its nested structural rules.");
    }
}
