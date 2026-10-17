using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionWholeStructuralRevisionsKeepContentAndSourceInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent whole-equation references use installed Windows fonts.");
        // Raw geometry and raster failures remain separate qualification evidence.
        CheckEquationRevisions(null, false, "whole-math-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionWholeStructuralRevisionsKeepActualSourceOwners()
    {
        foreach (string role in new[] { "root", "delimiter", "operator", "root-degree", "empty-root", "empty-delimiter", "empty-operator" })
        {
            var document = ReadMutation("CompositionProbes/whole-math-" + role + "-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var paragraph = document.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(equation.Display && equation.Revision?.Kind == DocxRevisionKind.Insertion,
                "The whole equation retains its actual display-wrapper revision.");
            TestAssert.True(DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) &&
                owners.Count == 1 && ReferenceEquals(owners[0], equation.Revision) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Literal or structural paint owns the actual source revision.");
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs) &&
                (runs.Count == 0) == role.StartsWith("empty-", StringComparison.Ordinal),
                "Structural-only equations do not invent text runs.");
            foreach (var invalid in new[] { equation with { Revision = equation.Revision! with { } },
                equation with { SourceRunIndex = -1 }, equation with { Content = new([]) } })
                TestAssert.True(!DocxInlineContent.TryGetEquationRevisionOwners(invalid, default, out _),
                    "Hints and value-equal metadata cannot replace actual source ownership.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
        }
    }

    public static void DocxInlineCompositionWholeStructuralRevisionsKeepBoundedDeletionFallbacks()
    {
        foreach (string role in new[] { "root", "delimiter", "operator", "root-degree" })
        {
            string fixture = "CompositionProbes/whole-math-" + role + "-del";
            var document = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false);
            var paragraph = document.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(equation.Revision is null && paragraph.Revisions.Count == 0 &&
                DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 0,
                "Final preserves structural paint without deleted literal text or author metadata.");
            TestAssert.True(!DocxRenderer.BuildWordReviewAuthorPalette(document, [], default).AuthorSlots.ContainsKey("Whole Math Reviewer"),
                "Excluded owners cannot reserve author ink in Final.");
            var unknown = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
                xml.Descendants(m + "oMath").First().Descendants(m + "e").First().ReplaceNodes(new XElement(m + "future")));
            TestAssert.Equal(0, unknown.Paragraphs.First().Equations.Count);
            var nested = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
            {
                var source = xml.Descendants(m + "oMath").First();
                source.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "999"),
                    new XAttribute(w + "author", "Nested Reviewer"), new XElement(source)));
            });
            TestAssert.Equal(0, nested.Paragraphs.First().Equations.Count);
        }
    }
}
