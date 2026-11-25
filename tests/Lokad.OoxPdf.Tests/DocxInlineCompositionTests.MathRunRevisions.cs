using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMathRunRevisionsRequireActualOwners()
    {
        var source = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = source.Paragraphs.First();
        var equation = paragraph.Equations.Single();
        var revision = paragraph.Revisions.Single();
        TestAssert.True(equation.Revision is null && equation.HasRunRevisions && paragraph.InlineContent is not null &&
            DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
            "The actual numerator run owns its revision without revising the denominator or whole equation.");
        foreach (var invalid in new[] { paragraph with { Revisions = [revision with { }] },
            paragraph with { Equations = [equation with { HasRunRevisions = false }] },
            paragraph with { Equations = [equation with { SourceRunIndex = -1 }] } })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(invalid, default),
                "Copied aggregate metadata and a discovery hint cannot invent source ownership.");
        foreach (string kind in new[] { "moveTo", "rPrChange" })
        {
            var unknown = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, m) => xml.Descendants(w + "ins").Single().Name = w + kind);
            TestAssert.True(!DocxMathLayout.TryGetRuns(unknown.Paragraphs.First().Equations.Single(), default, out _),
                "Unknown structural revisions retain the complete equation fallback.");
        }
        var empty = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
            (xml, w, m) => xml.Descendants(w + "ins").Single().Descendants(m + "t").Single().Value = string.Empty);
        TestAssert.True(!DocxMathLayout.TryGetRuns(empty.Paragraphs.First().Equations.Single(), default, out _),
            "Unqualified empty inner revision wrappers retain their existing fallback.");
        var mixed = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
            (xml, w, m) =>
            {
                XElement revision = xml.Descendants(w + "ins").Single();
                var mixedRun = new XElement(m + "r", new XElement(m + "t", "u"),
                    new XElement(w + "ins", revision.Attributes(), new XElement(m + "t", "x+2")));
                revision.ReplaceWith(mixedRun);
            });
        TestAssert.True(!DocxMathLayout.TryGetRuns(mixed.Paragraphs.First().Equations.Single(), default, out _),
            "An unchanged literal and a revision cannot be merged under the same owner.");
        foreach (bool wordProperties in new[] { false, true })
        {
            var conflicting = ReadMutation("CompositionProbes/math-run-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, m) =>
                {
                    XElement revision = xml.Descendants(w + "ins").Single();
                    XName property = (wordProperties ? w : m) + "rPr";
                    XElement outer = wordProperties
                        ? new XElement(property, new XElement(w + "color", new XAttribute(w + "val", "1840A0")))
                        : new XElement(property, new XElement(m + "sty", new XAttribute(m + "val", "p")));
                    XElement inner = wordProperties
                        ? new XElement(property, new XElement(w + "color", new XAttribute(w + "val", "A04018")))
                        : new XElement(property, new XElement(m + "sty", new XAttribute(m + "val", "i")));
                    revision.ReplaceWith(new XElement(m + "r", outer,
                        new XElement(w + "ins", revision.Attributes(), inner, new XElement(m + "t", "x+2"))));
                });
            TestAssert.True(!DocxMathLayout.TryGetRuns(conflicting.Paragraphs.First().Equations.Single(), default, out _),
                "Conflicting declarations of the same property kind retain the diagnosed fallback.");
        }
        var deleted = ReadMutation("CompositionProbes/math-run-revision-body-del", OoxPdfDocxMarkupMode.Final, false, false);
        var visible = deleted.Paragraphs.First();
        TestAssert.True(visible.Revisions.Count == 0 && !visible.Equations.Single().HasRunRevisions &&
            DocxMathLayout.TryGetRuns(visible.Equations.Single(), default, out var runs) &&
            runs.Count == 1 && runs[0].Run.Text.Text == "y",
            "Final retains the denominator and empty numerator structure without deleted glyph ownership.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
    }

    public static void DocxInlineCompositionMathRunRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent inner-math revision frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-run-revisions-office.json", splitMathRuns: true);
    }
}
