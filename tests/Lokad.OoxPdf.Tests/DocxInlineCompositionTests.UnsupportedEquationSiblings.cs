using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxUnsupportedEquationRevisionsRetainActualOwners()
    {
        var source = ReadUnsupportedEquationSibling("CompositionProbes/math-run-revision-body-ins",
            OoxPdfDocxMarkupMode.AllMarkup, "ins");
        var paragraph = source.Paragraphs.First();
        var equation = paragraph.Equations.Last();
        var revision = paragraph.Revisions.Single(item => item.Id == "904");
        TestAssert.True(equation.HasRunRevisions && !DocxMathLayout.TryGetRuns(equation, default, out _),
            "The retained revised literal does not make its unsupported accent renderable.");
        TestAssert.True(DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners) &&
            owners.Count == 1 && ReferenceEquals(owners[0], revision) &&
            DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
            "The actual retained source owns its revision without suppressing independent siblings.");
        TestAssert.True(DocxMathLayout.GetControlSources(equation, default).Count == 0,
            "Unknown grammar does not acquire metric or paint controls.");
        foreach (var invalid in new[] { equation with { SourceRunIndex = -1 },
            equation with { Revision = revision with { } }, equation with { HasRunRevisions = false } })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(
                paragraph with { Equations = [paragraph.Equations[0], invalid] }, default),
                "Missing source positions, copied owners and false discovery hints retain fallback.");

        var literal = DocxMathLayout.EnumerateRetainedSources(equation.Content, default)
            .OfType<DocxMathRun>().Single(item => item.Text.Text == "z");
        var empty = equation with { Content = new([literal with { Text = literal.Text with { Text = string.Empty } },
            new DocxMathUnsupported("acc", "urn:unknown", null, [])]) };
        TestAssert.True(!DocxInlineContent.TryGetEquationRevisionOwners(empty, default, out _),
            "An empty literal alone still cannot establish an unsupported equation's revision owner.");
        using var canceled = new CancellationTokenSource();
        using var walk = DocxMathLayout.EnumerateRetainedSources(equation.Content, canceled.Token).GetEnumerator();
        TestAssert.True(walk.MoveNext(), "Traversal starts from the retained root.");
        canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => walk.MoveNext());
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
    }

    internal static DocxDocument ReadUnsupportedEquationSibling(string fixture, OoxPdfDocxMarkupMode mode,
        string? revisionKind) => ReadMutation(fixture, mode, false, false, (xml, w, m) =>
    {
        var first = xml.Descendants(m + "oMath").First();
        XElement sibling = first;
        while (sibling.Parent!.Name != w + "p") sibling = sibling.Parent;
        var literal = new XElement(m + "r", new XElement(m + "t", "z"));
        XElement owner = revisionKind is null ? literal : new XElement(w + revisionKind,
            new XAttribute(w + "id", "904"), new XAttribute(w + "author", "Public synthetic review"), literal);
        sibling.AddAfterSelf(new XElement(m + "oMath", owner,
            new XElement(m + "acc", new XElement(m + "e",
                new XElement(m + "r", new XElement(m + "t", "q"))))));
    });
}
