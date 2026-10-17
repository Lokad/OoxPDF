using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCompoundRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent compound revision frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "compound-revisions-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionCompoundRevisionsRetainActualRowsOwnersAndFallbacks()
    {
        foreach (string order in new[] { "script-first", "fraction-first" })
        foreach (string kind in new[] { "ordinary", "inserted" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/compound-revision-" + order + "-" + kind, mode, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            bool hasOwners = DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners);
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 6 &&
                runs.Sum(run => run.Run.Text.Text.EnumerateRunes().Count()) == 8 &&
                hasOwners &&
                owners.Count == (kind == "inserted" ? 1 : 0) && equation.Revision is null &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Ordinary compound row grammar retains actual literal owners through transparent revision syntax.");
            if (kind == "inserted")
                TestAssert.True(equation.HasRunRevisions && paragraph.Revisions.All(owner => ReferenceEquals(owner, owners[0])),
                    "One actual wrapper owns every retained source; value-equal or synthetic owners are unnecessary.");
        }
        const string fixture = "CompositionProbes/compound-revision-script-first-inserted";
        var deleted = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, m) => xml.Descendants(w + "ins").First().Name = w + "del");
        var retained = deleted.Paragraphs.First().Equations.Single();
        TestAssert.True(!DocxMathLayout.TryGetRuns(retained, default, out _) &&
            retained.Content.Children.Single() is DocxMathUnsupported { ElementName: "del", SourceXml: null, Children.Count: 3 },
            "Unsupported compound deletion retains traversed source and an explicit fallback instead of silently dropping structural paint.");
        foreach (bool nested in new[] { false, true })
        {
            var unknown = ReadMutation(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
            {
                var wrapper = xml.Descendants(w + "ins").First();
                if (nested)
                {
                    var child = wrapper.Elements().First();
                    child.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "999"), new XElement(child)));
                }
                else wrapper.Add(new XElement(m + "future"));
            });
            TestAssert.True(!DocxMathLayout.TryGetRuns(unknown.Paragraphs.First().Equations.Single(), default, out _),
                "Unknown mixed rows and nested revision syntax retain explicit whole-equation fallback.");
        }
    }
}
