using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxEmptyStoredResultsKeepNeighborSourceAcrossPlacements()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadEmptyStoredResult(container, complex, mode);
            var paragraph = CachedMathParagraph(document);
            var field = paragraph.FieldReferences.Single(f => f.Instruction!.Trim() == "REF EmptyStoredResult");
            TestAssert.True(!field.HasCachedResult && !field.RendersCachedResult && !field.UsesPlaceholder,
                "An empty stored result does not invent a cache or a computed value.");
            TestAssert.Equal(0, field.TextRunCount);
            TestAssert.Equal(0, field.TextLength);
            TestAssert.Equal(2, paragraph.Equations.Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }

    public static void DocxEmptyStoredResultsPreservePaintAndNavigation()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(ReadEmptyStoredResult(container, complex, mode, baseline: true), null, default)
                .ToBlockingEnumerable().ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(ReadEmptyStoredResult(container, complex, mode, diagnosticSink: diagnostics.Add), diagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int index = 0; index < expected.Length; index++)
            {
                TestAssert.Equal(expected[index].Content, actual[index].Content);
                TestAssert.Equal(expected[index].Fonts.Count, actual[index].Fonts.Count);
                TestAssert.Equal(expected[index].Images.Count, actual[index].Images.Count);
                TestAssert.True(expected[index].Annotations.SequenceEqual(actual[index].Annotations),
                    "An empty source result leaves neighboring painted and navigation geometry unchanged.");
            }
            TestAssert.True(diagnostics.All(d => d.Id is not "DOCX_UNSUPPORTED_EQUATION" and not "DOCX_UNSUPPORTED_COMPLEX_FIELD"),
                "A complete empty result neither omits neighboring math nor reports ignored field content.");
        }
    }

    public static void DocxEmptyStoredResultsKeepNestedComputedAndIncompleteOwners()
    {
        foreach (bool complex in new[] { false, true })
        foreach (string instruction in new[] { "REF EmptyStoredResult", "PAGE", "NUMPAGES" })
        {
            var nested = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, m) => xml.Descendants(m + "oMath").First().AddBeforeSelf(
                    CachedMathField(w, CachedMathField(w, Array.Empty<XElement>(), complex, instruction), !complex)));
            var paragraph = CachedMathParagraph(nested);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.True(paragraph.FieldReferences.Where(f => f.TextLength == 0)
                .All(f => !f.HasCachedResult && !f.RendersCachedResult && !f.UsesPlaceholder),
                "Independent locks retain genuinely empty inner and outer results, including locked counters.");
        }
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (bool complex in new[] { false, true })
        {
            var computed = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, m) => xml.Descendants(m + "oMath").First().AddBeforeSelf(
                    CachedMathField(w, Array.Empty<XElement>(), complex, instruction, locked: false)));
            var field = CachedMathParagraph(computed).FieldReferences.Single(f => f.UsesPlaceholder);
            TestAssert.True(!field.HasCachedResult && !field.RendersCachedResult && field.Kind.IsDynamic(),
                "Unlocked empty counters still compute through their existing placeholder owner.");
            TestAssert.True(!DocxInlineContent.HasMaterializedFieldResults([field], []),
                "An unplaced computed value does not masquerade as an empty stored result.");
            TestAssert.True(DocxInlineContent.HasSourceFieldResults([field], []),
                "The page-input path still accepts its existing computed dependency.");
        }
        var diagnostics = new List<OoxPdfDiagnostic>();
        var incomplete = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, m) =>
            {
                var field = CachedMathField(w, Array.Empty<XElement>(), complex: true);
                xml.Descendants(m + "oMath").First().AddBeforeSelf(field.Where(run =>
                    run.Element(w + "fldChar")?.Attribute(w + "fldCharType")?.Value != "separate"));
            }, diagnosticSink: diagnostics.Add);
        var owner = CachedMathParagraph(incomplete);
        TestAssert.True(!owner.FieldReferences.Single(f => f.Instruction!.Trim() == "REF StoredMath").HasSeparate,
            "An instruction-only field has no complete result owner.");
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(owner).Count);
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_COMPLEX_FIELD"),
            "An incomplete result boundary retains its unsupported-field diagnostic.");
    }

    public static void DocxEmptyStoredResultsMatchIndependentWordContentAndOrigins() =>
        CheckIndependentRateTimeOrigins("empty-field-results-office.json",
            ["empty-field-simple-empty.docx", "empty-field-complex-empty.docx"]);

    private static DocxDocument ReadEmptyStoredResult(string container, bool complex, OoxPdfDocxMarkupMode mode, bool baseline = false,
        Action<OoxPdfDiagnostic>? diagnosticSink = null) =>
        ReadMutation("CompositionProbes/math-fields-simple-" + container, mode, true, true,
            baseline ? null : (xml, w, m) => xml.Descendants(m + "oMath").First().AddBeforeSelf(
                CachedMathField(w, Array.Empty<XElement>(), complex, "REF EmptyStoredResult")),
            partOverride: container is "header" or "footer" ? "word/math-fields-" + container + ".xml" : null,
            diagnosticSink: diagnosticSink);
}
