using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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

    public static void DocxEmptyStoredResultsMatchIndependentWordContentAndOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for independent empty-field Word origins.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "empty-field-results-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var pages = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], true).Where(show => pages[0].Fonts.Any(f =>
                f.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(f.Font, show.Font))).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("MathOrigins").EnumerateArray().ToArray();
            for (int index = 0; index < anchors.Length; index++)
            {
                var actual = math[4 * index]; var expected = anchors[index];
                TestAssert.Equal(expected.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "An empty result retains neighboring formula origins inside the existing one-point Word gate.");
            }
        }
    }

    private static DocxDocument ReadEmptyStoredResult(string container, bool complex, OoxPdfDocxMarkupMode mode, bool baseline = false,
        Action<OoxPdfDiagnostic>? diagnosticSink = null) =>
        ReadMutation("CompositionProbes/math-fields-simple-" + container, mode, true, true,
            baseline ? null : (xml, w, m) => xml.Descendants(m + "oMath").First().AddBeforeSelf(
                CachedMathField(w, Array.Empty<XElement>(), complex, "REF EmptyStoredResult")),
            partOverride: container is "header" or "footer" ? "word/math-fields-" + container + ".xml" : null,
            diagnosticSink: diagnosticSink);
}
