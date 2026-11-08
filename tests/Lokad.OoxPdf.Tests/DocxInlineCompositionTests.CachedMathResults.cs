using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCachedMathResultsKeepSourceOwnershipAcrossPlacements()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var source = ReadCachedMathResult(container, complex, mode);
            var paragraph = CachedMathParagraph(source);
            var field = paragraph.FieldReferences.Single(f => f.Instruction!.Trim() == "REF StoredMath");
            TestAssert.True(field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder,
                "A materialized equation is a stored result in either field form.");
            TestAssert.Equal(0, field.ResultRunCount); // Do not invent text runs for mathematics.
            TestAssert.Equal(2, paragraph.Equations.Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(2, TestAssert.NotNull(paragraph.InlineContent).OfType<DocxInlineEquation>().Count());
        }
    }

    public static void DocxCachedMathResultsPreservePaintNavigationAndDiagnostics()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var before = ReadCachedMathResult(container, complex, mode, baseline: true);
            var after = ReadCachedMathResult(container, complex, mode);
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var expected = renderer.RenderPagesAsync(before, null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(after, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int page = 0; page < expected.Length; page++)
            {
                TestAssert.Equal(expected[page].Content, actual[page].Content);
                TestAssert.Equal(expected[page].Fonts.Count, actual[page].Fonts.Count);
                TestAssert.Equal(expected[page].Images.Count, actual[page].Images.Count);
                TestAssert.True(expected[page].Annotations.SequenceEqual(actual[page].Annotations),
                    "A cached field wrapper leaves painted geometry and navigation unchanged.");
            }
            TestAssert.True(diagnostics.All(d => d.Id is not "DOCX_UNSUPPORTED_EQUATION" and not "DOCX_UNSUPPORTED_COMPLEX_FIELD"),
                "Stored supported math must not receive an ignored-content diagnostic.");
            string math = string.Concat(actual.SelectMany(p => DocxCommentsTests.ReadEmbeddedGlyphTextShows(p, true)
                .Where(show => p.Fonts.Any(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(f.Font, show.Font)))
                .Select(show => show.Text))).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal("ratetimeratetime", math);
        }
    }

    public static void DocxCachedMathResultsKeepNestedLocksAndMetadataOwners()
    {
        foreach (bool complex in new[] { false, true })
        foreach (string wrapper in new[] { "customXml", "smartTag", "sdt" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-fields-simple-body", mode, true, true, (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                var inner = CachedMathField(w, new XElement(equation), complex);
                var outer = CachedMathField(w, inner, !complex, "NUMPAGES");
                var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), outer);
                var content = wrapper == "sdt" ? new XElement(w + "sdt", new XElement(w + "sdtContent", link))
                    : new XElement(w + wrapper, link);
                equation.ReplaceWith(content);
            });
            var paragraph = CachedMathParagraph(document);
            TestAssert.Equal(4, paragraph.FieldReferences.Count);
            TestAssert.True(paragraph.FieldReferences.All(f => f.HasCachedResult && f.RendersCachedResult && !f.UsesPlaceholder),
                "Nested locks keep their own stored result, independently of transparent source wrappers.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.SelectMany(p => p.Annotations).Any(a => a.Destination is not null),
                "The enclosing equation link still resolves its retained bookmark.");
            TestAssert.True(diagnostics.All(d => d.Id is not "DOCX_UNSUPPORTED_EQUATION" and not "DOCX_UNSUPPORTED_COMPLEX_FIELD"),
                "The diagnostic scan follows the same supported metadata owners as the reader.");
        }
    }

    public static void DocxCachedMathResultsKeepEmptyComputedAndExcludedFallbacks()
    {
        foreach (bool complex in new[] { false, true })
        {
            var empty = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, m) => xml.Descendants(m + "oMath").First().AddBeforeSelf(CachedMathField(w, Array.Empty<XElement>(), complex)));
            var emptyParagraph = CachedMathParagraph(empty);
            var emptyField = emptyParagraph.FieldReferences.Single(f => f.Instruction!.Trim() == "REF StoredMath");
            TestAssert.True(!emptyField.HasCachedResult && !emptyField.RendersCachedResult,
                "An absent result does not become materialized content.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(emptyParagraph).Count);
            var computed = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                equation.ReplaceWith(CachedMathField(w, new XElement(equation), complex, "PAGE", locked: false));
            });
            var field = CachedMathParagraph(computed).FieldReferences.Single(f => f.Kind == DocxFieldKind.Page);
            TestAssert.True(field.UsesPlaceholder && !field.RendersCachedResult, "Computed counters retain their existing replacement owner.");
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original })
            {
                var revision = ReadMutation("CompositionProbes/math-fields-simple-body", mode, false, false, (xml, w, m) =>
                {
                    var equation = xml.Descendants(m + "oMath").First();
                    equation.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "811"),
                        new XAttribute(w + "author", "Public"), new XAttribute(w + "date", "2026-01-01T00:00:00Z"),
                        CachedMathField(w, new XElement(equation), complex)));
                });
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final ? 2 : 1, CachedMathParagraph(revision).Equations.Count);
            }
        }
    }

    public static void DocxCachedMathResultsMatchIndependentWordContentAndOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent cached-math Word origins.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "cached-math-results-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(input, mode);
            var pages = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], true).Where(show => pages[0].Fonts.Any(f =>
                f.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(f.Font, show.Font))).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            var anchors = row.GetProperty("MathOrigins").EnumerateArray().ToArray();
            TestAssert.Equal(16, math.Length);
            for (int index = 0; index < anchors.Length; index++)
            {
                var actual = math[4 * index]; var expected = anchors[index];
                TestAssert.Equal(expected.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "Stored math retains the existing independent one-point Word origin gate: " + input + "/" + mode);
            }
        }
    }

    private static DocxDocument ReadCachedMathResult(string container, bool complex, OoxPdfDocxMarkupMode mode, bool baseline = false) =>
        ReadMutation("CompositionProbes/math-fields-simple-" + container, mode, true, true,
            baseline ? null : (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                equation.ReplaceWith(CachedMathField(w, new XElement(equation), complex));
            }, partOverride: container is "header" or "footer" ? "word/math-fields-" + container + ".xml" : null);

    private static DocxParagraph CachedMathParagraph(DocxDocument document) =>
        DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).Concat(document.EquationParts.SelectMany(part =>
            DocxBlockTraversal.EnumerateBodyParagraphs(part.BodyElements))).First(p => p.Equations.Count != 0);

    private static OoxPdfDocxMarkupGeometryMode CachedMathGeometry(OoxPdfDocxMarkupMode mode) => mode == OoxPdfDocxMarkupMode.Final
        ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;

    private static XElement[] CachedMathField(XNamespace w, object content, bool complex, string instruction = "REF StoredMath", bool locked = true)
    {
        if (!complex) return [new XElement(w + "fldSimple", new XAttribute(w + "instr", " " + instruction + " "),
            new XAttribute(w + "fldLock", locked ? "true" : "false"), content)];
        return [new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin"),
                new XAttribute(w + "fldLock", locked ? "true" : "false"))),
            new XElement(w + "r", new XElement(w + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), " " + instruction + " ")),
            new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate"))),
            new XElement(w + "customXml", content),
            new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end")))];
    }
}
