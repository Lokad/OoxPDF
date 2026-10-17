using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPageNotesRenumberOwnedRunsWithoutDiscardingMath()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (string composition in new[] { "", "compound-", "explicit-page-" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            bool compound = composition.Length != 0;
            var source = ReadMutation("CompositionProbes/math-page-notes-" + composition + typography, mode, true, true);
            TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(source), "Typography purposes retain their existing plain-body qualification.");
            var math = source.Paragraphs.SelectMany(p => p.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var numbered, out var pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Complete measured marker/note pages can number supported mixed paragraphs independently of typography.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            foreach (var paragraph in numbered.Paragraphs.Where(p => p.Equations.Count != 0))
            {
                var reference = paragraph.InlineReferences.Single();
                var run = TestAssert.NotNull(reference.DisplayRun);
                TestAssert.True(paragraph.Runs.Any(owner => ReferenceEquals(owner, run)) &&
                    paragraph.InlineContent!.OfType<DocxInlineText>().Any(item => ReferenceEquals(item.Run, run) && item.Length == run.Text.Length),
                    "Numbering publishes one new marker run shared by references and the ordered source view.");
                TestAssert.Equal(compound ? 2 : 1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
                if (compound)
                {
                    TestAssert.True(paragraph.BookmarkAnchors.Count != 0, "Numbering retains the authored bookmark metadata.");
                    var image = paragraph.InlineContent!.OfType<DocxInlinePicture>().Single();
                    TestAssert.True(ReferenceEquals(image.Image, paragraph.Images.Single()), "Numbering retains the actual picture owner.");
                }
            }
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(numbered, final, pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Changed labels retain complete source/note page ownership.");
        }
    }

    public static void DocxInlineCompositionPageNotesMatchIndependentLabelsAndRetainMathematics()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent page-note controls use installed Windows fonts.");
        CheckPageNoteMath(null);
    }

    public static void DocxInlineCompositionCompoundPageNotesRetainIndependentLabelsMathAndImages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent compound page-note controls use installed Windows fonts.");
        CheckPageNoteMath(null, "math-page-notes-compound-office.json");
    }

    public static void DocxInlineCompositionExplicitPageNotesRetainIndependentLabelsMathAndImages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent page dependency controls use installed Windows fonts.");
        CheckPageNoteMath(null, "math-page-notes-explicit-page-office.json");
    }

    internal static void CheckPageNoteMath(IFontResolver? resolver, string referenceFile = "math-page-notes-office.json")
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, referenceFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = ReadMutation("CompositionProbes/" + id, mode, true, true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(row.GetProperty("ExpectedEquations").GetInt32(), text.Split("ratetime", StringSplitOptions.None).Length - 1);
            if (row.TryGetProperty("ImagePlacements", out var expectedImages))
            {
                int images = pages.Sum(page => page.Images.Sum(image => System.Text.RegularExpressions.Regex.Matches(page.Content,
                    "/" + System.Text.RegularExpressions.Regex.Escape(image.ResourceName) + @"\s+Do\b").Count));
                TestAssert.Equal(expectedImages.GetInt32(), images);
            }
            foreach (string unchanged in new[] { "105", "108", "Public destination Alpha.", "Public destination Delta." }) TestAssert.Contains(unchanged, text);
            var labels = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                .Select(show => show.Text.Trim()).Where(label => label is "09" or "10").ToArray();
            TestAssert.Equal(string.Join(',', row.GetProperty("ExpectedDisplayedLabels").EnumerateArray().Select(x => x.GetString())), string.Join(',', labels));
            var links = pages.SelectMany(page => page.Annotations).Where(link => link.Destination is not null).ToArray();
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Measured numbering cannot discard supported mathematics.");
        }
    }
}
