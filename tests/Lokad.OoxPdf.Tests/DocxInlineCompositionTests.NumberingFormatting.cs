using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    private static readonly string[] NumberingFormattingCases =
    [
        "numbering-mark-plain-body", "numbering-mark-plain-cell", "numbering-mark-plain-header",
        "list-decoration-mark-inserted", "numbering-mark-math-body", "numbering-mark-math-cell",
        "numbering-mark-math-header", "numbering-mark-math-footer", "numbering-counter-second-mark"
    ];

    public static void DocxMarkFormattingPreservesNumberingTimelines()
    {
        foreach (string id in NumberingFormattingCases)
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (int index in new[] { 0, 1, 2 })
        {
            var baseline = NumberedCounterParagraphs(ReadNumberingFormatting(id, mode));
            var actual = NumberedCounterParagraphs(ReadNumberingFormatting(id, mode, index));
            TestAssert.Equal(string.Join('|', baseline.Select(p => p.ListLabel!.DisplayText)),
                string.Join('|', actual.Select(p => p.ListLabel!.DisplayText)));
            var insertion = actual.SelectMany(p => p.ParagraphMarkRun?.Revisions ?? [])
                .Single(r => r.Kind == DocxRevisionKind.Insertion);
            TestAssert.True(actual.Where(p => p.ListLabel!.NumberingChange is not null).All(p =>
                ReferenceEquals(insertion, p.ListLabel!.NumberingChange!.SourceMark)),
                "Formatting does not replace the actual inserted mark as the counter trigger.");
            var revisions = actual[index].ParagraphMarkRun!.Revisions;
            var formatting = revisions.Single(r => r.Kind == DocxRevisionKind.RunPropertiesChange);
            TestAssert.True(formatting.SourceElement == "rPrChange" && formatting.PropertyChangeFamily == DocxRevisionPropertyFamily.Run &&
                actual[index].Revisions.Any(r => ReferenceEquals(r, formatting)) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(actual[index], default),
                "The one retained formatting snapshot has its actual source identity.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() =>
                DocxInlineContent.HasSupportedParagraphRevisionOwners(actual[index], canceled.Token));
            if (mode != OoxPdfDocxMarkupMode.AllMarkup) continue;
            var duplicate = NumberedCounterParagraphs(ReadNumberingFormatting(id, mode, index, duplicate: true));
            TestAssert.Equal(string.Join('|', baseline.Take(index).Select(p => p.ListLabel!.DisplayText)),
                string.Join('|', duplicate.Take(index).Select(p => p.ListLabel!.DisplayText)));
            TestAssert.True(duplicate.Skip(index).All(p => p.ListLabel!.NumberingChange is null),
                "Ambiguous history preserves the established fallback from its own source boundary onward.");
        }
    }

    public static void DocxMarkFormattingPreservesNumberingPaintAndSourceContent()
    {
        using var golden = ReadNumberingCounterGolden("numbering-counters-office.json");
        foreach (string id in NumberingFormattingCases)
        foreach (bool final in new[] { true, false })
        {
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string modeName = final ? "final-preserve" : "all-word-compatible";
            var reference = golden.RootElement.GetProperty("Rows").EnumerateArray().Single(r =>
                r.GetProperty("Id").GetString() == id && r.GetProperty("Mode").GetString() == modeName);
            TestAssert.True(reference.GetProperty("QualifiesInsertedMarkTimeline").GetBoolean(),
                "Keep the independent-list continuation residual outside this qualification.");
            string expectedNumbers = string.Join('|', reference.GetProperty("Paragraphs").EnumerateArray()
                .SelectMany(p => p.GetProperty("Paints").EnumerateArray()).Select(p => p.GetProperty("Text").GetString()));
            var baseline = ReadNumberingCounterDocument(reference);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().Single();
            for (int index = 0; index < 3; index++)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(ReadNumberingFormatting(id, mode, index), diagnostics.Add, default)
                    .ToBlockingEnumerable().Single();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual, true, includeTextTransforms: true).ToArray();
                var numbers = shows.Where(s => Regex.IsMatch(s.Text.Trim(), @"^\d+[.)]$")).ToArray();
                TestAssert.Equal(expectedNumbers, string.Join('|', numbers.Select(s => s.Text.Trim())));
                TestAssert.True(numbers.All(s => s.X >= 0 && s.X < actual.Width && s.Y > 0 && s.Y < actual.Height),
                    "Both previous and current list numbers remain on the page.");
                TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(actual));
                static string MathText(Lokad.OoxPdf.Pdf.PdfPage page) => string.Concat(
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(s => page.Fonts.Any(f =>
                        f.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(f.Font, s.Font)))
                    .Select(s => s.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(MathText(control), MathText(actual));
                TestAssert.Equal(control.Images.Count, actual.Images.Count);
                TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                    "Preserving counter history cannot discard supported mathematics.");
                Lokad.OoxPdf.Pdf.PdfContentValidator.ValidatePage(actual, 0, default);
            }
        }
    }

    public static void DocxMarkFormattingKeepsNumberingDecorationSource()
    {
        foreach (string id in NumberingFormattingCases.Concat(new[] {
            "numbering-mark-decoration-mark-red-double", "numbering-mark-decoration-level-green-none",
            "numbering-mark-decoration-level-green-double" }))
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int index in new[] { 0, 1, 2 })
        {
            var control = NumberedCounterParagraphs(ReadNumberingFormatting(id, mode))[index];
            var actual = NumberedCounterParagraphs(ReadNumberingFormatting(id, mode, index))[index];
            static DocxTextRun LabelRun(DocxParagraph p) => DocxLayoutEngine.CreateListLabelRun(
                p.ListLabel!, p.Runs.First(), p.Runs.First().EffectiveProperties.FontSize);
            TestAssert.Equal(LabelRun(control).EffectiveProperties, LabelRun(actual).EffectiveProperties);
            TestAssert.True(actual.ListLabel!.DecorationSource is not null,
                "A known formatting snapshot cannot switch numbering to the first text run's hyperlink ink.");
        }
    }

    private static DocxDocument ReadNumberingFormatting(string id, OoxPdfDocxMarkupMode mode,
        int? index = null, bool duplicate = false) =>
        ReadMutation("CompositionProbes/" + id, mode, false, false, index is null ? null : (xml, w, _) =>
        {
            var paragraphs = xml.Descendants(w + "p").Where(p =>
                p.Element(w + "pPr")?.Element(w + "numPr") is not null).ToArray();
            TestAssert.Equal(3, paragraphs.Length);
            var properties = paragraphs[index.Value].Element(w + "pPr")!;
            var mark = properties.Element(w + "rPr");
            if (mark is null) { mark = new XElement(w + "rPr"); properties.Add(mark); }
            var history = new XElement(w + "rPrChange", new XAttribute(w + "id", "3400"),
                new XAttribute(w + "author", "Format Reviewer"),
                new XElement(w + "rPr", new XElement(w + "b", new XAttribute(w + "val", "1"))));
            mark.Add(history);
            if (duplicate) mark.Add(new XElement(history));
        }, partOverride: id.Contains("header", StringComparison.Ordinal) ? "word/header1.xml" :
            id is "list-decoration-mark-inserted" or "numbering-counter-second-mark" ||
                id.StartsWith("numbering-mark-decoration-", StringComparison.Ordinal) || id.EndsWith("footer", StringComparison.Ordinal)
                ? "word/footer1.xml" : null);
}
