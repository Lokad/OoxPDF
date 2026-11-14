using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNumberedTextRowsRetainContinuationMetricsAndPages()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string content in new[] { "plain", "picture", "math", "mixed" })
        {
            var control = RenderTextRowMetrics(content, "ordinary", mode);
            var source = ReadNumberedTextRows(content, "ordinary", mode);
            var label = source.Paragraphs[0].ListLabel;
            var actual = RenderTextRowMetricDocument(source, mode);
            TestAssert.Equal(1, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            AssertNumberedTextRowContent(actual, content);
            TestAssert.True(ReferenceEquals(label, source.Paragraphs[0].ListLabel), "Measurement preserves the authored numbering owner.");
        }
    }

    public static void DocxNumberedTextRowsShareKeepAndWidowAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string admission in new[] { "keepLines", "keepNext", "widow" })
        {
            var control = RenderTextRowMetrics("picture", admission, mode);
            var actual = RenderTextRowMetricDocument(ReadNumberedTextRows("picture", admission, mode), mode);
            TestAssert.Equal(admission == "keepLines" ? 1 : 2, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            int headingPage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Large heading", StringComparison.Ordinal));
            TestAssert.Equal(admission == "keepNext" ? 1 : 0, headingPage);
            AssertNumberedTextRowContent(actual, "picture");
        }
    }

    public static void DocxNumberedTextRowsRetainDefaultAndExplicitListSpacing()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string content in new[] { "plain", "picture" })
        foreach (bool explicitSpacing in new[] { false, true })
        {
            var original = ReadNumberedTextRows(content, "ordinary", OoxPdfDocxMarkupMode.Final);
            var paragraph = original.Paragraphs[0] with
            {
                LineSpacingFactor = 1.15d,
                Spacing = original.Paragraphs[0].Spacing with { LineValue = explicitSpacing ? "276" : null, LineRuleValue = "auto" }
            };
            var source = original with { BodyElements = [new DocxParagraphElement(paragraph)], Paragraphs = [paragraph] };
            var plan = DocxFontPlan.Create(source, resolver, default);
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(plan, resolver.Resolve(new FontRequest("DejaVu Sans", false, false)),
                default, resolver).GetAwaiter().GetResult();
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, measurer, default);
            var rows = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>()
                .Where(line => line.Text.StartsWith("Small row", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(7, rows.Length);
            foreach (var row in rows)
            {
                TestAssert.Equal(11d, row.FontSize);
                TestAssert.Equal(explicitSpacing ? 1.15d : 1.16d, row.EffectiveLineSpacingFactor!.Value);
                TestAssert.True(ReferenceEquals(paragraph, row.SourceParagraph), "The source paragraph retains list spacing ownership.");
            }
            TestAssert.Equal(1, layout.Pages.Count);
        }
    }

    private static DocxDocument ReadNumberedTextRows(string content, string admission, OoxPdfDocxMarkupMode mode) =>
        ReadTextRowMetrics(content, admission, mode, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(w + "t").Any(text => text.Value == "Large heading"));
            paragraph.Element(w + "pPr")!.Add(new XElement(w + "numPr",
                new XElement(w + "ilvl", new XAttribute(w + "val", "0")), new XElement(w + "numId", new XAttribute(w + "val", "1"))));
        });

    private static void AssertNumberedTextRowContent(PdfPage[] pages, string content)
    {
        string text = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            honorActualText: false)).Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
        TestAssert.Equal(1, Regex.Matches(text, @"1\.").Count);
        for (int row = 1; row <= 7; row++) TestAssert.True(text.Contains($"Small row {row:00}", StringComparison.Ordinal), "Every authored continuation row paints once.");
        TestAssert.Equal(content is "picture" or "mixed" ? 1 : 0, pages.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
        string math = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false)
            .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
            .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
        TestAssert.Equal(content is "math" or "mixed" ? "x" : "", math);
    }
}
