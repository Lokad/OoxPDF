using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxHeaderFooterTests
{
    public static void DocxStaticTokenWrappingMatchesBodyAcrossZeroAndTinyInsets()
    {
        foreach (string text in new[] { new string('W', 80), string.Concat(Enumerable.Repeat("Alpha Beta ", 8)),
            string.Concat(Enumerable.Repeat("ABCDEFGHIJ/", 8)), "ABCDEFGHIJKLM\u00ADNOPQRSTUVWXYZ" + new string('W', 60),
            new string('W', 40) + "\u00A0" + new string('W', 40) })
        foreach (double? inset in new double?[] { null, 0, .05 })
        foreach (bool split in new[] { false, true })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph(text, 10, 12) with
            {
                LineSpacingPoints = null,
                Indent = DocxParagraphIndent.Empty with { LeftPoints = inset, LeftValue = inset is null ? null : inset == 0 ? "0" : "1" }
            };
            if (split) paragraph = paragraph with { Runs = [paragraph.Runs[0] with { Text = text[..20] }, paragraph.Runs[0] with { Text = text[20..] }] };
            paragraph = paragraph with { Runs = paragraph.Runs.Select((run, index) => run with { SourceRunIndex = index }).ToArray() };
            paragraph = paragraph with { InlineContent = paragraph.Runs.Select(run => (DocxInlineItem)new DocxInlineText(run, 0, run.Text.Length)).ToArray() };
            var settings = DocxPageSettings.Empty with
            {
                HeaderDistancePoints = 12, FooterDistancePoints = 12,
                HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph] },
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph] }
            };
            var document = new DocxDocument(160, 1000, 10, 10, 300, 300, settings, [], [], [],
                [new DocxParagraphElement(paragraph)]);
            DocxLayoutPage page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
            var body = page.Items.OfType<DocxTextLineLayout>().ToArray();
            foreach (DocxStoryKind kind in new[] { DocxStoryKind.Header, DocxStoryKind.Footer })
            {
                var story = page.StaticTextLines.Where(line => line.Story?.Kind == kind).ToArray();
                TestAssert.True(body.Select(line => line.Text.TrimEnd(' ')).SequenceEqual(story.Select(line => line.Text.TrimEnd(' '))),
                    $"{kind} uses the same token breaks as body text with {inset}pt inset and split={split}.");
                if (text.Contains('\u00A0')) TestAssert.Equal(1, story.Length); // The existing protected-compound overflow fallback remains.
                else
                {
                    TestAssert.True(story.Length > 1, "Ordinary overlong content wraps in every story.");
                    TestAssert.True(story.SelectMany(line => line.Segments).Where(segment => !string.IsNullOrWhiteSpace(segment.Text))
                        .All(segment => segment.X >= 0 && segment.X + segment.Width <= page.Width + 1e-7), "Visible static text must remain inside the physical page.");
                }
            }
        }
    }
}
