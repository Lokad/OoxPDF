using System.Text;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMarkerRowsRetainLocalMetricsVisibleContentAndNoteTargets()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string content in new[] { "plain", "picture", "math", "mixed" })
        {
            var control = MarkerRowOrigins(RenderTextRowMetrics(content, "ordinary", mode));
            foreach (bool endnote in new[] { false, true })
            foreach (bool before in new[] { false, true })
            foreach (bool sameRun in new[] { false, true })
            {
                string label = endnote ? "i" : "1";
                var source = ReadNoteTextRows(content, "ordinary", mode, endnote, sameRun, markerRow: 4, beforeText: before);
                var actual = RenderTextRowMetricDocument(source, mode);
                TestAssert.Equal(1, actual.Length);
                AssertMarkerRowOrigins(control, MarkerRowOrigins(actual));
                string prose = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    true, includeTextTransforms: true, honorActualText: false,
                    fontResourceFilter: name => !name.StartsWith("Eq", StringComparison.Ordinal))).Select(show => show.Text));
                string expected = "Largeheading" + string.Concat(Enumerable.Range(1, 7).Select(row =>
                    (row == 4 && before ? label : "") + $"Smallrow{row:00}" + (row == 4 && !before ? label : ""))) + label + "Sourcenote";
                TestAssert.Equal(expected, string.Concat(prose.Normalize(NormalizationForm.FormKC).Where(ch => !char.IsWhiteSpace(ch))));
                string math = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    true, includeTextTransforms: true, honorActualText: false,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))).Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(content is "math" or "mixed" ? "x" : "", math);
                TestAssert.True(actual[0].Annotations.Any(annotation => annotation.Destination?.PageIndex == 0),
                    "The measured reference retains its actual same-page note target.");
                TestAssert.Equal(content is "picture" or "mixed" ? 1 : 0, actual.Sum(page => page.Images.Count));
            }
        }
    }

    public static void DocxMarkerRowsShareKeepWidowAndCompleteNoteAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string admission in new[] { "keepLines", "keepNext", "widow" })
        {
            var control = RenderTextRowMetricDocument(ReadNoteTextRows("plain", admission, mode, false, false, markerRow: 4), mode);
            var actual = RenderTextRowMetricDocument(ReadNoteTextRows("picture", admission, mode, false, true, markerRow: 4), mode);
            TestAssert.Equal(2, actual.Length);
            var rows = MarkerRowOrigins(actual);
            AssertMarkerRowOrigins(MarkerRowOrigins(control), rows);
            int notePage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Source note", StringComparison.Ordinal));
            TestAssert.Equal(rows[3].Page, notePage);
            TestAssert.True(actual[rows[3].Page].Annotations.Any(annotation => annotation.Destination?.PageIndex == notePage),
                "Keep and widow admission preserve the actual late marker's note page.");
        }
    }

    public static void DocxMarkerRowsKeepActualLargeStrutsDeclaredSlotsAndUnknownFallbacks()
    {
        foreach (string variant in new[] { "ordinary", "large-marker", "exact", "unknown-source", "missing-run" })
        {
            var original = ReadNoteTextRows("picture", "ordinary", OoxPdfDocxMarkupMode.Final, false, false, markerRow: 4);
            var paragraph = original.Paragraphs[0]; var reference = paragraph.InlineReferences.Single();
            if (variant == "large-marker")
            {
                var marker = TestAssert.NotNull(reference.DisplayRun); var enlarged = marker with { FontSize = 36d };
                paragraph = paragraph with { Runs = paragraph.Runs.Select(run => ReferenceEquals(run, marker) ? enlarged : run).ToArray(),
                    InlineReferences = [reference with { DisplayRun = enlarged }] };
                paragraph = paragraph with { InlineContent = DocxInlineContent.FromPrefix(paragraph.Runs, paragraph.Images, default) };
            }
            if (variant == "exact") paragraph = paragraph with { LineSpacingPoints = 12d,
                Spacing = paragraph.Spacing with { LineValue = "240", LineRuleValue = "exact" } };
            if (variant is "unknown-source" or "missing-run") paragraph = paragraph with { InlineReferences =
                [variant == "unknown-source" ? reference with { SourceRunIndex = -1 } : reference with { DisplayRun = null }] };
            var source = original with { BodyElements = [new DocxParagraphElement(paragraph)] };
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var plan = DocxFontPlan.Create(source, resolver, default);
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(plan, resolver.Resolve(new FontRequest("DejaVu Sans", false, false)),
                default, resolver).GetAwaiter().GetResult();
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, measurer, default);
            var rows = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>()
                .Where(line => line.Text.StartsWith("Small row", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(7, rows.Length);
            if (variant == "ordinary") TestAssert.True(rows.All(row => row.FontSize == 11d), "The actual marker is measured with its short row.");
            else if (variant == "large-marker")
            {
                TestAssert.Equal(36d, rows[3].FontSize);
                TestAssert.True(rows.Where((_, index) => index != 3).All(row => row.FontSize == 11d), "Only the actual large marker grows its own row.");
            }
            else if (variant == "exact") TestAssert.True(rows.All(row => row.LineHeight == 12d), "Authored exact slots remain exact.");
            else TestAssert.True(rows.All(row => row.FontSize == 36d), "Ambiguous source owners retain their conservative frame.");
        }
    }

    private static (int Page, double Y)[] MarkerRowOrigins(PdfPage[] pages) => Enumerable.Range(1, 7).Select(row =>
        pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, includeTextTransforms: true,
                honorActualText: false, fontResourceFilter: name => !name.StartsWith("Eq", StringComparison.Ordinal))
            .GroupBy(show => show.Y).Where(group => string.Concat(group.Select(show => show.Text)).Contains($"Small row {row:00}", StringComparison.Ordinal))
            .Select(group => (Page: index, Y: group.Key))).Single()).ToArray();

    private static void AssertMarkerRowOrigins((int Page, double Y)[] control, (int Page, double Y)[] actual)
    {
        for (int index = 0; index < control.Length; index++)
        {
            TestAssert.Equal(control[index].Page, actual[index].Page);
            TestAssert.True(Math.Abs(control[index].Y - actual[index].Y) < .001d,
                $"Row{index + 1} retains actual text/reference metrics independently of a distant heading.");
        }
    }
}
