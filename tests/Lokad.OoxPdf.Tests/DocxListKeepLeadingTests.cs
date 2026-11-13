using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxListKeepLeadingTests
{
    public static void DocxKeptListLabelsShareFirstLineLeadingWithPlacement()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (double labelSize in new[] { 12d, 36d, 72d })
        foreach (bool boundary in new[] { false, true })
        foreach (bool keep in new[] { false, true })
        {
            var source = Document(mode, labelSize, boundary, keep, counter: false);
            var pages = Renderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            int expectedHeading = boundary && labelSize > 12d && keep ? 1 : 0;
            int expectedFollowing = boundary && labelSize > 12d ? 1 : 0;
            TestAssert.Equal(expectedFollowing + 1, pages.Length);
            TestAssert.Equal(expectedHeading, PageContaining(pages, "BIGLABELHEADING"));
            TestAssert.Equal(expectedFollowing, PageContaining(pages, "FOLLOWINGBODY"));
            int labels = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true))
                .Count(show => show.Text == "1." && Math.Abs(show.Size - labelSize) < .001d);
            TestAssert.True(labels == 1, "The source list label paints exactly once at its authored size.");
        }
    }

    public static void DocxKeptListLabelCountersFollowTheAdmittedFrame()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = Document(mode, 72d, boundary: true, keep: true, counter: true);
            var renderer = Renderer(mode);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(1, PageContaining(pages, "BIGLABELHEADINGPAGE2"));
            TestAssert.Equal(1, PageContaining(pages, "FOLLOWINGBODY"));
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupMode mode) =>
        new(new PortableRenderingTests.PinnedResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

    private static int PageContaining(IReadOnlyList<Lokad.OoxPdf.Pdf.PdfPage> pages, string expected) =>
        Enumerable.Range(0, pages.Count).Single(index => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[index], true)
            .SelectMany(show => show.Text).Where(c => !char.IsWhiteSpace(c))).Contains(expected, StringComparison.Ordinal));

    private static DocxDocument Document(OoxPdfDocxMarkupMode mode, double labelSize, bool boundary, bool keep, bool counter)
    {
        var heading = DocxTests.CreateDocxLayoutParagraph(counter ? "BIG LABEL HEADING PAGE {PAGE}" : "BIG LABEL HEADING", 12d, 12d) with
        {
            LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
            KeepRules = DocxParagraphKeepRules.Empty with { KeepNext = keep },
            ListLabel = new DocxListLabel("1.", "decimal", "%1.", "space", "1", 0,
                DocxNumberingIndent.Empty with { LeftPoints = 72d, HangingPoints = 72d },
                DocxTextRunStyle.Empty with { FontSize = labelSize, FontFamily = "Arial" })
        };
        var following = DocxTests.CreateDocxLayoutParagraph("FOLLOWING BODY", 12d, 24d) with
        { Spacing = DocxParagraphSpacing.Empty with { LineValue = "480", LineRuleValue = "exact" } };
        var paragraphs = new List<DocxParagraph>();
        if (boundary)
        {
            double filler = labelSize >= 72d ? 568d : 596d;
            paragraphs.Add(DocxTests.CreateDocxLayoutParagraph(string.Empty, 12d, filler) with
            { Spacing = DocxParagraphSpacing.Empty with { LineValue = (filler * 20d).ToString(System.Globalization.CultureInfo.InvariantCulture), LineRuleValue = "exact" } });
        }
        paragraphs.Add(heading); paragraphs.Add(following);
        return DocxTests.CreateLayoutTestDocument(paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray(), []) with
        {
            PageWidthPoints = 612d, PageHeightPoints = 792d, MarginTopPoints = 72d, MarginBottomPoints = 72d,
            MarginLeftPoints = 72d, MarginRightPoints = 72d, MarkupMode = mode
        };
    }
}
