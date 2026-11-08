using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxReviewTableFrameTests
{
    public static void DocxReviewAuthoredRowsStayInsideTheirPrintedPages()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
                     OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach (bool tableFirst in new[] { true, false })
        foreach (string? heightRule in new string?[] { null, "atLeast", "exact" })
        foreach (var paper in new[] { (Width: 432d, Height: 432d), (Width: 612d, Height: 792d) })
        {
            var source = CreateDocument(paper.Width, paper.Height, tableFirst, heightRule, nestedHeader: false);
            AssertVisibleRows(source, geometry);
        }
    }

    public static void DocxReviewNestedHeaderAndContinuationSharePrintedCapacity()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
                     OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach (bool tableFirst in new[] { true, false })
        {
            var source = CreateDocument(432d, 432d, tableFirst, null, nestedHeader: true);
            AssertVisibleRows(source, geometry);
            var renderer = Renderer(geometry);
            var layout = renderer.InspectLayout(source);
            TestAssert.True(layout.Pages.Count > 1 && layout.Pages.All(page => page.TableRows.Any(row => row.IsHeader)),
                "Every continued fragment keeps the nested source header in its own admitted frame.");
            TestAssert.Equal(50, layout.Pages.Sum(page => page.TableRows.Count(row => !row.IsHeader)));
        }
    }

    private static void AssertVisibleRows(DocxDocument source, OoxPdfDocxMarkupGeometryMode geometry)
    {
        var renderer = Renderer(geometry);
        var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
        var rows = pages.SelectMany((page, index) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .GroupBy(show => show.Y).Select(group => group.OrderBy(show => show.X).ToArray())
                .Where(line => string.Concat(line.Select(show => show.Text)).StartsWith("ROW", StringComparison.Ordinal))
                .Select(line => (Page: index, Y: line[0].Y, Size: line[0].Size,
                    Text: string.Concat(line.Select(show => show.Text))))).ToArray();
        TestAssert.Equal(50, rows.Length);
        TestAssert.Equal(50, rows.Select(row => row.Text.Trim()).Distinct(StringComparer.Ordinal).Count());
        foreach (var row in rows)
        {
            TestAssert.True(row.Y - row.Size > 0d && row.Y + row.Size < pages[row.Page].Height,
                $"Admitted rows must remain on the paper after painting applies its frame. " +
                $"{geometry}, {source.PageWidthPoints}x{source.PageHeightPoints}, " +
                $"page {row.Page}, {row.Text}, Y={row.Y}, size={row.Size}.");
        }
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry) =>
        new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);

    private static DocxDocument CreateDocument(double width, double height, bool tableFirst,
        string? heightRule, bool nestedHeader)
    {
        var anchor = Paragraph("Anchor") with
        {
            InlineReferences = [new(DocxRelatedStoryKind.Comment, "1", null, DisplayText: null,
                SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 0)],
            CommentRanges = [new("1", 0, 0, 1, 6, 1, 0)]
        };
        var rows = new List<DocxTableRow>();
        if (nestedHeader)
        {
            var child = new DocxTable("fixed", [160d], [new([Cell(Paragraph("Child"))], null)]);
            var cell = Cell(Paragraph(string.Empty)) with
            { BodyElements = [new DocxTableElement(child), new DocxParagraphElement(Paragraph(string.Empty))] };
            rows.Add(new([cell], null) { IsHeader = true, CantSplit = true });
        }
        for (int i = 0; i < 50; i++) rows.Add(new([Cell(Paragraph($"ROW{i:00}"))], 20d)
            { CantSplit = true, HeightRuleValue = heightRule });
        var table = new DocxTable("fixed", [160d], rows);
        return DocxTests.CreateAllMarkupWrapProbeDocument([anchor]) with
        {
            PageWidthPoints = width,
            PageHeightPoints = height,
            BodyElements = tableFirst
                ? [new DocxTableElement(table), new DocxParagraphElement(anchor)]
                : [new DocxParagraphElement(anchor), new DocxTableElement(table)],
            FallbackTables = [table],
            RelatedStories = [new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(Paragraph("Public comment"))], [], [], null)]
        };
    }

    private static DocxParagraph Paragraph(string text) => DocxTests.CreateDocxLayoutParagraph(text, 12d, 20d)
        with { LineSpacingPoints = null };
    private static DocxTableCell Cell(DocxParagraph paragraph) =>
        new(paragraph.Runs[0].Text, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
}
