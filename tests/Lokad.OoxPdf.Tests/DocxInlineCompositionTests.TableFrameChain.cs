using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionTableFrameChainMatchesIndependentWordGapsAndFollowingFlow()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-font table-frame controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("table-frame-chain-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = NestedMathCasePath(id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(path, mode);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var layout = renderer.InspectLayout(source);
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), layout.Pages.Count);
            var tables = layout.Pages[0].TableRows;
            var preceding = tables.Last(table => table.TableIndex == 0);
            var next = tables.First(table => table.TableIndex == 1);
            double gap = preceding.Y - next.Y - next.Height;
            TestAssert.True(Math.Abs(gap - row.GetProperty("TableGap").GetDouble()) < .5d,
                $"Distinct tables keep Word's shared body frame and authored intervening paragraph: {id}/{mode}/gap={gap}.");
            var page = renderer.RenderBlankPages(source, null, default).Single();
            var following = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Single(show => show.Text == "Following tables.");
            TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < .5d,
                $"Following prose retains the same qualified printed origin: {id}/{mode}/Y={following.Y}.");
        }
    }

    public static void DocxInlineCompositionTableFrameChainPreservesOneOriginAcrossMetricTransitions()
    {
        foreach (double? headingSize in new double?[] { null, 9d, 18d, 24d })
        foreach (var sizes in new[] { (First: 9d, Second: 24d), (First: 17d, Second: 12d) })
        foreach (double scale in new[] { .5d, .75d, .9d })
        foreach (bool separator in new[] { false, true })
        {
            DocxParagraph Prose(string text, double size) => DocxTests.CreateDocxLayoutParagraph(text, size, size)
                with { LineSpacingPoints = null };
            DocxTable Table(string text, double size) => new(null, [468d], [new DocxTableRow(
                [new DocxTableCell(text, [Prose(text, size)], null, null, null, null, [], DocxTableCellMargins.Empty)], null)]);
            var first = Table("First cell", sizes.First); var second = Table("Second cell", sizes.Second);
            var tail = Prose("Following body", 11d);
            var elements = new List<DocxBodyElement>();
            DocxParagraph? heading = headingSize is { } headingPoints ? Prose("Heading", headingPoints) : null;
            if (heading is not null)
            {
                elements.Add(new DocxParagraphElement(heading));
                elements.Add(new DocxParagraphElement(Prose("Body introduction", 11d)));
            }
            elements.Add(new DocxTableElement(first));
            if (separator) elements.Add(new DocxParagraphElement(Prose("Between tables", 11d)));
            elements.Add(new DocxTableElement(second)); elements.Add(new DocxParagraphElement(tail));
            var source = DocxTests.CreateLayoutTestDocument(elements, [first, second]) with
            { PageWidthPoints = 612d, PageHeightPoints = 792d, MarginLeftPoints = 72d, MarginRightPoints = 72d,
                MarginTopPoints = 72d, MarginBottomPoints = 72d };
            var natural = new DocxTests.FamilyWidthTextMeasurer();
            var final = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, natural, default).Pages.Single();
            var review = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: scale)
                .Create(source, new DocxTests.ScaledLayoutTextMeasurer(natural, scale, scale), default,
                    unscaledTextMeasurer: natural).Pages.Single();
            var before = final.Items.OfType<DocxTableRowLayout>().ToArray();
            var after = review.Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(2, before.Length); TestAssert.Equal(2, after.Length);
            double beforeGap = before[0].Y - before[1].Y - before[1].Height;
            double afterGap = after[0].Y - after[1].Y - after[1].Height;
            TestAssert.True(Math.Abs(afterGap - beforeGap * scale) < .000001d,
                $"A table does not reset its body's font origin: heading={headingSize}/cells={sizes}/scale={scale}/separator={separator}.");
            double Origin(DocxLayoutPage page, DocxTableRowLayout row) => heading is null ? row.Y + row.Height :
                page.Items.OfType<DocxTextLineLayout>().Single(line => ReferenceEquals(line.SourceParagraph, heading)).BaselineY;
            double Following(DocxLayoutPage page) => page.Items.OfType<DocxTextLineLayout>()
                .Single(line => ReferenceEquals(line.SourceParagraph, tail)).BaselineY;
            double expected = (Origin(final, before[0]) - Following(final)) * scale;
            double actual = Origin(review, after[0]) - Following(review);
            TestAssert.True(Math.Abs(expected - actual) < .000001d,
                "The following paragraph consumes the same physical frame across mixed cell/body metrics.");
        }
    }
}
