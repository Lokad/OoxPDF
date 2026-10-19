using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxFollowingFlowMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent following-flow origins use installed Windows fonts.");
        using var gold = ReadInlineBoxMathGold("inline-box-following-flow-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = ReadMutation("InlineBoxMath/" + id, mode, false, false);
            var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
            var following = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true).GroupBy(show => show.Y).Select(line =>
                    (Page: index + 1, Height: page.Height, Y: line.Key, X: line.Min(show => show.X),
                    Text: string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)).Trim())))
                .Single(line => line.Text == row.GetProperty("FollowingText").GetString());
            var expected = row.GetProperty("WordFollowingAnchor");
            TestAssert.Equal(expected.GetProperty("Page").GetInt32(), following.Page);
            TestAssert.True(Math.Abs(following.X - expected.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(following.Height - following.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                id + "/" + row.GetProperty("Mode").GetString() + ": following prose retains one-point independent Word origins.");
        }
    }

    public static void DocxInlineBoxOccupiedFramesOwnZeroSpacingAdmissionInBodyAndCells()
    {
        var box = new DocxInlineTextBox("1371600", "457200")
        { BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Inside", 10d, 12d))] };
        var owner = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d,
            null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) { InlineTextBoxes = [box] };
        var following = DocxTests.CreateDocxLayoutParagraph("Following", 12d, 12d);
        var paragraphs = new[] { owner, owner with { }, following };
        foreach (bool cell in new[] { false, true })
        {
            var table = new DocxTable(null, [200d], [new DocxTableRow([
                new DocxTableCell(string.Empty, paragraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], null)]);
            DocxDocument document = new(300d, 156d, 36d, 36d, 36d, 36d, DocxPageSettings.Empty, [], [], [],
                cell ? [new DocxTableElement(table)] : paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(),
                [], cell ? [table] : []);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.Equal(1, layout.Pages.Count);
            var row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().SingleOrDefault();
            var boxes = cell ? row!.Cells.Single().InlineTextBoxes : layout.Pages[0].Items.OfType<DocxInlineTextBoxLayout>().ToArray();
            TestAssert.Equal(2, boxes.Count);
            TestAssert.True(Math.Abs(boxes[0].BoxTop - boxes[0].BoxHeight - boxes[1].BoxTop) < .001d,
                "Zero paragraph spacing joins complete box frames without an extra gap.");
            if (cell) TestAssert.True(Math.Abs(row!.Height - 84d) < .001d,
                "Cell measurement reserves two36pt boxes and the12pt following line once.");
        }
    }
}
