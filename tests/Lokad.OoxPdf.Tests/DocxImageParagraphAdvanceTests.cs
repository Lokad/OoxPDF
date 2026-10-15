using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxImageParagraphAdvanceTests
{
    public static void DocxBodyImageParagraphAdvanceUsesAuthoredSpacingWithoutExtraPadding()
    {
        foreach ((double height, double after, double before, string type) in new[]
        {
            (18d, 0d, 0d, "image/png"), (36d, 8d, 0d, "image/svg+xml"),
            (24d, 12d, 4d, "image/png"), (18d, 8d, 12d, "image/svg+xml")
        })
        {
            DocxParagraph first = ImageParagraph(height, type) with { SpacingAfterPoints = after };
            DocxParagraph second = ImageParagraph(height, type) with { SpacingBeforePoints = before };
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(first), new DocxParagraphElement(second)], []);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            DocxInlineImageLayout[] images = layout.Pages.Single().Items.OfType<DocxInlineImageLayout>().ToArray();
            TestAssert.Equal(2, images.Length);
            double visibleGap = images[0].Y - (images[1].Y + images[1].Height);
            TestAssert.True(Math.Abs(visibleGap - Math.Max(after, before)) < .000001,
                "Image lines must use the collapsed paragraph gap without another fixed padding.");
        }
    }

    public static void DocxBodyImageKeepNextEstimateSharesTheOccupiedImageAdvance()
    {
        foreach ((double nextHeight, int expectedPages) in new[] { (20d, 1), (28d, 2) })
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Fill", 10d, 10d);
            DocxParagraph image = ImageParagraph(24d, "image/png") with
            {
                KeepRules = new DocxParagraphKeepRules(true, null, null, null, null, null)
            };
            DocxParagraph next = DocxTests.CreateDocxLayoutParagraph("Next", 10d, nextHeight);
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(filler), new DocxParagraphElement(image), new DocxParagraphElement(next)], [])
                with { PageHeightPoints = 80d };
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            TestAssert.Equal(expectedPages, layout.Pages.Count);
            var imagePage = layout.Pages.Single(page => page.Items.OfType<DocxInlineImageLayout>().Any());
            TestAssert.True(imagePage.Items.OfType<DocxTextLineLayout>().Any(line => line.Text == "Next"),
                "The image must stay with its following paragraph when the pair moves.");
        }
    }

    private static DocxParagraph ImageParagraph(double height, string type) =>
        DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with
        {
            Runs = [], Images = [new DocxInlineImage(24d, height, type, [], null)]
        };
}
