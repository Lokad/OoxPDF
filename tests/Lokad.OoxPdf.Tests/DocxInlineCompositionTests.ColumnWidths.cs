using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFixedPictureSkipsUnfitContinuationColumn()
    {
        var source = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false);
        var original = source.Paragraphs.First(p => p.Equations.Count != 0);
        var equation = original.Equations.Single();
        var image = new DocxInlineImage(30d, 32d, "image/png", null, "boundary.png");
        var paragraph = original with { Runs = [], Images = [image], LineSpacingPoints = null, LineSpacingFactor = 1d,
            Spacing = original.Spacing with { LineRuleValue = "auto", LineValue = "240" },
            InlineContent = [new DocxInlineEquation(equation), new DocxInlineBreak(original.Runs[0], 0), new DocxInlinePicture(image)] };
        var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
            { [equation] = new(equation, new(18d, 10d, 4d, []), 10d, 16d) };
        var document = new DocxDocument(120d, 60d, 10d, 10d, 10d, 10d, DocxPageSettings.Empty,
            [], [], [], [new DocxParagraphElement(paragraph)], [], [])
        {
            FinalSectionBreak = new(DocxPageSettings.Empty, null, "2", "0", "200", [new("1600", "200"), new("200", null)])
        };
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TrackingMeasurer(), default, mathEquations: equations);
        TestAssert.Equal(2, layout.Pages.Count);
        var placed = layout.Pages.SelectMany((page, pi) => page.Items.OfType<DocxInlineImageLayout>().Select(item => (item, pi))).Single();
        TestAssert.Equal(1, placed.pi);
        TestAssert.True(placed.item.X < 90d, "The 30pt image must skip the 10pt continuation column without consuming its source.");
        TestAssert.Equal(1, layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()).Sum(line => line.InlineMath.Count));
    }
}
