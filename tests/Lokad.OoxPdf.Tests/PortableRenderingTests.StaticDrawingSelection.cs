using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingStaticDrawingSelectionKeepsPaintAndFooterInk()
    {
        const string footer = "Public synthetic composition portfolio";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false);
            var baseline = RenderAlignmentProbe(document, mode);
            var expectedFooters = baseline.Select(page => ReadPaintedSourceLabel(page, footer)).ToArray();
            var drawing = CreateStaticSelectionPicture(document);
            var inactive = document with
            {
                PageSettings = document.PageSettings with { TitlePage = false, EvenAndOddHeaders = false }
            };
            foreach (bool header in new[] { true, false })
            {
                foreach (string variant in new[] { "first", "even" })
                {
                    var unused = RenderAlignmentProbe(WithStaticSelectionDrawing(inactive, drawing, header, variant), mode);
                    TestAssert.Equal(baseline.Length, unused.Length);
                    for (int index = 0; index < unused.Length; index++)
                    {
                        TestAssert.Equal(baseline[index].Content, unused[index].Content);
                        ReadPaintedSourceLabel(unused[index], footer);
                    }
                }

                var pages = RenderAlignmentProbe(WithStaticSelectionDrawing(document, drawing, header, "default"), mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    var expected = expectedFooters[index];
                    var actual = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(expected.Count, actual.Count);
                    AssertSamePaintOrigin(expected[0], actual[0]);
                    AssertStaticSelectionPictureFrame(pages[index]);
                }
            }
        }
    }

    private static DocxDocument WithStaticSelectionDrawing(DocxDocument document, DocxFloatingDrawing drawing,
        bool header, string variant)
    {
        var settings = document.PageSettings;
        var drawings = (header ? settings.HeaderFloatingDrawingsByType : settings.FooterFloatingDrawingsByType)
            .ToDictionary(pair => pair.Key, pair => pair.Value);
        drawings[variant] = [drawing];
        return document with
        {
            PageSettings = header ? settings with { HeaderFloatingDrawingsByType = drawings } :
                settings with { FooterFloatingDrawingsByType = drawings }
        };
    }

    private static void AssertStaticSelectionPictureFrame(PdfPage page) =>
        TestAssert.True(DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Any(frame => frame.Block.StartsWith("/", StringComparison.Ordinal) &&
            Math.Abs(frame.Transform.M11 - 20d) < .001d && Math.Abs(frame.Transform.M22 - 20d) < .001d &&
            Math.Abs(frame.Transform.M12) < .001d && Math.Abs(frame.Transform.M21) < .001d &&
            Math.Abs(frame.Transform.M31 - 420d) < .001d && Math.Abs(frame.Transform.M32 - (page.Height - 44d)) < .001d),
            "The floating picture retains its authored page frame independently of static text placement.");

    private static DocxFloatingDrawing CreateStaticSelectionPicture(DocxDocument document) =>
        new("0", "0", "0", "0", "0", "0", "1", "0", "1", "1", "254000", "254000",
            "page", null, "5334000", "page", null, "304800", DocxFloatingWrapKind.None, null,
            "public-static-picture", document.Paragraphs.SelectMany(paragraph => paragraph.Images).Single(), 0, 0);
}
