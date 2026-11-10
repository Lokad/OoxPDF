using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingUnusedStaticDrawingsPreserveSelectedStoryFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false);
            var baseline = RenderAlignmentProbe(document, mode);
            var drawing = CreateStaticSelectionPicture(document);
            foreach (bool header in new[] { true, false })
            foreach (string variant in new[] { "first", "even" })
            {
                var settings = document.PageSettings with { TitlePage = false, EvenAndOddHeaders = false };
                var drawings = (header ? settings.HeaderFloatingDrawingsByType : settings.FooterFloatingDrawingsByType)
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
                drawings[variant] = [drawing];
                settings = header ? settings with { HeaderFloatingDrawingsByType = drawings } :
                    settings with { FooterFloatingDrawingsByType = drawings };
                var pages = RenderAlignmentProbe(document with { PageSettings = settings }, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    TestAssert.Equal(baseline[index].Content, pages[index].Content);
                    ReadPaintedSourceLabel(pages[index], "Public synthetic composition portfolio");
                }
            }
        }
    }

    public static void PortableRenderingHeaderDrawingKeepsCounterpartFooterVisible()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false);
            var baseline = RenderAlignmentProbe(document, mode);
            var drawings = document.PageSettings.HeaderFloatingDrawingsByType.ToDictionary(pair => pair.Key, pair => pair.Value);
            drawings["default"] = [CreateStaticSelectionPicture(document)];
            var source = document with { PageSettings = document.PageSettings with { HeaderFloatingDrawingsByType = drawings } };
            var pages = RenderAlignmentProbe(source, mode);
            TestAssert.Equal(baseline.Length, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                const string footer = "Public synthetic composition portfolio";
                var expected = ReadPaintedSourceLabel(baseline[index], footer);
                var actual = ReadPaintedSourceLabel(pages[index], footer);
                TestAssert.Equal(expected.Count, actual.Count);
                AssertSamePaintOrigin(expected[0], actual[0]);
            }
        }
    }

    private static DocxFloatingDrawing CreateStaticSelectionPicture(DocxDocument document) =>
        new("0", "0", "0", "0", "0", "0", "1", "0", "1", "1", "254000", "254000",
            "page", null, "5334000", "page", null, "304800", DocxFloatingWrapKind.None, null,
            "public-static-picture", document.Paragraphs.SelectMany(paragraph => paragraph.Images).Single(), 0, 0);
}
