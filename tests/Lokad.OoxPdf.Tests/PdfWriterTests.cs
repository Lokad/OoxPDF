using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class PdfWriterTests
{
    public static void WritesSingleBlankPagePdfStructure()
    {
        string pdf = WritePdfText(new[] { new PdfPage(612, 792) });

        TestAssert.Contains("%PDF-1.7", pdf);
        TestAssert.Contains("<< /Type /Catalog /Pages 2 0 R >>", pdf);
        TestAssert.Contains("<< /Type /Pages /Count 1 /Kids [3 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 612 792]", pdf);
        TestAssert.Contains("xref", pdf);
        TestAssert.Contains("trailer", pdf);
        TestAssert.Contains("/Root 1 0 R", pdf);
        TestAssert.Contains("%%EOF", pdf);
    }

    public static void WritesMultipleBlankPagesWithPageSizes()
    {
        string pdf = WritePdfText(new[] { new PdfPage(960, 540), new PdfPage(595.276, 841.89) });

        TestAssert.Contains("<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 960 540]", pdf);
        TestAssert.Contains("/MediaBox [0 0 595.276 841.89]", pdf);
        TestAssert.Equal(2, CountOccurrences(pdf, "/Type /Page /Parent"));
    }

    public static void WritesDrawingOperators()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetFillRgb(255, 0, 0);
        graphics.FillRectangle(10, 20, 30, 40);
        graphics.SetStrokeRgb(0, 0, 255);
        graphics.SetLineWidth(2);
        graphics.StrokeLine(0, 0, 100, 100);
        graphics.FillEllipse(10, 10, 20, 20);

        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString()) });

        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("10 20 30 40 re f", pdf);
        TestAssert.Contains("0 0 1 RG", pdf);
        TestAssert.Contains("2 w", pdf);
        TestAssert.Contains("0 0 m 100 100 l S", pdf);
        TestAssert.Contains(" c", pdf);
    }

    public static void WritesEvenOddClippingOperators()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.ClipRectangleEvenOdd(10, 20, 30, 40);

        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString()) });

        TestAssert.Contains("10 20 30 40 re W* n", pdf);
    }

    public static void WritesOpenEvenOddClippingOperators()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.ClipOpenRectangleEvenOdd(10, 20, 30, 40);

        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString()) });

        TestAssert.Contains("10 60 m", pdf);
        TestAssert.Contains("40 60 l", pdf);
        TestAssert.Contains("40 20 l", pdf);
        TestAssert.Contains("10 20 l", pdf);
        TestAssert.Contains("W* n", pdf);
        TestAssert.DoesNotContain(" h", pdf);
        TestAssert.DoesNotContain(" re W* n", pdf);
    }

    public static void WritesEvenOddFillOperators()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.FillRectangleEvenOdd(10, 20, 30, 40);
        graphics.FillPolygonEvenOdd([(0d, 0d), (10d, 0d), (10d, 10d)]);

        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString()) });

        TestAssert.Contains("10 20 30 40 re f*", pdf);
        TestAssert.Contains("0 0 m", pdf);
        TestAssert.Contains("f*", pdf);
    }

    public static void WritesEmbeddedTrueTypeFontObjects()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        OpenTypeFont font = OpenTypeFont.Load(arial);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c));
        var page = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", embedded)]);

        string pdf = WritePdfText(new[] { page });

        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains("/Encoding /Identity-H", pdf);
        TestAssert.Contains("/Subtype /CIDFontType2", pdf);
        TestAssert.Contains("/CIDToGIDMap /Identity", pdf);
        TestAssert.Contains("/Type /FontDescriptor", pdf);
        TestAssert.Contains("/FontFile2", pdf);
        TestAssert.Contains("/ToUnicode", pdf);
        TestAssert.Contains("beginbfchar", pdf);
        TestAssert.Contains("<0041>", pdf);
        TestAssert.Contains("/F1 ", pdf);
    }

    public static void WritesAxialShadingResources()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.ClipRectangle(10, 20, 30, 40);
        graphics.PaintAxialShading(10, 20, 40, 60, 255, 0, 0, 0, 0, 255);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/Shading << /Sh1 ", pdf);
        TestAssert.Contains("/ShadingType 2", pdf);
        TestAssert.Contains("/Coords [10 20 40 60]", pdf);
        TestAssert.Contains("/C0 [1 0 0]", pdf);
        TestAssert.Contains("/C1 [0 0 1]", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
    }

    public static void WritesJpegImageColorSpaceFromFrameMetadata()
    {
        PdfImageXObject image = PdfImageXObject.Jpeg(1, 1, [0xFF, 0xD8, 0xFF, 0xD9], componentCount: 1, bitsPerComponent: 8);
        var page = new PdfPage(100, 100, string.Empty, [], [new PdfImageResource("Im1", image)]);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/ColorSpace /DeviceGray", pdf);
        TestAssert.Contains("/BitsPerComponent 8", pdf);
        TestAssert.Contains("/Filter /DCTDecode", pdf);
    }

    public static void WritesStitchedAxialShadingFunctionForMultipleStops()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.PaintAxialShading(0, 0, 100, 0, [
            new PdfShadingStop(0d, 255, 0, 0),
            new PdfShadingStop(0.5d, 0, 255, 0),
            new PdfShadingStop(1d, 0, 0, 255)
        ]);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/FunctionType 3", pdf);
        TestAssert.Contains("/Bounds [0.5]", pdf);
        TestAssert.Contains("/Encode [0 1 0 1]", pdf);
        TestAssert.Contains("/C0 [1 0 0]", pdf);
        TestAssert.Contains("/C1 [0 1 0]", pdf);
        TestAssert.Contains("/C1 [0 0 1]", pdf);
    }

    private static string WritePdfText(IReadOnlyList<PdfPage> pages)
    {
        using var stream = new MemoryStream();
        PdfDocumentWriter.WriteBlank(stream, pages);
        return Encoding.ASCII.GetString(stream.ToArray());
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int start = 0;
        while (true)
        {
            int index = text.IndexOf(value, start, StringComparison.Ordinal);
            if (index < 0)
            {
                return count;
            }

            count++;
            start = index + value.Length;
        }
    }
}
