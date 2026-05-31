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

    public static void WritesGrayColorOperatorsForEqualRgbChannels()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetFillRgb(128, 128, 128);
        graphics.FillRectangle(10, 20, 30, 40);
        graphics.SetStrokeRgb(0, 0, 0);
        graphics.StrokeRectangle(10, 20, 30, 40);
        graphics.DrawGlyphText("F1", 12, 20, 30, 255, 255, 255, "0041", textRenderingMode: 1, strokeRed: 128, strokeGreen: 128, strokeBlue: 128, strokeWidth: 0.5d);

        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString()) });

        TestAssert.Contains("0.502 g", pdf);
        TestAssert.Contains("0 G", pdf);
        TestAssert.Contains("1 g", pdf);
        TestAssert.Contains("0.502 G", pdf);
        TestAssert.DoesNotContain("0.502 0.502 0.502 rg", pdf);
        TestAssert.DoesNotContain("0 0 0 RG", pdf);
        TestAssert.DoesNotContain("1 1 1 rg", pdf);
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

    public static void WritesTilingPatternResources()
    {
        var graphics = new PdfGraphicsBuilder();
        var pattern = PdfTilingPattern.DiagonalLines(4d, up: true, 1d, 47, 133, 106);
        graphics.FillRectangleWithTilingPattern(10, 20, 30, 40, pattern);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings, graphics.Patterns);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/Pattern << /P1 ", pdf);
        TestAssert.Contains("/PatternType 1", pdf);
        TestAssert.Contains("/PaintType 1", pdf);
        TestAssert.Contains("/TilingType 1", pdf);
        TestAssert.Contains("/BBox [0 0 4 4]", pdf);
        TestAssert.Contains("/Pattern cs /P1 scn", pdf);
        TestAssert.Contains("10 20 30 40 re f", pdf);
    }

    public static void WritesScaledTilingPatternMatrix()
    {
        var graphics = new PdfGraphicsBuilder();
        var pattern = PdfTilingPattern.OfficeScaledDiagonalLines(up: true, 1d, 47, 133, 106);
        graphics.FillRectangleWithTilingPattern(10, 20, 30, 40, pattern);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings, graphics.Patterns);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/TilingType 2", pdf);
        TestAssert.Contains("/BBox [0 0 16 16]", pdf);
        TestAssert.Contains("/Matrix [0.375 0 0 0.375 0 0]", pdf);
        TestAssert.Contains("/XStep 16 /YStep 16", pdf);
    }

    public static void WritesImageBackedTilingPatternResources()
    {
        var graphics = new PdfGraphicsBuilder();
        var pattern = PdfTilingPattern.OfficeBitmapDiagonalLines(up: true, 47, 133, 106, 191, 191, 191);
        graphics.FillRectangleWithTilingPattern(10, 20, 30, 40, pattern);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings, graphics.Patterns);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/Pattern << /P1 ", pdf);
        TestAssert.Contains("/PatternType 1", pdf);
        TestAssert.Contains("/TilingType 2", pdf);
        TestAssert.Contains("/Resources << /XObject << /ImPattern ", pdf);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("q 16 0 0 16 0 0 cm /ImPattern Do Q", pdf);
        TestAssert.Contains("/Pattern cs /P1 scn", pdf);
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

    public static void WritesDistinctImageObjectsForDifferentSoftMasks()
    {
        byte[] rgb = [255, 0, 0, 0, 0, 255];
        PdfImageXObject first = PdfImageXObject.RgbPng(2, 1, rgb, [255, 64]);
        PdfImageXObject second = PdfImageXObject.RgbPng(2, 1, rgb, [64, 255]);
        var page = new PdfPage(100, 100, string.Empty, [], [
            new PdfImageResource("Im1", first),
            new PdfImageResource("Im2", second)
        ]);

        string pdf = WritePdfText([page]);

        TestAssert.Equal(4, CountOccurrences(pdf, "/Subtype /Image"));
        TestAssert.Equal(2, CountOccurrences(pdf, "/SMask"));
        TestAssert.Contains("/Im1 ", pdf);
        TestAssert.Contains("/Im2 ", pdf);
    }

    public static void WritesLuminositySoftMaskFormXObject()
    {
        PdfImageXObject image = PdfImageXObject.Jpeg(2, 1, [0xFF, 0xD8, 0xFF, 0xD9]);
        var mask = new PdfLuminositySoftMask(image, 10, 20, 30, 40, 0.1d, 0.2d, 0.3d, 0.4d);
        var graphics = new PdfGraphicsBuilder();
        graphics.SetLuminositySoftMask(mask, 0.5d, 1d);
        graphics.FillRectangle(10, 20, 30, 40);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/SMask << /S /Luminosity /G ", pdf);
        TestAssert.Contains("/Subtype /Form", pdf);
        TestAssert.Contains("/Group << /S /Transparency /CS /DeviceRGB >>", pdf);
        TestAssert.Contains("/BBox [10 20 40 60]", pdf);
        TestAssert.Contains("/ImMask ", pdf);
        TestAssert.Contains("/DCTDecode", pdf);
        TestAssert.Contains("/GSM1 gs", pdf);
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
