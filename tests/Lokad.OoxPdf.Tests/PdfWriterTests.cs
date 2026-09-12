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
        graphics.DrawGlyphText("F1", 12, 20, 30, 255, 255, 255, "0041", textRenderingMode: 1, strokeRed: 128, strokeGreen: 128, strokeBlue: 128, strokeWidth: 0.5d, italic: false, characterSpacing: 0d);

        PdfEmbeddedFont grayFont = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), new[] { 65 }, CancellationToken.None);
        string pdf = WritePdfText(new[] { new PdfPage(200, 200, graphics.ToString(), [new PdfFontResource("F1", grayFont)]) });

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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c), CancellationToken.None);
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

    public static void WritesLinkAnnotations()
    {
        var page = new PdfPage(
            200,
            200,
            string.Empty,
            [],
            [],
            [],
            [],
            [],
            [new PdfLinkAnnotation(10, 20, 30, 40, "https://example.invalid/a?b=(c)\\d", Destination: null)]);

        string pdf = WritePdfText([page]);

        TestAssert.Contains("/Annots [5 0 R]", pdf);
        TestAssert.Contains("/Type /Annot /Subtype /Link", pdf);
        TestAssert.Contains("/Rect [10 20 40 60]", pdf);
        TestAssert.Contains("/Border [0 0 0]", pdf);
        TestAssert.Contains(@"/A << /S /URI /URI (https://example.invalid/a?b=\(c\)\\d) >>", pdf);
    }

    public static void WritesInternalLinkDestinations()
    {
        var firstPage = new PdfPage(
            200,
            200,
            string.Empty,
            [],
            [],
            [],
            [],
            [],
            [PdfLinkAnnotation.ToDestination(10, 20, 30, 40, new PdfLinkDestination(1, null, 180, null))]);
        var secondPage = new PdfPage(200, 200);

        string pdf = WritePdfText([firstPage, secondPage]);

        TestAssert.Contains("/Annots [7 0 R]", pdf);
        TestAssert.Contains("/Type /Annot /Subtype /Link", pdf);
        TestAssert.Contains("/Rect [10 20 40 60]", pdf);
        TestAssert.Contains("/Dest [5 0 R /XYZ null 180 null]", pdf);
        TestAssert.DoesNotContain("/S /URI", pdf);
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
        PdfImageXObject image = PdfImageXObject.Jpeg(2, 1, [0xFF, 0xD8, 0xFF, 0xD9], 3, 8);
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

    public static void SharedSpaceGlyphExtractsAsPlainSpace()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        TestAssert.Equal(font.MapCodePoint(0x20), font.MapCodePoint(0xA0));

        // Both insertion orders resolve the shared space glyph to plain
        // space: Office extracts spaces where sources mix NBSPs (see P03).
        int[][] orders = [ [0x20, 0xA0], [0xA0, 0x20] ];
        foreach (int[] codePoints in orders)
        {
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, codePoints, CancellationToken.None);
            TestAssert.True(embedded.TryGetEncodedCid(font.MapCodePoint(0x20), out ushort cid), "Shared space glyph must be encoded.");
            TestAssert.Equal(0x20, embedded.UnicodeByCid[cid]);
        }
    }

    public static void WritesSubsetBaseFontWithSixLetterTagAndPlus()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c), CancellationToken.None);
        var page = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", embedded)]);

        string pdf = WritePdfText(new[] { page });

        TestAssert.True(embedded.BaseFontName.Length >= 8, "Expected subset BaseFont name to carry a tag prefix.");
        TestAssert.True(embedded.BaseFontName[6] == '+', "Expected six-letter tag followed by plus, got: " + embedded.BaseFontName);
        for (int i = 0; i < 6; i++)
        {
            TestAssert.True(embedded.BaseFontName[i] >= 'A' && embedded.BaseFontName[i] <= 'Z', "Expected uppercase tag letters, got: " + embedded.BaseFontName);
        }

        TestAssert.Contains("/BaseFont /" + embedded.BaseFontName, pdf);
        TestAssert.Contains("/FontName /" + embedded.BaseFontName, pdf);
    }

    public static void WritesDistinctSubsetBaseFontsForDifferentCodepointSets()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, "ABC".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(font, "XYZ".Select(c => (int)c), CancellationToken.None);
        var firstPage = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", first)]);
        var secondPage = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", second)]);

        string pdf = WritePdfText(new[] { firstPage, secondPage });

        TestAssert.True(first.BaseFontName != second.BaseFontName, "Differing subsets must have distinct BaseFont names.");
        TestAssert.Contains("/BaseFont /" + first.BaseFontName, pdf);
        TestAssert.Contains("/BaseFont /" + second.BaseFontName, pdf);
    }

    public static void WritesFontFile2WithDecodedLength()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, "Az".Select(c => (int)c), CancellationToken.None);
        var page = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", embedded)]);

        string pdf = WritePdfText(new[] { page });

        TestAssert.Contains("/FontFile2", pdf);
        TestAssert.Contains("/Length1 " + embedded.FontProgramBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), pdf);
    }

    public static void WritesCMapBatchesWithinOneHundredMappings()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        OpenTypeFont font = OpenTypeFont.Load(File.ReadAllBytes(arial));
        int[] codePoints = Enumerable.Range(0x20, 300).ToArray();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, codePoints, CancellationToken.None);
        var page = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", embedded)]);

        string pdf = WritePdfText(new[] { page });
        string cmap = embedded.BuildToUnicodeCMap(CancellationToken.None);
        TestAssert.True(embedded.UnicodeByCid.Count > 100, "Expected large probe to exceed one CMap block.");
        foreach (string line in cmap.Split('\n'))
        {
            string trimmed = line.Trim();
            if (trimmed.EndsWith("beginbfchar", StringComparison.Ordinal))
            {
                string countText = trimmed.Substring(0, trimmed.Length - "beginbfchar".Length).Trim();
                int count = int.Parse(countText, System.Globalization.CultureInfo.InvariantCulture);
                TestAssert.True(count >= 1 && count <= 100, "Each bfchar block must hold 1..100 mappings.");
            }
        }

        TestAssert.Contains("beginbfchar", pdf);
    }
    public static void RejectsNonFinitePdfNumbers()
    {
        TestAssert.Throws<ArgumentOutOfRangeException>(() => PdfDocumentWriter.FormatNumber(double.NaN));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => PdfDocumentWriter.FormatNumber(double.PositiveInfinity));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => PdfDocumentWriter.FormatNumber(double.NegativeInfinity));
        TestAssert.Equal("1.5", PdfDocumentWriter.FormatNumber(1.5));
    }

    public static void RejectsInvalidPageDimensions()
    {
        TestAssert.Throws<ArgumentOutOfRangeException>(() => WritePdfBytes(new[] { new PdfPage(double.NaN, 100) }));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => WritePdfBytes(new[] { new PdfPage(100, double.PositiveInfinity) }));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => WritePdfBytes(new[] { new PdfPage(0, 100) }));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => WritePdfBytes(new[] { new PdfPage(100, -5) }));
    }

    public static void PercentEncodesNonAsciiLinkUris()
    {
        var page = new PdfPage(200, 200, string.Empty, [], [], [], [], [], [PdfLinkAnnotation.ToUri(10, 20, 30, 40, "https://example.invalid/caf\u00e9?q=a b")]);
        string pdf = WritePdfText([page]);

        TestAssert.Contains("/URI (https://example.invalid/caf%C3%A9?q=a%20b)", pdf);
    }

    public static void OmitsDocumentInfoWithoutCreationDate()
    {
        string pdf = WritePdfText([new PdfPage(200, 200)]);

        TestAssert.DoesNotContain("/Info", pdf);
    }

    public static void RecordsFixedCreationDateInDocumentInfo()
    {
        var date = new DateTimeOffset(2026, 9, 9, 12, 30, 0, TimeSpan.FromHours(2));
        string first = WritePdfText([new PdfPage(200, 200)], date);
        string second = WritePdfText([new PdfPage(200, 200)], date);

        TestAssert.Contains("/Info", first);
        TestAssert.Contains("/CreationDate (D:20260909123000+02\u002700\u0027)", first);
        TestAssert.Contains("/Producer (Lokad.OoxPdf", first);
        TestAssert.Equal(first, second);
    }

    public static void DistinctCreationDatesProduceDistinctOutput()
    {
        var firstDate = new DateTimeOffset(2026, 9, 9, 12, 30, 0, TimeSpan.Zero);
        var secondDate = new DateTimeOffset(2026, 9, 10, 12, 30, 0, TimeSpan.Zero);
        string first = WritePdfText([new PdfPage(200, 200)], firstDate);
        string second = WritePdfText([new PdfPage(200, 200)], secondDate);

        TestAssert.True(!first.Equals(second, StringComparison.Ordinal), "Different creation dates must produce different bytes.");
    }

    public static void RejectsDuplicateFontResourceNames()
    {
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), new[] { 65 }, CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), new[] { 66 }, CancellationToken.None);
        var page = new PdfPage(200, 200, string.Empty, [new PdfFontResource("F1", first), new PdfFontResource("F1", second)]);
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsDanglingFontReference()
    {
        var page = new PdfPage(200, 200, "BT /F9 12 Tf <0041> Tj ET\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsDanglingImageReference()
    {
        var page = new PdfPage(200, 200, "q 10 0 0 10 0 0 cm /Im9 Do Q\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsDanglingGraphicsStateReference()
    {
        var page = new PdfPage(200, 200, "/GS9 gs\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsDanglingShadingReference()
    {
        var page = new PdfPage(200, 200, "/Sh9 sh\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsDanglingPatternReference()
    {
        var page = new PdfPage(200, 200, "/Pattern cs /P9 scn\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsUnbalancedGraphicsState()
    {
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([new PdfPage(200, 200, "q q Q\n")]));
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([new PdfPage(200, 200, "Q\n")]));
    }

    public static void RejectsUnbalancedTextState()
    {
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([new PdfPage(200, 200, "BT <0041> Tj\n")]));
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([new PdfPage(200, 200, "ET\n")]));
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([new PdfPage(200, 200, "BT BT <0041> Tj ET ET\n")]));
    }

    public static void AcceptsBalancedContentWithAllResourceKinds()
    {
        PdfEmbeddedFont font = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), new[] { 65 }, CancellationToken.None);
        PdfImageXObject image = PdfImageXObject.RgbPng(1, 1, [0, 0, 0], null);
        var graphics = new PdfGraphicsBuilder();
        graphics.SetAlpha(0.5d, 1d);
        graphics.DrawGlyphText("F1", 12, 20, 30, 0, 0, 0, "0041", false, 0d, 0, 0, 0, 0, 0d);
        graphics.DrawImage("Im1", 0, 0, 10, 10);
        graphics.PaintAxialShading(0, 0, 10, 0, 255, 0, 0, 0, 0, 255);
        var pattern = PdfTilingPattern.DiagonalLines(4d, up: true, 1d, 0, 0, 0);
        graphics.FillRectangleWithTilingPattern(0, 0, 10, 10, pattern);
        var page = new PdfPage(
            200,
            200,
            graphics.ToString(),
            [new PdfFontResource("F1", font)],
            [new PdfImageResource("Im1", image)],
            graphics.ExtGStates,
            graphics.Shadings,
            graphics.Patterns);
        string pdf = WritePdfText([page]);
        TestAssert.Contains("<0041> Tj", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.Contains("/Pattern cs /P1 scn", pdf);
    }

    public static void RejectsDanglingPatternImageReference()
    {
        var pattern = new PdfTilingPattern(4, 4, "q /ImMissing Do Q\n");
        var page = new PdfPage(200, 200, "/Pattern cs /P1 scn\n", [], [], [], [], [new PdfTilingPatternResource("P1", pattern)]);
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsUnbalancedPatternGraphicsState()
    {
        var pattern = new PdfTilingPattern(4, 4, "q q Q\n");
        var page = new PdfPage(200, 200, "/Pattern cs /P1 scn\n", [], [], [], [], [new PdfTilingPatternResource("P1", pattern)]);
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void AcceptsPatternWithMatchingImages()
    {
        PdfImageXObject image = PdfImageXObject.RgbPng(1, 1, [0, 0, 0], null);
        var pattern = new PdfTilingPattern(4, 4, 4, 4, null, 2, "q 1 0 0 1 0 0 cm /Im1 Do Q\n", [new PdfImageResource("Im1", image)]);
        var page = new PdfPage(200, 200, "/Pattern cs /P1 scn\n", [], [], [], [], [new PdfTilingPatternResource("P1", pattern)]);
        string pdf = WritePdfText([page]);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void RejectsNonAsciiPageContent()
    {
        var page = new PdfPage(200, 200, "q Q\n% caf\u00e9\n");
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void RejectsNonAsciiPatternContent()
    {
        var pattern = new PdfTilingPattern(4, 4, "q Q\n% caf\u00e9\n");
        var page = new PdfPage(200, 200, "/Pattern cs /P1 scn\n", [], [], [], [], [new PdfTilingPatternResource("P1", pattern)]);
        TestAssert.Throws<InvalidDataException>(() => WritePdfBytes([page]));
    }

    public static void XrefTablePointsAtObjectHeaders()
    {
        var creationDate = new DateTimeOffset(2026, 5, 4, 12, 0, 0, TimeSpan.Zero);
        byte[] pdf = WritePdfBytes(RichAuditDocument(), creationDate);

        TestAssert.True(StartsWithBytes(pdf, 0, "%PDF-1.7\n"), "File header must declare PDF 1.7.");
        TestAssert.True(EndsWithBytes(pdf, "%%EOF\n"), "File must end with %%EOF.");

        int startxrefMarker = LastIndexOfBytes(pdf, "startxref\n");
        TestAssert.True(startxrefMarker >= 0, "File must contain a startxref marker.");
        int offsetLineStart = startxrefMarker + "startxref\n".Length;
        int offsetLineEnd = IndexOfByte(pdf, 10, offsetLineStart);
        long xrefOffset = ParseDecimalLong(pdf, offsetLineStart, offsetLineEnd);

        int position = checked((int)xrefOffset);
        TestAssert.Equal("xref", ReadAsciiLine(pdf, ref position));
        string subsection = ReadAsciiLine(pdf, ref position);
        TestAssert.True(subsection.StartsWith("0 ", StringComparison.Ordinal), "Cross-reference table must declare a single subsection starting at object 0.");
        int entryCount = int.Parse(subsection.Substring(2), System.Globalization.CultureInfo.InvariantCulture);
        TestAssert.True(entryCount >= 2, "Cross-reference table must describe at least the catalog and pages.");
        TestAssert.Equal("0000000000 65535 f ", ReadAsciiLine(pdf, ref position));
        var offsets = new long[entryCount - 1];
        for (int i = 0; i < offsets.Length; i++)
        {
            string entry = ReadAsciiLine(pdf, ref position);
            TestAssert.Equal(19, entry.Length);
            TestAssert.True(entry.EndsWith(" 00000 n ", StringComparison.Ordinal), "Cross-reference entries must be uncompressed single-generation entries.");
            offsets[i] = ParseDecimalLong(pdf, position - 20, position - 10);
        }

        for (int i = 0; i < offsets.Length; i++)
        {
            string header = (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj";
            for (int j = 0; j < header.Length; j++)
            {
                TestAssert.Equal((byte)header[j], pdf[offsets[i] + j]);
            }
        }

        TestAssert.Equal("trailer", ReadAsciiLine(pdf, ref position));
        string trailerDict = ReadAsciiLine(pdf, ref position);
        TestAssert.True(trailerDict.Contains("/Size " + entryCount.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal), "Trailer Size must agree with the cross-reference table.");
        TestAssert.True(trailerDict.Contains("/Root 1 0 R", StringComparison.Ordinal), "Trailer must reference the catalog.");
        TestAssert.True(trailerDict.Contains("/Info " + (entryCount - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R", StringComparison.Ordinal), "Trailer must reference the document information dictionary.");
        TestAssert.Equal("startxref", ReadAsciiLine(pdf, ref position));
        TestAssert.Equal(xrefOffset, ParseDecimalLong(pdf, position, IndexOfByte(pdf, 10, position)));
    }

    public static void StreamLengthsMatchPayloads()
    {
        byte[] pdf = WritePdfBytes(RichAuditDocument(), new DateTimeOffset(2026, 5, 4, 12, 0, 0, TimeSpan.Zero));
        int validatedStreams = 0;
        int searchFrom = 0;
        while (true)
        {
            int lengthAt = IndexOfBytes(pdf, "/Length ", searchFrom);
            if (lengthAt < 0)
            {
                break;
            }

            int valueStart = lengthAt + "/Length ".Length;
            int valueEnd = valueStart;
            while (valueEnd < pdf.Length && pdf[valueEnd] >= 48 && pdf[valueEnd] <= 57)
            {
                valueEnd++;
            }

            long length = ParseDecimalLong(pdf, valueStart, valueEnd);
            int streamMarker = IndexOfBytes(pdf, "stream\n", valueEnd);
            TestAssert.True(streamMarker >= 0, "Every Length declaration must be followed by a stream payload.");
            long payloadStart = (long)streamMarker + "stream\n".Length;
            long payloadEnd = payloadStart + length;
            TestAssert.True(payloadEnd <= pdf.Length, "Stream payload must fit inside the file.");
            int trailerAt = checked((int)payloadEnd);
            if (trailerAt < pdf.Length && pdf[trailerAt] == 13)
            {
                trailerAt++;
            }

            if (trailerAt < pdf.Length && pdf[trailerAt] == 10)
            {
                trailerAt++;
            }

            TestAssert.True(StartsWithBytes(pdf, trailerAt, "endstream"), "Stream payload must be followed by endstream.");
            validatedStreams++;
            searchFrom = trailerAt + "endstream".Length;
        }

        TestAssert.Equal(7, validatedStreams);
        // Two page contents (the blank page still gets an empty stream), font program, ToUnicode map, image, soft mask, tiling pattern.
    }

    private static IReadOnlyList<PdfPage> RichAuditDocument()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, TestFontBuilder.AllCodePoints(), CancellationToken.None);
        byte[] rgb = [255, 0, 0, 0, 0, 255];
        PdfImageXObject image = PdfImageXObject.RgbPng(2, 1, rgb, [255, 64]);
        var graphics = new PdfGraphicsBuilder();
        graphics.DrawGlyphText("F1", 12, 20, 30, 0, 0, 0, "0041", false, 0d, 0, 0, 0, 0, 0d);
        graphics.DrawImage("Im1", 0, 0, 10, 10);
        graphics.PaintAxialShading(0, 0, 10, 0, 255, 0, 0, 0, 0, 255);
        var pattern = PdfTilingPattern.DiagonalLines(4d, up: true, 1d, 0, 0, 0);
        graphics.FillRectangleWithTilingPattern(0, 0, 10, 10, pattern);
        var first = new PdfPage(
            200,
            200,
            graphics.ToString(),
            [new PdfFontResource("F1", embedded)],
            [new PdfImageResource("Im1", image)],
            graphics.ExtGStates,
            graphics.Shadings,
            graphics.Patterns,
            [PdfLinkAnnotation.ToUri(10, 20, 30, 40, "https://example.invalid/")]);
        return [first, new PdfPage(100, 100)];
    }

    private static bool StartsWithBytes(byte[] pdf, int position, string text)
    {
        if (position < 0 || position + text.Length > pdf.Length)
        {
            return false;
        }

        for (int i = 0; i < text.Length; i++)
        {
            if (pdf[position + i] != (byte)text[i])
            {
                return false;
            }
        }

        return true;
    }

    private static bool EndsWithBytes(byte[] pdf, string text)
    {
        return StartsWithBytes(pdf, pdf.Length - text.Length, text);
    }

    private static int IndexOfBytes(byte[] pdf, string text, int start)
    {
        for (int i = start; i + text.Length <= pdf.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < text.Length; j++)
            {
                if (pdf[i + j] != (byte)text[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
            {
                return i;
            }
        }

        return -1;
    }

    private static int LastIndexOfBytes(byte[] pdf, string text)
    {
        for (int i = pdf.Length - text.Length; i >= 0; i--)
        {
            bool match = true;
            for (int j = 0; j < text.Length; j++)
            {
                if (pdf[i + j] != (byte)text[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
            {
                return i;
            }
        }

        return -1;
    }

    private static int IndexOfByte(byte[] pdf, int value, int start)
    {
        for (int i = start; i < pdf.Length; i++)
        {
            if (pdf[i] == value)
            {
                return i;
            }
        }

        return -1;
    }

    private static long ParseDecimalLong(byte[] pdf, int start, int end)
    {
        long value = 0;
        for (int i = start; i < end; i++)
        {
            TestAssert.True(pdf[i] >= 48 && pdf[i] <= 57, "Expected ASCII digits while parsing a cross-reference value.");
            value = checked(value * 10 + (pdf[i] - 48));
        }

        return value;
    }

    private static string ReadAsciiLine(byte[] pdf, ref int position)
    {
        int end = IndexOfByte(pdf, 10, position);
        TestAssert.True(end >= 0, "Unexpected end of file while reading a cross-reference line.");
        var chars = new char[end - position];
        for (int i = 0; i < chars.Length; i++)
        {
            chars[i] = (char)pdf[position + i];
        }

        position = end + 1;
        return new string(chars);
    }

    private static string WritePdfText(IReadOnlyList<PdfPage> pages)
    {
        return WritePdfText(pages, null);
    }

    private static string WritePdfText(IReadOnlyList<PdfPage> pages, DateTimeOffset? creationDate)
    {
        return Encoding.ASCII.GetString(WritePdfBytes(pages, creationDate));
    }

    private static byte[] WritePdfBytes(IReadOnlyList<PdfPage> pages)
    {
        return WritePdfBytes(pages, null);
    }

    public static void TruncatedContentMarkRewindsPaintAndResources()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetFillRgb(255, 0, 0);
        graphics.FillRectangle(10, 20, 30, 40);
        int survivingLength = graphics.ToString().Length;
        PdfGraphicsBuilder.ContentMark mark = graphics.MarkContent();

        graphics.SetFillRgb(0, 0, 255);
        graphics.FillRectangle(50, 60, 70, 80);
        graphics.SaveState();
        graphics.SetAlpha(0.5d, 0.5d);
        graphics.PaintAxialShading(0, 0, 100, 100, 255, 0, 0, 0, 0, 255);
        var pattern = PdfTilingPattern.DiagonalLines(4d, up: true, 1d, 47, 133, 106);
        graphics.FillRectangleWithTilingPattern(1, 2, 3, 4, pattern);

        graphics.TruncateContent(mark);

        string content = graphics.ToString();
        TestAssert.Equal(survivingLength, content.Length);
        TestAssert.Contains("1 0 0 rg", content);
        TestAssert.Contains("10 20 30 40 re f", content);
        TestAssert.DoesNotContain("0 0 1 rg", content);
        TestAssert.DoesNotContain("50 60 70 80 re f", content);
        TestAssert.Equal(0, graphics.ExtGStates.Count);
        TestAssert.Equal(0, graphics.Shadings.Count);
        TestAssert.Equal(0, graphics.Patterns.Count);
        TestAssert.Equal(0, graphics.StateDepth);
    }

    public static void TruncatedContentMarksNestAcrossNodes()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetFillRgb(255, 0, 0);
        graphics.FillRectangle(10, 20, 30, 40);
        PdfGraphicsBuilder.ContentMark first = graphics.MarkContent();
        graphics.SetFillRgb(0, 1, 0);
        graphics.FillRectangle(50, 60, 70, 80);
        PdfGraphicsBuilder.ContentMark second = graphics.MarkContent();
        graphics.SetFillRgb(0, 0, 255);
        graphics.FillRectangle(90, 100, 110, 120);

        graphics.TruncateContent(second);
        string middle = graphics.ToString();
        TestAssert.Contains("10 20 30 40 re f", middle);
        TestAssert.Contains("50 60 70 80 re f", middle);
        TestAssert.DoesNotContain("90 100 110 120 re f", middle);

        graphics.TruncateContent(first);
        string rolledBack = graphics.ToString();
        TestAssert.Contains("10 20 30 40 re f", rolledBack);
        TestAssert.DoesNotContain("50 60 70 80 re f", rolledBack);
    }

    private static byte[] WritePdfBytes(IReadOnlyList<PdfPage> pages, DateTimeOffset? creationDate)
    {
        using var stream = new MemoryStream();
        PdfDocumentWriter.WriteBlank(stream, pages, CancellationToken.None, creationDate);
        return stream.ToArray();
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

    public static void DrawGlyphTextEmitsQuarterTurnTextMatrices()
    {
        TestAssert.Contains("1 0 0 1 100 200 Tm", DrawGlyphTextContent(0));
        TestAssert.Contains("0 -1 1 0 100 200 Tm", DrawGlyphTextContent(1));
        TestAssert.Contains("-1 0 0 -1 100 200 Tm", DrawGlyphTextContent(2));
        TestAssert.Contains("0 1 -1 0 100 200 Tm", DrawGlyphTextContent(3));
        TestAssert.Contains("1 0 0 1 100 200 Tm", DrawGlyphTextContent(4));
        TestAssert.Contains("0 -1 1 0 100 200 Tm", DrawGlyphTextContent(-3));

        static string DrawGlyphTextContent(int quarterTurns)
        {
            var graphics = new PdfGraphicsBuilder();
            graphics.DrawGlyphText("F1", 24d, 100d, 200d, 0, 0, 0, "0041", false, 0d, 0, 0, 0, 0, 0d, quarterTurns);
            System.Reflection.FieldInfo? field = typeof(PdfGraphicsBuilder).GetField("builder", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var content = (System.Text.StringBuilder?)field?.GetValue(graphics);
            return content?.ToString() ?? throw new System.InvalidOperationException("Expected builder content.");
        }
    }
}
