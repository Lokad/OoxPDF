using System.Buffers.Binary;
using System.Globalization;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathBoxTests
{
    public static void DocxMathBoxVariantSymbolsKeepTheirSourceIdentityAndDistinctStyleGlyphs()
    {
        int[] source = [0x2202, 0x03F5, 0x03D1, 0x03F0, 0x03D5, 0x03F1, 0x03D6];
        foreach ((string style, int start) in new[] { ("i", 0x1D715), ("b", 0x1D6DB), ("bi", 0x1D74F) })
        foreach ((int point, int offset) in source.Select((point, offset) => (point, offset)))
        {
            string text = new System.Text.Rune(point).ToString();
            var row = new DocxMathRow([new DocxMathRun(Run(text), style, null)]);
            TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out var points, out _),
                "A supported variant symbol must retain its specific mathematical glyph.");
            TestAssert.Equal(start + offset, points.Single());
            TestAssert.Equal(text, ((DocxMathRun)row.Children[0]).Text.Text);
        }
        var theta = new DocxMathRow([new DocxMathRun(Run("θϑ"), "bi", null)]);
        TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(theta, CancellationToken.None, out var distinct, out _),
            "Ordinary and variant theta must remain discoverable together.");
        TestAssert.True(distinct.Count == 2 && distinct[0] != distinct[1],
            "Variant symbols cannot collapse to visually similar ordinary Greek letters.");
    }

    public static void DocxMathBoxGreekStylesPreserveEveryBasicLetterThroughUnicodeNormalization()
    {
        foreach (string style in new[] { "b", "bi" })
        foreach (int point in Enumerable.Range(0x0391, 0x03A9 - 0x0391 + 1)
            .Where(point => point != 0x03A2).Concat(Enumerable.Range(0x03B1, 0x03C9 - 0x03B1 + 1)))
        {
            string text = new System.Text.Rune(point).ToString();
            var row = new DocxMathRow([new DocxMathRun(Run(text), style, null)]);
            TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out var points, out _),
                "A basic Greek letter must resolve its explicit mathematical weight/style.");
            int rendered = points.Single();
            TestAssert.True(rendered is >= 0x1D6A8 and <= 0x1D74E && rendered != point,
                "Explicit Greek styles use the Unicode mathematical alphabet, including the capital-sigma gap.");
            TestAssert.Equal(text, new System.Text.Rune(rendered).ToString().Normalize(System.Text.NormalizationForm.FormKC));
            TestAssert.Equal(text, ((DocxMathRun)row.Children[0]).Text.Text);
        }
    }

    public static void DocxMathBoxExplicitStylesPreserveSourceAndSelectTheMathematicalAlphabet()
    {
        foreach ((string style, int[] expected) in new[]
        {
            ("p", new[] { 0x41, 0x68, 0x2B, 0x3D, 0x78, 0x32 }),
            ("i", new[] { 0x1D434, 0x210E, 0x2B, 0x3D, 0x1D465, 0x32 }),
            ("b", new[] { 0x1D400, 0x1D421, 0x2B, 0x3D, 0x1D431, 0x1D7D0 }),
            ("bi", new[] { 0x1D468, 0x1D489, 0x2B, 0x3D, 0x1D499, 0x1D7D0 })
        })
        {
            var row = new DocxMathRow([new DocxMathRun(Run("Ah+h=x2"), style, "roman")]);
            TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out var points, out _),
                "An explicit supported style must discover its mathematical alphabet.");
            TestAssert.True(points.SequenceEqual(expected), "Styled letters/digits own their Unicode glyphs; operators remain ordinary.");
            TestAssert.Equal("Ah+h=x2", ((DocxMathRun)row.Children[0]).Text.Text);
        }
    }

    public static void DocxMathBoxExplicitStylesRetainUnqualifiedAlphabetAndWordPropertyFallbacks()
    {
        foreach (var run in new[]
        {
            new DocxMathRun(Run("Ж"), "b", null),
            new DocxMathRun(Run("A"), "future", null),
            new DocxMathRun(Run("A"), "b", "double-struck"),
            new DocxMathRun(Run("A") with { Bold = true }, "b", null),
            new DocxMathRun(Run("A") with { Underline = true }, "bi", null)
        })
        {
            TestAssert.True(!DocxMathRowLayout.TryGetRenderCodePoints(new([run]), CancellationToken.None, out var points, out _) &&
                points.Count == 0, "Unqualified source styling must retain the complete fallback without partial glyph discovery.");
        }
    }

    public static void DocxMathBoxOperatorSpacingOwnerPreservesGlyphAdvanceAndRejectsInvalidSizes()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A+B");
        TestAssert.True(DocxMathRowLayout.TryMeasure(Plain("A+B"), _ => font, CancellationToken.None, out var normal, out _),
            "Expected ordinary operator spacing.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(Plain("A+B"), _ => font, CancellationToken.None, out var parent, out _,
            operatorSpacingFontSize: 6d), "Expected independently owned parent operator spacing.");
        DocxMathBox original = TestAssert.NotNull(normal), owned = TestAssert.NotNull(parent);
        TestAssert.True(owned.Width < original.Width && owned.Glyphs.Zip(original.Glyphs).All(pair =>
            pair.First.FontSize == pair.Second.FontSize && pair.First.Advance == pair.Second.Advance &&
            pair.First.EncodedGlyph == pair.Second.EncodedGlyph),
            "Changing the math-unit owner must preserve physical glyph sizing, advance and encoding.");
        foreach (double invalid in new[] { 0d, -1d, double.NaN, double.PositiveInfinity })
        {
            TestAssert.True(!DocxMathRowLayout.TryMeasure(Plain("A+B"), _ => font, CancellationToken.None,
                out var box, out _, operatorSpacingFontSize: invalid) && box is null,
                "Invalid operator spacing ownership cannot publish partial geometry.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathRowLayout.TryMeasure(Plain("A+B"), _ => font,
            cancelled.Token, out _, out _, operatorSpacingFontSize: double.NaN));
    }

    public static void DocxMathBoxPainterUsesTheMeasuredGlyphsPositionsAndResources()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A+B=C");
        DocxMathRow row = Plain("A+B=C");
        TestAssert.True(DocxMathRowLayout.TryMeasure(row, _ => font, CancellationToken.None, out DocxMathBox? box, out string? fallback), fallback ?? "Expected measured math row.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(5, measured.Glyphs.Count);
        TestAssert.True(measured.AboveBaseline > 0d && measured.Width > 0d, "The box must retain measured extents.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        string content = graphics.ToString();
        TestAssert.Equal(5, System.Text.RegularExpressions.Regex.Matches(content, " Tj").Count);
        foreach (DocxMathGlyphPlacement glyph in measured.Glyphs)
        {
            TestAssert.Contains("<" + glyph.EncodedGlyph + "> Tj", content);
            string x = (20d + glyph.X).ToString("0.###", CultureInfo.InvariantCulture);
            TestAssert.Contains(" " + x + " 50 Tm", content);
        }
        TestAssert.True(graphics.UsedFontResourceNames.SetEquals(new[] { "Math1" }), "Painting must use the measured font resource.");
        TestAssert.Equal(0, graphics.StateDepth);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 100d, content, [font])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Measured native math must produce a validated PDF.");
    }

    public static void DocxMathBoxUnarySignsAndRelationsUseSemanticSpacing()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A+B=-");
        DocxMathBox Measure(string text)
        {
            TestAssert.True(DocxMathRowLayout.TryMeasure(Plain(text), _ => font, CancellationToken.None, out DocxMathBox? box, out _), "Expected a literal row.");
            return TestAssert.NotNull(box);
        }
        DocxMathBox unary = Measure("-A"), binary = Measure("A-A"), relation = Measure("A=A"), afterRelation = Measure("A=-A");
        double hyphenAdvance = unary.Glyphs[0].Advance;
        TestAssert.True(Math.Abs(unary.Glyphs[1].X - hyphenAdvance) < 1e-9, "An initial unary sign has no binary-operator glue.");
        double binaryGap = binary.Glyphs[1].X - binary.Glyphs[0].Advance;
        double relationGap = relation.Glyphs[1].X - relation.Glyphs[0].Advance;
        TestAssert.True(binaryGap > 0d && relationGap > binaryGap, "Relation spacing must exceed binary spacing.");
        TestAssert.True(Math.Abs(afterRelation.Glyphs[3].X - afterRelation.Glyphs[2].X - afterRelation.Glyphs[2].Advance) < 1e-9, "A unary sign after a relation must not gain binary spacing.");
    }

    public static void DocxMathBoxRetainsSourceMeaningWhenSelectingItalicAlphabetGlyphs()
    {
        DocxMathRow row = new([new DocxMathRun(Run("Aahαβγ"), null, null)]);
        TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out IReadOnlyList<int> points, out _), "Expected supported mathematical alphabet.");
        TestAssert.Equal("1D434,1D44E,210E,1D6FC,1D6FD,1D6FE", string.Join(',', points.Select(point => point.ToString("X", CultureInfo.InvariantCulture))));
        TestAssert.Equal("Aahαβγ", ((DocxMathRun)row.Children.Single()).Text.Text);
        TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(Plain("Aahαβγ"), CancellationToken.None, out points, out _), "Plain math style preserves the declared alphabet.");
        TestAssert.Equal("41,61,68,3B1,3B2,3B3", string.Join(',', points.Select(point => point.ToString("X", CultureInfo.InvariantCulture))));
        var unsupported = DocxMathParser.Read(System.Xml.Linq.XElement.Parse("""
            <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:nor/></m:rPr><m:t>A</m:t></m:r></m:oMath>
            """), (_, text) => Run(text), CancellationToken.None);
        TestAssert.True(((DocxMathRun)unsupported.Children.Single()).HasUnsupportedProperties, "Unsupported source properties must survive parser admission.");
        TestAssert.True(!DocxMathRowLayout.TryGetRenderCodePoints(unsupported, CancellationToken.None, out _, out _), "Unsupported source properties cannot silently become ordinary math.");
    }

    public static void DocxMathBoxFallbackDoesNotPublishPartialOrUnencodedMathematics()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A");
        void Reject(DocxMathRow row, Func<DocxMathRun, PdfFontResource?> resolver, string reason)
        {
            TestAssert.True(!DocxMathRowLayout.TryMeasure(row, resolver, CancellationToken.None, out DocxMathBox? box, out string? fallback), "Unsupported mathematics must retain fallback.");
            TestAssert.True(box is null, "A failed box cannot publish partially painted mathematical meaning.");
            TestAssert.Equal(reason, fallback);
        }
        Reject(Plain("AB"), _ => font, "MissingMathEncoding");
        Reject(new([new DocxMathRun(Run("A"), null, null)]), _ => font, "MissingMathGlyph");
        Reject(Plain("A"), _ => null, "MissingMathFont");
        Reject(new([new DocxMathFraction(Plain("A"), Plain("A"), null)]), _ => font, "UnsupportedMathStructureOrStyle");
        Reject(new([new DocxMathRun(Run("A"), "p", null) { HasUnsupportedProperties = true }]), _ => font, "UnsupportedMathStructureOrStyle");
        Reject(new([new DocxMathRun(Run("A") with { FontSize = double.NaN }, "p", null)]), _ => font, "InvalidMathFontSize");
        Reject(new([new DocxMathRun(Run("AA") with { FontSize = double.MaxValue }, "p", null)]), _ => font, "InvalidMathExtent");
    }

    public static void DocxMathBoxMeasurementAndPaintingObserveCancellation()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathRowLayout.TryMeasure(Plain("A"), _ => font, cancelled.Token, out _, out _));
        using var duringResolution = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathRowLayout.TryMeasure(Plain("AA"), _ =>
        { duringResolution.Cancel(); return font; }, duringResolution.Token, out _, out _));
        TestAssert.True(DocxMathRowLayout.TryMeasure(Plain("A"), _ => font, CancellationToken.None, out DocxMathBox? box, out _), "Expected a measured box.");
        TestAssert.Throws<OperationCanceledException>(() => DocxMathBoxPainter.Paint(TestAssert.NotNull(box), new PdfGraphicsBuilder(), 0d, 0d, cancelled.Token));
    }

    public static void DocxMathBoxExpansionUsesTheExistingCumulativeSceneBudget()
    {
        var font = Resource(TestFontBuilder.LoadTestFont(), "A");
        using OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 2 });
        TestAssert.True(DocxMathRowLayout.TryGetRenderCodePoints(Plain("A"), CancellationToken.None, out _, out _), "First expansion must fit.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(Plain("A"), _ => font, CancellationToken.None, out _, out _), "Second expansion must fit.");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathRowLayout.TryMeasure(Plain("A"), _ => font, CancellationToken.None, out _, out _));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathRowLayout.TryGetRenderCodePoints(Plain(new string('A', 100_000)), CancellationToken.None, out _, out _));
    }

    public static void MathFontItalicCorrectionsReadBothCoverageFormatsAndSignedValues()
    {
        foreach (bool ranges in new[] { false, true })
        {
            byte[] table = ItalicTable(ranges);
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
            TestAssert.True(font.TryGetMathItalicCorrection(34, out short value), "Covered italic correction must be readable.");
            TestAssert.Equal((short)-71, value);
            TestAssert.True(font.TryGetMathItalicCorrection(35, out value), "Second covered glyph must use its own correction.");
            TestAssert.Equal((short)91, value);
            TestAssert.True(font.TryGetMathItalicCorrection(33, out value), "Uncovered glyphs have zero italic correction.");
            TestAssert.Equal((short)0, value);
            for (int length = 1; length < table.Length; length++)
            {
                OpenTypeFont truncated = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table[..length]));
                TestAssert.True(!truncated.TryGetMathItalicCorrection(34, out _), "Truncated italic metrics must not escape their table.");
            }
        }
    }

    public static void MathFontGlyphBoundsAgreeWithOutlinesWithoutExpandingThem()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        foreach (int codePoint in new[] { 'A', 'B', 'g', 0x1F600 })
        {
            ushort glyph = font.MapCodePoint(codePoint);
            TestAssert.True(font.TryGetGlyphBounds(glyph, out var bounds), "Expected glyph header bounds.");
            TestAssert.True(font.TryReadGlyphOutline(glyph, out var outline), "Expected a test outline.");
            TestAssert.Equal(outline.Bounds, bounds);
        }
        TestAssert.True(!font.TryGetGlyphBounds(ushort.MaxValue, out _), "Out-of-range glyphs must not expose bounds.");
    }

    private static DocxTextRun Run(string text) => new(text, 12d, null, false, false, false, null, "TestFont");
    private static DocxMathRow Plain(string text) => new([new DocxMathRun(Run(text), "p", null)]);
    private static PdfFontResource Resource(OpenTypeFont font, string text) => new("Math1", PdfEmbeddedFont.Create(font, text.EnumerateRunes().Select(rune => rune.Value), CancellationToken.None));
    private static byte[] ItalicTable(bool ranges)
    {
        byte[] bytes = new byte[ranges ? 40 : 38];
        void U(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(at), value);
        U(0, 1); U(6, 10); U(10, 8); U(18, 12); U(20, 2);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(22), -71);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(26), 91);
        U(30, ranges ? (ushort)2 : (ushort)1); U(32, ranges ? (ushort)1 : (ushort)2); U(34, 34);
        U(36, 35); if (ranges) { U(38, 0); }
        return bytes;
    }
}
