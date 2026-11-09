using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
    public static void DocxMathAssemblyRepeatedPartsAdmitConstructionBeforeMaterializing()
    {
        var assembly = new OpenTypeMathGlyphAssembly(0, 0, [new(34, 0, 0, 1, true)]);
        foreach (bool refuse in new[] { false, true })
        {
            var budget = new TestMemoryBudget(refuse ? 8192 : 1024 * 1024);
            using (var storage = OoxMemory.Reserve(budget, 4096))
            {
                if (refuse)
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        DocxMathGlyphAssembly.TryArrange(assembly, 1000d, default, out _, out _, storage)).Kind);
                else
                {
                    TestAssert.True(DocxMathGlyphAssembly.TryArrange(assembly, 1000d, default, out var result, out double height, storage),
                        "Repeated extenders keep complete source order and exact height.");
                    TestAssert.Equal(1000, TestAssert.NotNull(result).Count);
                    TestAssert.Equal(1000d, height);
                    TestAssert.True(budget.Reserved >= 4096 + 1000 * 8,
                        "Repeated output capacity is additional to earlier equation storage.");
                }
            }
            TestAssert.Equal(0L, budget.Reserved);
        }
    }

    public static void DocxMathGlyphAssemblyAvoidsRedundantLargePartCopies()
    {
        var warmup = new OpenTypeMathGlyphAssembly(0, 0, [new(34, 0, 0, 1, false)]);
        TestAssert.True(DocxMathGlyphAssembly.TryArrange(warmup, 1d, default, out _, out _), "The bounded arranger is ready.");
        foreach (int mode in new[] { 0, 1, 2 })
        {
            bool growing = mode == 2;
            var part = new OpenTypeMathGlyphPart(34, growing ? (ushort)10 : (ushort)0,
                growing ? (ushort)10 : (ushort)0, growing ? (ushort)10 : (ushort)1, mode == 1);
            var source = new OpenTypeMathGlyphPart[ushort.MaxValue];
            Array.Fill(source, part);
            var assembly = new OpenTypeMathGlyphAssembly(0, 0, Array.AsReadOnly(source));
            long before = GC.GetAllocatedBytesForCurrentThread();
            bool success = DocxMathGlyphAssembly.TryArrange(assembly, growing ? 16d : 1d, default, out var placed, out double height);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            TestAssert.True(success && placed is not null, "Fixed, unrepeated-extender and growing-overlap sequences remain complete.");
            var positions = TestAssert.NotNull(placed);
            TestAssert.Equal(source.Length, positions.Count);
            TestAssert.True(Math.Abs(height - (growing ? 16d : source.Length)) < 0.000001d, "The full assembly retains its exact enclosing height.");
            double step = growing ? 6d / (source.Length - 1) : 1d;
            for (int index = 0; index < positions.Count; index++)
            {
                TestAssert.Equal(part, positions[index].Part);
                TestAssert.True(Math.Abs(positions[index].Offset - index * step) < 0.000001d,
                    "All source parts retain their order and equal overlap distribution.");
            }
            TestAssert.True(allocated < 2_000_000, "Result storage and required growth scratch do not need duplicate source lists: " + allocated);
        }
    }

    public static void MathFontRadicalAssembliesKeepOneLogicalSymbolAndAllPhysicalParts()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "PortableFonts", "DejaVuMathTeXGyre.ttf"));
        var font = OpenTypeFont.Load(File.ReadAllBytes(path));
        var mappings = new List<(ushort Glyph, int Unicode)> { (font.MapCodePoint('A'), 'A') };
        ushort canonical = font.MapCodePoint(0x221A);
        for (double height = 0d; font.TryGetMathVerticalVariant(canonical, height, default, out var variant); height = variant.Advance + 1d)
            mappings.Add((variant.Glyph, 0x221A));
        var missingParts = new PdfFontResource("EqMissingParts", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
        TestAssert.True(font.TryGetMathVerticalAssembly(canonical, default, out var assembly), "The pinned font supplies actual root assembly parts.");
        foreach (var part in TestAssert.NotNull(assembly).Parts) mappings.Add((part.Glyph, 0x221A));
        var resource = new PdfFontResource("EqAssembly", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
        var body = new DocxMathRow([new DocxMathRun(new("A", 60d, "008000", false, false, false, null, "Math"), null, null)
            { NormalTextValue = true }]);
        var radical = new DocxMathRadical(body, new([]), true)
            { ControlText = new("", 12d, "FF0000", false, false, false, null, "Math") };
        var equation = new DocxEquation(new([radical]), 0, 0, false, null, null, null) { ParentFontSize = 11d };
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var measured, out var reason,
            resolveControlFont: _ => resource), reason ?? "The font-defined root assembly must enclose its larger literal child.");
        var box = TestAssert.NotNull(measured);
        TestAssert.Equal(2, box.Glyphs.Count);
        var root = box.Glyphs[0];
        TestAssert.True(root.SourceCodePoint == 0x221A && root.AssemblyParts.Count >= 3 &&
            ReferenceEquals(root.ControlOwner, radical) && root.FontSize == 12d && root.Color?.Red == 255,
            "One logical root retains its authored structural control and its physical pieces.");
        foreach (var part in root.AssemblyParts)
        {
            TestAssert.True(font.TryGetGlyphBounds(part.OriginalGlyph, out var ink), "Every physical root part has readable bounds.");
            double scale = root.FontSize / font.UnitsPerEm;
            TestAssert.True(root.X + part.X + ink.XMax * scale <= box.Width + 1e-9 &&
                root.BaselineY + part.BaselineY + ink.YMax * scale <= box.AboveBaseline + 1e-9 &&
                root.BaselineY + part.BaselineY + ink.YMin * scale >= -box.BelowBaseline - 1e-9,
                "The measured root contains the complete assembled ink.");
        }
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(box, graphics, 10d, 10d, default);
        var page = new PdfPage(100d, 100d, graphics.ToString(), [resource]);
        PdfContentValidator.ValidatePage(page, 0, default);
        TestAssert.Equal("\u221AA", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text)));
        TestAssert.Equal(root.AssemblyParts.Count + 1,
            System.Text.RegularExpressions.Regex.Matches(page.Content, @"<[^>]+> Tj").Count);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), new OoxConversionLimits(), diagnosticSink: null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
        TestAssert.True(output.Length > 0, "The one-symbol assembly writes a complete PDF with mapped physical glyphs.");
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => missingParts, default, out var absent, out _,
            resolveControlFont: _ => missingParts) && absent is null, "A missing late assembly CID cannot publish a partial root or child.");
    }

    public static void MathFontVerticalAssemblyPlacementKeepsConnectorBoundsAndBalancedGrowth()
    {
        var assembly = new OpenTypeMathGlyphAssembly(0, 40,
            new OpenTypeMathGlyphPart[] { new(34, 0, 300, 500, false), new(35, 200, 200, 200, true), new(36, 250, 0, 400, false) });
        foreach ((double target, double expected, int pieces) in new[] { (100d, 650d, 2), (700d, 700d, 2),
            (860d, 860d, 2), (900d, 900d, 3), (1500d, 1500d, 6) })
        {
            TestAssert.True(DocxMathGlyphAssembly.TryArrange(assembly, target, default, out var arranged, out double height),
                "A bounded font-defined assembly must reach its target.");
            var positions = TestAssert.NotNull(arranged);
            TestAssert.Equal(expected, height);
            TestAssert.Equal(pieces, positions.Count);
            TestAssert.Equal((ushort)34, positions[0].Part.Glyph);
            TestAssert.Equal((ushort)36, positions[^1].Part.Glyph);
            for (int index = 1; index < positions.Count; index++)
            {
                var previous = positions[index - 1]; var current = positions[index];
                double overlap = previous.Offset + previous.Part.FullAdvance - current.Offset;
                TestAssert.True(overlap >= 40d && overlap <= Math.Min(previous.Part.EndConnectorLength, current.Part.StartConnectorLength),
                    "Every connector stays within the font's available material and minimum overlap.");
            }
            if (target == 900d) TestAssert.Equal(400d, positions[1].Offset);
        }
        var zeroGrowth = new OpenTypeMathGlyphAssembly(0, 40,
            new OpenTypeMathGlyphPart[] { new(34, 0, 100, 500, false), new(35, 40, 40, 40, true),
                new(37, 200, 200, 200, true), new(36, 250, 0, 400, false) });
        TestAssert.True(DocxMathGlyphAssembly.TryArrange(zeroGrowth, 2000d, default, out var result, out double resultHeight) &&
            Math.Abs(resultHeight - 2000d) < 1e-9 && result!.Count(part => part.Part.Glyph == 35) == 1,
            "A zero-growth extender is retained once while useful extenders repeat.");
    }

    public static void MathFontVerticalAssemblyPlacementKeepsFallbackCancellationAndExistingWorkBounds()
    {
        var assembly = new OpenTypeMathGlyphAssembly(0, 40,
            new OpenTypeMathGlyphPart[] { new(34, 0, 100, 500, false), new(35, 200, 200, 200, true), new(36, 250, 0, 400, false) });
        foreach (double target in new[] { -1d, double.NaN, double.PositiveInfinity, 65536d })
            TestAssert.True(!DocxMathGlyphAssembly.TryArrange(assembly, target, default, out var absent, out _) && absent is null,
                "Invalid design-unit targets cannot publish placements.");
        var disconnected = assembly with { Parts = new OpenTypeMathGlyphPart[] { new(34, 0, 0, 500, false), new(36, 0, 0, 400, false) } };
        TestAssert.True(!DocxMathGlyphAssembly.TryArrange(disconnected, 1000d, default, out _, out _),
            "Parts without enough connector material retain explicit fallback.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathGlyphAssembly.TryArrange(assembly, 1000d, cancellation.Token, out _, out _));
    }
}
