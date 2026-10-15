using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathMatrixLayoutTests
{
    public static void MathFontMatrixLiteralCellsShareSlotsAxisAndMetricOnlyControl()
    {
        var metric = MathFont(); var plain = TestFontBuilder.LoadTestFont();
        var first = Resource("First", plain, "BC"); var rest = Resource("Rest", plain, "BC");
        var a = Run("B", "0000FF"); var b = Run("CC"); var c = Run("BBB", "008000"); var d = Run("C");
        foreach (double size in new[] { 12d, 20d })
        {
            var matrix = Matrix([a, b], [c, d]) with { ControlText = new("", size, "FF0000", false, false, false, null, "TestFont") };
            TestAssert.True(DocxMathLayout.TryGetRuns(Equation(matrix), CancellationToken.None, out var sources) && sources.Count == 4 &&
                ReferenceEquals(sources[0].ControlOwner, matrix) && sources.Skip(1).All(s => s.ControlOwner is null),
                "One invisible owner retains ordered cell source without taking over painted fonts.");
            TestAssert.True(DocxMathLayout.GetControlCodePoints(matrix).Count == 0, "A matrix control requests metrics without glyph mappings.");
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), r => ReferenceEquals(r, a) ? first : rest, CancellationToken.None,
                out var box, out var reason, resolveControlMetrics: _ => metric), reason ?? "Expected literal grid.");
            var measured = TestAssert.NotNull(box); double advance = plain.GetAdvanceWidth(plain.MapCodePoint('B')) * 12d / plain.UnitsPerEm;
            double otherAdvance = plain.GetAdvanceWidth(plain.MapCodePoint('C')) * 12d / plain.UnitsPerEm;
            TestAssert.Equal("BCCBBBC", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.True(Math.Abs(measured.Glyphs[0].X - advance) < 1e-9 && measured.Glyphs[3].X == 0d &&
                Math.Abs(measured.Glyphs[1].X - 3d * advance - size) < 1e-9 &&
                Math.Abs(measured.Glyphs[6].X - 3d * advance - otherAdvance / 2d - size) < 1e-9,
                "Each column centers its complete logical cell width and consumes one control-sized em between slots.");
            double export = OfficePdfTextEmissionProfile.FontSize(size);
            double line = DocxLineMetrics.MeasureTypographicLineHeight(metric, export), above = DocxLineMetrics.MeasureTypographicBaselineInset(metric, export);
            TestAssert.True(Math.Abs(measured.Glyphs[0].BaselineY - measured.Glyphs[3].BaselineY - line) < 1e-9,
                "Single row spacing reuses typographic frames instead of independent per-cell baseline corrections.");
            double frameTop = measured.Glyphs[0].BaselineY + above, frameBottom = measured.Glyphs[3].BaselineY - (line - above);
            TestAssert.True(Math.Abs((frameTop + frameBottom) / 2d - size / 4d) < 1e-9,
                "The whole grid frame centers on the metric font's mathematical axis.");
            TestAssert.True(measured.ParentMetrics is { IncludeTypographicFrame: false } owner && ReferenceEquals(owner.Font, metric) && owner.Size == size &&
                ReferenceEquals(measured.Glyphs[0].Font, first) && measured.Glyphs[0].Color.Blue == 255 &&
                measured.Glyphs.Skip(3).Take(3).All(g => ReferenceEquals(g.Font, rest) && g.Color.Green == 128) && measured.Rules.Count == 0,
                "Internal control frames, source paint and occupied outer ink retain distinct ownership, without implicit delimiters.");
            var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(measured, graphics, 20d, 80d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(250d, 200d, graphics.ToString(), [first, rest])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "Only painted cell resources are required; the metric font needs no empty subset.");
        }
    }

    public static void MathFontMatrixDefaultControlUsesParentFrameAndRetainsIndependentCellSizes()
    {
        var font = MathFont(); var resource = Resource("Cells", font, "BC");
        var first = Run("B") with { Text = Run("B").Text with { FontSize = 18d } };
        var matrix = Matrix([first, Run("C")], [Run("C"), Run("B")]);
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var box, out var reason,
            resolveControlMetrics: _ => font), reason ?? "Expected surrounding control inheritance.");
        var measured = TestAssert.NotNull(box);
        TestAssert.True(measured.ParentMetrics?.Size == 12d && measured.Glyphs[0].FontSize == 18d && measured.Glyphs.Skip(1).All(g => g.FontSize == 12d),
            "An absent matrix control uses the surrounding size while every painted cell keeps its authored size.");
    }

    public static void MathFontMatrixFallbackCancellationAndWorkRejectCompleteGrids()
    {
        var font = MathFont(); var resource = Resource("Grid", font, "BC"); var last = Run("C");
        var matrix = Matrix([Run("B"), Run("C")], [Run("B"), last]);
        var missing = Resource("MissingLast", font, "B");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix), r => ReferenceEquals(r, last) ? missing : resource, CancellationToken.None,
            out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
            "A late missing cell encoding cannot publish earlier cells or rows.");
        foreach (var invalid in new[] { matrix with { HasUnsupportedProperties = true }, matrix with { Rows = [matrix.Rows[0], new([new([Run("B")])])] },
            matrix with { Rows = [new([new([]), new([Run("B")])])] },
            matrix with { Rows = [new([new([new DocxMathUnsupported("future", "test", "<future/>", [])])])] },
            matrix with { ControlText = new("", double.NaN, null, false, false, false, null, "TestFont") } })
            TestAssert.True(!DocxMathLayout.TryMeasure(Equation(invalid), _ => resource, CancellationToken.None, out var rejected, out _,
                resolveControlMetrics: _ => font) && rejected is null, "Unsupported rows, empty placeholders, properties and invalid metrics retain complete fallback.");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out absent, out _,
            resolveControlMetrics: _ => TestFontBuilder.LoadTestFont()) && absent is null, "A missing mathematical metric owner cannot publish a default-shaped grid.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(matrix), cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out _, out _, resolveControlMetrics: _ => font));
    }

    private static DocxMathRun Run(string text, string? color = null) => new(new(text, 12d, color, false, false, false, null, "TestFont"), "p", null) { NormalTextValue = true };
    private static DocxMathMatrix Matrix(DocxMathRun[] first, DocxMathRun[] second) => new([new(first.Select(r => new DocxMathRow([r])).ToArray()), new(second.Select(r => new DocxMathRow([r])).ToArray())]);
    private static DocxEquation Equation(DocxMathMatrix matrix) => new(new([matrix]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
    private static PdfFontResource Resource(string name, OpenTypeFont font, string text) => new(name, PdfEmbeddedFont.Create(font, text.Select(c => (int)c), CancellationToken.None));
    private static OpenTypeFont MathFont()
    {
        byte[] table = new byte[224]; BinaryPrimitives.WriteUInt16BigEndian(table, 1); BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(4), 10);
        BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(10), 73);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)OpenTypeMathConstant.AxisHeight * 4), 250);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)OpenTypeMathConstant.MathLeading * 4), 120);
        return OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
    }
}
