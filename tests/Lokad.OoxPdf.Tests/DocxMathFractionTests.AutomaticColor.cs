using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
    public static void MathFontAutomaticColorsKeepGlyphAndRuleOwnershipUntilPainting()
    {
        var resource = RadicalResource();
        foreach (string? declared in new string?[] { null, "auto", "000000", "FFFFFF" })
        {
            DocxMathRow Child(string text) => new([new DocxMathRun(new(text, 12d, declared,
                false, false, false, null, "TestFont"), null, null) { NormalTextValue = true }]);
            var fraction = new DocxMathFraction(Child("B"), Child("C"), "bar");
            var equation = new DocxEquation(new([fraction]), 0, 0, false, null, null, null) { ParentFontSize = 12d };
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var box, out var reason), reason ?? "Expected supported fraction.");
            var measured = TestAssert.NotNull(box);
            RgbColor? authored = DocxTextColor.Read(declared);
            TestAssert.True(measured.Glyphs.All(glyph => glyph.Color == authored) && measured.Rules.Single().Color == authored,
                "Mathematical text and structural rules retain automatic versus authored colors.");
            foreach (var automatic in new[] { new RgbColor(255, 255, 255), default })
            {
                var graphics = new PdfGraphicsBuilder();
                DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, default, automaticColor: automatic);
                string content = graphics.ToString();
                string ink = (authored ?? automatic).Red == 255 ? "1 g" : "0 g";
                TestAssert.Equal(measured.Glyphs.Count + measured.Rules.Count,
                    content.Split(ink + Environment.NewLine, StringSplitOptions.None).Length - 1);
                using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => DocxMathBoxPainter.Paint(measured, new(), 0d, 0d, cancelled.Token, automaticColor: automatic));
            }
        }
    }
}
