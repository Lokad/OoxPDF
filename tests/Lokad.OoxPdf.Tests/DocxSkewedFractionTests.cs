using System.Security.Cryptography;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxSkewedFractionTests
{
    public static void MathFontSkewedFractionsKeepLiteralAndControlOwnersAndCompleteFallbacks()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "PortableFonts", "DejaVuMathTeXGyre.ttf"));
        byte[] bytes = File.ReadAllBytes(path);
        TestAssert.Equal("75A6CB527B0175F9D13AF29C33A6915AE38645016C6FE9881653155CD1F62485",
            Convert.ToHexString(SHA256.HashData(bytes)));
        var font = OpenTypeFont.Load(bytes);
        var mappings = new List<(ushort Glyph, int CodePoint)>();
        foreach (int point in new[] { (int)'B', (int)'C' }) mappings.Add((font.MapCodePoint(point), point));
        ushort canonical = font.MapCodePoint(0x2044);
        for (double height = 0d; font.TryGetMathVerticalVariant(canonical, height, default, out var variant); height = variant.Advance + 1d)
            mappings.Add((variant.Glyph, 0x2044));
        var complete = new PdfFontResource("Slash", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
        var missing = new PdfFontResource("NoSlash", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings.Where(m => m.CodePoint != 0x2044), default));
        var text = new PdfFontResource("Literal", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), "BC".Select(c => (int)c), default));
        DocxMathRow Literal(string value) => new([new DocxMathRun(new(value, 12d, "008000", false, false, false, null, "Text"), null, null)
            { NormalTextValue = true }]);
        foreach (bool display in new[] { false, true })
        {
            var fraction = new DocxMathFraction(Literal("B"), Literal("C"), "skw")
                { ControlText = new("", 20d, "FF0000", false, false, false, null, "Math") };
            var equation = new DocxEquation(new([fraction]), 0, 0, display, null, null, null) { ParentFontSize = 11d };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var sources) && sources.Count == 2 &&
                sources.All(source => source.ScriptLevel == 0) && ReferenceEquals(fraction, sources[0].ControlOwner),
                "The actual slash control is discovered without reducing its literal children to script style.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => text, default, out var measured, out var reason,
                resolveControlFont: _ => complete), reason ?? "Expected a bounded skewed fraction with independent text/control fonts.");
            var box = TestAssert.NotNull(measured);
            TestAssert.Equal(3, box.Glyphs.Count);
            TestAssert.Equal(0, box.Rules.Count);
            TestAssert.Equal(20d, box.ParentMetrics!.Size);
            var slash = box.Glyphs.Single(g => g.SourceCodePoint == 0x2044);
            TestAssert.True(ReferenceEquals(slash.ControlOwner, fraction) && ReferenceEquals(slash.Font, complete) && slash.FontSize == 20.04d &&
                slash.Color.Red == 255 && slash.Color.Green == 0d && slash.Color.Blue == 0d,
                "Authored structural font, size and red ink stay on the slash.");
            TestAssert.True(box.Glyphs.Take(2).All(g => ReferenceEquals(g.Font, text) && g.FontSize == 12d && g.ControlOwner is null) &&
                box.Glyphs[0].BaselineY > box.Glyphs[1].BaselineY && box.Glyphs[0].X < box.Glyphs[1].X,
                "The ordered literal sources retain their own size/font while occupying diagonal roles.");
            foreach (var glyph in box.Glyphs)
            {
                TestAssert.True(glyph.Font.Font.Font.TryGetGlyphBounds(glyph.OriginalGlyph, out var ink), "Every emitted glyph has readable ink bounds.");
                double scale = glyph.FontSize / glyph.Font.Font.Font.UnitsPerEm;
                TestAssert.True(glyph.X + ink.XMin * scale >= -1e-9 && glyph.X + ink.XMax * scale <= box.Width + 1e-9 &&
                    glyph.BaselineY + ink.YMax * scale <= box.AboveBaseline + 1e-9 &&
                    glyph.BaselineY + ink.YMin * scale >= -box.BelowBaseline - 1e-9, "The occupied box contains actual child/slash ink.");
            }
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => text, default, out var absent, out _,
                resolveControlFont: _ => missing) && absent is null, "A missing slash CID cannot publish a numerator/denominator prefix.");
            var matrix = equation with { Content = new([new DocxMathMatrix([new([new([fraction])])])]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(matrix, default, out var matrixSources) &&
                matrixSources[0].ControlOwners.Any(owner => ReferenceEquals(owner, fraction)) &&
                DocxMathLayout.TryMeasure(matrix, _ => text, default, out var matrixBox, out _,
                    resolveControlFont: _ => complete, resolveControlMetrics: _ => font) &&
                matrixBox!.Glyphs.Count(g => ReferenceEquals(g.ControlOwner, fraction)) == 1,
                "A matrix cell discovers and paints its actual nested slash control once.");
            foreach (var invalid in new[] { fraction with { TypeValue = "future" }, fraction with { HasUnsupportedProperties = true },
                fraction with { Numerator = new([]) }, fraction with { Numerator = new([fraction]) },
                fraction with { Denominator = new([new DocxMathScripts(Literal("B"), null, Literal("C"))]) } })
                TestAssert.True(!DocxMathLayout.TryMeasure(equation with { Content = new([invalid]) }, _ => text, default,
                    out var partial, out _, resolveControlFont: _ => complete) && partial is null,
                    "Unknown, empty or deeper skewed arguments retain complete fallback.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancellation.Token, out _));
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => text, cancellation.Token, out _, out _));
        }
    }
}
