using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathNaryTests
{
    public static void MathFontNaryIntegralFamilySharesDefaultsVariantsAndCompleteFallback()
    {
        foreach (int point in new[] { 0x222B, 0x222C, 0x222D })
        {
            var font = OpenTypeFont.Load(FontBytes(point, extendedOperator: true, growthVariants: true));
            var mappings = new[] { ((ushort)127, point), ((ushort)128, point), ((ushort)130, point),
                (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') };
            var resource = new PdfFontResource("IntegralFamily", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
            foreach (bool display in new[] { false, true })
            foreach (string location in new[] { "subSup", "undOvr" })
            {
                var nary = Nary(char.ConvertFromUtf32(point)) with { LimitLocationValue = null };
                var equation = Equation(nary) with { Display = display,
                    LimitSettings = Settings($"<m:naryLim m:val='future'/><m:intLim m:val='{location}'/>") };
                TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var inherited, out var reason),
                    reason ?? "Every supported integral uses its independent integral default.");
                var box = TestAssert.NotNull(inherited);
                TestAssert.Equal(point, box.Glyphs[0].SourceCodePoint);
                TestAssert.Equal(4, box.Glyphs.Count);
                TestAssert.True(DocxMathLayout.TryMeasure(equation with { Content = new([nary with { LimitLocationValue = location }]) },
                    _ => resource, default, out var explicitBox, out _), "Equivalent authored placement remains supported.");
                TestAssert.True(PaintPlacements(box).SequenceEqual(PaintPlacements(TestAssert.NotNull(explicitBox))),
                    "The family shares measured placement, while its source default remains unauthored.");
                TestAssert.True(nary.LimitLocationValue is null, "Resolution cannot rewrite source settings.");

                var malformed = equation with { LimitSettings = Settings("<m:intLim m:val='future'/>") };
                TestAssert.True(!DocxMathLayout.TryMeasure(malformed, _ => resource, default, out var absent, out _) && absent is null,
                    "A malformed used integral default rejects the complete equation.");
                TestAssert.True(DocxMathLayout.TryMeasure(malformed with { Content = new([nary with { LimitLocationValue = location }]) },
                    _ => resource, default, out _, out _), "An explicit node still overrides malformed defaults.");
                var missing = new PdfFontResource("MissingIntegralVariant", PdfEmbeddedFont.CreateForGlyphMappings(font,
                    mappings.Where(mapping => mapping.Item1 != box.Glyphs[0].OriginalGlyph), default));
                TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => resource, default, out absent, out _,
                    resolveControlFont: _ => missing) && absent is null, "A missing selected CID cannot publish a partial integral.");
                using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
            }
        }
    }
}
