using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathNaryTests
{
    public static void MathFontDisplayNaryMinimumAndSlantedLimitsHaveOneOwner()
    {
        foreach (bool subformulaLarger in new[] { true, false })
        foreach (short correction in new short[] { -200, 200 })
        {
            byte[] bytes = FontBytes(0x222B, extendedOperator: true, growthVariants: true);
            var range = TestFontBuilder.GetTableRange(bytes, "MATH");
            byte[] math = bytes.AsSpan(range.Offset, range.Length).ToArray();
            BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(14), (ushort)(subformulaLarger ? 3000 : 1500));
            BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(16), (ushort)(subformulaLarger ? 1500 : 3000));
            int info = BinaryPrimitives.ReadUInt16BigEndian(math.AsSpan(6));
            BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(info + 12), 127);
            BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(info + 26), 127);
            BinaryPrimitives.WriteInt16BigEndian(math.AsSpan(info + 18), correction);
            var font = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(bytes, "MATH", math));
            var mappings = new[] { ((ushort)127, 0x222B), ((ushort)128, 0x222B), ((ushort)130, 0x222B),
                (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') };
            var resource = new PdfFontResource("DisplayIntegral", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
            var nary = Nary("\u222B") with { Superscript = Row("BBBB"), Subscript = Row("CCCCC") };
            var equation = Equation(nary);
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var measured, out var reason,
                resolveControlFont: _ => resource), reason ?? "The complete display group must use both font minima.");
            var box = TestAssert.NotNull(measured);
            var op = box.Glyphs[0]; var upper = box.Glyphs.Skip(2).Take(4).ToArray(); var lower = box.Glyphs.Skip(6).Take(5).ToArray();
            TestAssert.True(op.OriginalGlyph == 127 && ReferenceEquals(op.ControlOwner, nary) && op.Color.Red == 255,
                "The selected display variant retains its actual control font and ink.");
            double center = op.X + op.Advance / 2d;
            double UpperCenter() => (upper[0].X + upper[^1].X + upper[^1].Advance) / 2d;
            double LowerCenter() => (lower[0].X + lower[^1].X + lower[^1].Advance) / 2d;
            double halfCorrection = correction * op.FontSize / font.UnitsPerEm / 2d;
            TestAssert.True(Math.Abs(UpperCenter() - center - halfCorrection) < 1e-9 &&
                Math.Abs(LowerCenter() - center + halfCorrection) < 1e-9,
                "Signed italic correction belongs to the selected glyph, with opposite half shifts for stacked limits.");
            double groupRight = Math.Max(op.X + op.Advance, Math.Max(upper[^1].X + upper[^1].Advance,
                lower[^1].X + lower[^1].Advance));
            TestAssert.True(op.X >= 0d && upper[0].X >= 0d && lower[0].X >= 0d && box.Glyphs[1].X >= groupRight,
                "Wide limits translate as a complete nonnegative group before the body.");
            foreach (var glyph in box.Glyphs)
            {
                TestAssert.True(font.TryGetGlyphBounds(glyph.OriginalGlyph, out var ink), "Every placed glyph has readable bounds.");
                double scale = glyph.FontSize / font.UnitsPerEm;
                TestAssert.True(glyph.X + ink.XMax * scale <= box.Width + 1e-9 &&
                    glyph.BaselineY + ink.YMax * scale <= box.AboveBaseline + 1e-9 &&
                    glyph.BaselineY + ink.YMin * scale >= -box.BelowBaseline - 1e-9,
                    "The complete display box owns all corrected limit ink.");
            }
            var missing = new PdfFontResource("MissingDisplay", PdfEmbeddedFont.CreateForGlyphMappings(font,
                mappings.Where(mapping => mapping.Item1 != 127).ToArray(), default));
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => missing, default, out var absent, out _,
                resolveControlFont: _ => missing) && absent is null, "Missing selected display encoding never publishes partial content.");
        }
    }
}
