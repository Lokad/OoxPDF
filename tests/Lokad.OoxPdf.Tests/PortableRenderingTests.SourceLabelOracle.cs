using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingSourceLabelOracleKeepsScalarAndGlyphUnits()
    {
        var resolver = new PinnedResolver();
        var font = TestAssert.NotNull(FontProgramLoader.Load(resolver.Resolve(new FontRequest("DejaVu Sans")), default));
        foreach (string text in new[] { "A", "\U0001f600", "\ufb01", "A\U0001f600\ufb01" })
        {
            var page = SourceLabelOraclePage(font, text);
            PdfContentValidator.ValidatePage(page, 0, default);
            var ink = ReadPaintedSourceLabel(page, text);
            TestAssert.Equal(text == "A\U0001f600\ufb01" ? 3 : 1, ink.Count);
            TestAssert.True(ink.All(glyph => glyph.Ink.Width > 0 && glyph.Ink.Height > 0), "Every authored glyph has visible bounded ink.");
        }
    }

    public static void PortableRenderingSourceLabelOracleRejectsMissingClippedAndPartialGlyphs()
    {
        var resolver = new PinnedResolver();
        var font = TestAssert.NotNull(FontProgramLoader.Load(resolver.Resolve(new FontRequest("DejaVu Sans")), default));
        var missing = TestAssert.Throws<InvalidOperationException>(() => ReadPaintedSourceLabel(SourceLabelOraclePage(font, "A"), "B"));
        TestAssert.Contains("Each complete source label paints once", missing.Message);
        var clipped = TestAssert.Throws<InvalidOperationException>(() => ReadPaintedSourceLabel(
            SourceLabelOraclePage(font, "\U0001f600", clipped: true), "\U0001f600"));
        TestAssert.Contains("Every source glyph must remain inside the paper and effective clips", clipped.Message);
        var partial = TestAssert.Throws<InvalidOperationException>(() => ReadPaintedSourceLabel(SourceLabelOraclePage(font, "\ufb01"), "f"));
        TestAssert.Contains("A source label must own complete glyphs", partial.Message);
    }

    public static void PortableRenderingSourceLabelOracleRejectsDisplacedAndAlternateGlyphs()
    {
        var font = TestAssert.NotNull(FontProgramLoader.Load(new PinnedResolver().Resolve(new FontRequest("DejaVu Sans")), default));
        var embedded = PdfEmbeddedFont.Create(font, ['A', 'B'], default);
        string hex = embedded.EncodeGlyphHex("AB");
        foreach (var (spacing, show) in new[] { (100, $"<{hex}> Tj"), (0, $"[<{hex[..4]}> -10000 <{hex[4..]}>] TJ") })
        {
            var page = new PdfPage(100, 100,
                $"q\n0 0 70 100 re W n\nBT\n/F1 12 Tf\n{spacing} Tc\n1 0 0 1 10 50 Tm\n{show}\nET\nQ\n",
                [new PdfFontResource("F1", embedded)]);
            PdfContentValidator.ValidatePage(page, 0, default);
            var failure = TestAssert.Throws<InvalidOperationException>(() => ReadPaintedSourceLabel(page, "AB"));
            TestAssert.Contains("Every source glyph must remain inside the paper and effective clips", failure.Message);
        }
        // The logical I has the physical W outline; cmap(I) cannot certify its ink.
        var alternate = PdfEmbeddedFont.CreateForGlyphMappings(font, [(font.MapCodePoint('W'), (int)'I')], default);
        TestAssert.True(alternate.TryGetEncodedCid(font.MapCodePoint('W'), out ushort cid), "The authored alternate outline is embedded.");
        var variant = new PdfPage(100, 100,
            $"q\n0 0 16 100 re W n\nBT\n/F1 24 Tf\n0 Tc\n1 0 0 1 10 50 Tm\n<{cid:X4}> Tj\nET\nQ\n",
            [new PdfFontResource("F1", alternate)]);
        PdfContentValidator.ValidatePage(variant, 0, default);
        var clipped = TestAssert.Throws<InvalidOperationException>(() => ReadPaintedSourceLabel(variant, "I"));
        TestAssert.Contains("Every source glyph must remain inside the paper and effective clips", clipped.Message);
    }

    public static void PortableRenderingSourceLabelOracleKeepsEmittedAdvancesAndTextTransforms()
    {
        var font = TestAssert.NotNull(FontProgramLoader.Load(new PinnedResolver().Resolve(new FontRequest("DejaVu Sans")), default));
        var embedded = PdfEmbeddedFont.Create(font, ['A', 'B'], default);
        string hex = embedded.EncodeGlyphHex("AB");
        foreach (var (spacing, show) in new[] { (3, $"<{hex}> Tj"), (0, $"[<{hex[..4]}> -250 <{hex[4..]}>] TJ") })
        {
            var page = new PdfPage(200, 200,
                $"BT\n/F1 12 Tf\n{spacing} Tc\n1 0.5 0 1 10 50 Tm\n{show}\nET\n",
                [new PdfFontResource("F1", embedded)]);
            PdfContentValidator.ValidatePage(page, 0, default);
            var glyphs = ReadPaintedSourceLabel(page, "AB");
            TestAssert.Equal(2, glyphs.Count);
            // Pinned DejaVu Sans A: emitted /W 684, 12pt, plus 3pt tracking.
            TestAssert.True(Math.Abs(glyphs[1].Origin.X - 21.208d) < .001d && Math.Abs(glyphs[1].Origin.Y - 55.604d) < .001d,
                "The second glyph uses emitted PDF width, spacing and the authored text matrix.");
        }
    }

    private static PdfPage SourceLabelOraclePage(OpenTypeFont font, string text, bool clipped = false)
    {
        var embedded = PdfEmbeddedFont.Create(font, text.EnumerateRunes().Select(rune => rune.Value), default);
        var graphics = new PdfGraphicsBuilder();
        if (clipped) { graphics.SaveState(); graphics.ClipRectangle(0, 0, 10, 10); }
        graphics.DrawGlyphText("F1", 24, 40, 80, 0, 0, 0, embedded.EncodeGlyphHex(text), false, 0, 0, 0, 0, 0, 0);
        if (clipped) graphics.RestoreState();
        return new PdfPage(200, 200, graphics.ToString(), [new PdfFontResource("F1", embedded)]);
    }
}
