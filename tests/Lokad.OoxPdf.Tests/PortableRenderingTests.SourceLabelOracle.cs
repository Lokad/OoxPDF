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
