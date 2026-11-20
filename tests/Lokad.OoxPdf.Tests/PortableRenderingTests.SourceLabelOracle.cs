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

    public static void PortableRenderingSourceLabelOracleChecksPaintAndRestoredState()
    {
        var font = TestAssert.NotNull(FontProgramLoader.Load(new PinnedResolver().Resolve(new FontRequest("DejaVu Sans")), default));
        var source = SourceLabelOraclePage(font, "A");
        PdfExtGStateResource[] states = [new("Off", 0d, 1d), new("Dim", .5d, .5d), new("Tiny", .0004d, 1d)];
        foreach (var (prefix, mode, visible) in new (string, int, bool)[]
        {
            ("", 0, true), ("", 3, false), ("/Off gs\n", 0, false), ("/Tiny gs\n", 0, false),
            ("/Dim gs\n", 0, true), ("q\n/Off gs\nQ\n", 0, true), ("q\n3 Tr\nQ\n", 0, true),
            ("/Off gs\nq\n/Dim gs\nQ\n", 0, false), ("3 Tr\nq\n0 Tr\nQ\n", 0, false),
            ("BT\n3 Tr\nET\n", 0, false), ("", 1, false), ("", 7, false),
            ("", 2, true), ("/Dim gs\n", 2, true), ("/Off gs\n", 2, false), ("/Tiny gs\n", 2, false)
        })
        {
            // Mode 0 here leaves the inherited state alone, as ordinary generated
            // text does. Tr persists across BT/ET and belongs to the q/Q state.
            string content = prefix + (mode == 0 ? source.Content : source.Content.Replace("BT", $"BT\n{mode} Tr", StringComparison.Ordinal));
            var page = new PdfPage(200, 200, content, source.Fonts, [], states, []);
            PdfContentValidator.ValidatePage(page, 0, default);
            if (visible) TestAssert.Equal(1, ReadPaintedSourceLabel(page, "A").Count);
            else TestAssert.Contains("known visible fill paint", TestAssert.Throws<InvalidOperationException>(() =>
                ReadPaintedSourceLabel(page, "A")).Message);
        }
        var mask = new PdfTransparencyGroup(new PdfRectangle(0, 0, 200, 200), "", [], [], [], []);
        foreach (string mode in new[] { "", "2 Tr\n" })
        {
            var masked = new PdfPage(200, 200, "/Masked gs\n/Dim gs\n" + mode + source.Content, source.Fonts, [],
                [.. states, new("Masked", 1d, 1d, "Mask")], [], [], [], groups: [new("Mask", mask)]);
            PdfContentValidator.ValidatePage(masked, 0, default);
            TestAssert.Contains("known visible fill paint", TestAssert.Throws<InvalidOperationException>(() =>
                ReadPaintedSourceLabel(masked, "A")).Message);
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
