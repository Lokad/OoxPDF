using System.Numerics;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingJustifiedInlineBoxChildKeepsPreparedParentContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid",
                mode, false, false), mode).Single();
            var expected = ReadPaintedSourceLabel(baseline, "x+2y", mathOnly: true);
            TestAssert.Equal(4, expected.Count);
            foreach (bool unsupportedChild in new[] { false, true })
            {
                var source = DocxInlineCompositionTests.ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid", mode, false, false,
                    (xml, w, m) =>
                    {
                        var paragraph = xml.Descendants(w + "txbxContent").Single().Elements(w + "p").Single();
                        if (unsupportedChild) paragraph.Add(new XElement(m + "oMath", new XElement(m + "future")));
                        else SetJustified(paragraph, w);
                    });
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = RenderAlignmentProbe(source, mode, diagnostics.Add).Single();
                var actual = ReadPaintedSourceLabel(page, "x+2y", mathOnly: true);
                TestAssert.Equal(expected.Count, actual.Count);
                foreach (var pair in expected.Zip(actual)) AssertSamePaintOrigin(pair.First, pair.Second);
                ReadPaintedSourceLabel(page, "Inline box content.");
                TestAssert.Equal(unsupportedChild ? 1 : 0, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            }
        }
    }

    public static void PortableRenderingJustifiedStaticParagraphsKeepFooterInkAndPhysicalFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (string part in new[] { "word/header1.xml", "word/footer1.xml" })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                    SetJustified(xml.Descendants(w + "p").First(p => !p.Descendants(m + "oMath").Any() &&
                        p.Descendants(w + "t").Any(t => t.Value.Length != 0)), w), partOverride: part);
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var before = ReadPaintedSourceLabel(baseline[index], footer);
                    var after = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(before.Count, after.Count);
                    AssertSamePaintOrigin(before[0], after[0]);
                    AssertSameStaticMathPaint(baseline[index], pages[index]);
                }
            }
        }
    }

    public static void PortableRenderingPlainStaticCountersKeepFooterInkAndIndependentFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (string part in new[] { "word/header1.xml", "word/footer1.xml" })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var paragraph = xml.Descendants(w + "p").First(p => !p.Descendants(m + "oMath").Any() &&
                        p.Descendants(w + "t").Any(t => t.Value.Length != 0));
                    void Text(string value) => paragraph.Add(new XElement(w + "r", new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), value)));
                    void Field(string instruction) => paragraph.Add(new XElement(w + "fldSimple",
                        new XAttribute(w + "instr", instruction), new XElement(w + "r", new XElement(w + "t", "99"))));
                    Text(" Page "); Field(" PAGE "); Text(" of "); Field(" NUMPAGES ");
                }, partOverride: part);
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var before = ReadPaintedSourceLabel(baseline[index], footer);
                    var after = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(before.Count, after.Count);
                    AssertSamePaintOrigin(before[0], after[0]);
                    ReadPaintedSourceLabel(pages[index], "Page " + (index + 1) + " of " + pages.Length);
                    AssertSameStaticMathPaint(baseline[index], pages[index]);
                }
            }
        }
    }

    private static void AssertSameStaticMathPaint(PdfPage before, PdfPage after)
    {
        var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
        var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
        TestAssert.True(expected.Select(show => (show.Text, show.Size, show.X, show.Y))
            .SequenceEqual(actual.Select(show => (show.Text, show.Size, show.X, show.Y))),
            "Equivalent static paragraph edits retain neighboring prepared math and physical frames.");
    }

    private static PdfPage[] RenderAlignmentProbe(DocxDocument document, OoxPdfDocxMarkupMode mode,
        Action<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>? diagnosticSink = null) =>
        new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, diagnosticSink, default).ToBlockingEnumerable().ToArray();

    private static void SetJustified(XElement paragraph, XNamespace w)
    {
        var properties = paragraph.Element(w + "pPr");
        if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
        properties.Element(w + "jc")?.Remove();
        properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "both")));
    }

    private static void AssertSamePaintOrigin((Vector2 Origin, double Size) before, (Vector2 Origin, double Size) after) =>
        TestAssert.True(Vector2.Distance(before.Origin, after.Origin) < .002f && Math.Abs(before.Size - after.Size) < .001d,
            "The equivalent single-line alignment keeps its source origin and physical font size.");

    private static IReadOnlyList<(Vector2 Origin, double Size)> ReadPaintedSourceLabel(PdfPage page,
        string label, bool mathOnly = false)
    {
        // Match source characters across text-show boundaries, then check their
        // font ink against each physical paint frame. Existing readers own PDF syntax.
        var characters = new List<(int Offset, ushort Glyph, PdfEmbeddedFont Font, double Size,
            double X, double Y, PdfRectangle Clip, Matrix3x2 Transform)>();
        var text = new StringBuilder();
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts),
            usePageCoordinates: false, honorActualText: false,
            fontResourceFilter: mathOnly ? name => name.StartsWith("Eq", StringComparison.Ordinal) : null))
        {
            double x = show.X, scale = show.Size / show.Font.Font.UnitsPerEm;
            foreach (var rune in show.Text.EnumerateRunes())
            {
                ushort glyph = show.Font.Font.MapCodePoint(rune.Value);
                if (!Rune.IsWhiteSpace(rune))
                {
                    characters.Add((text.Length, glyph, show.Font, show.Size, x, show.Y, clip, transform));
                    text.Append(rune.ToString().Normalize(NormalizationForm.FormKC));
                }
                x += show.Font.Font.GetAdvanceWidth(glyph) * scale;
            }
        }
        string wanted = string.Concat(label.Where(c => !char.IsWhiteSpace(c))).Normalize(NormalizationForm.FormKC);
        string actual = text.ToString();
        int start = actual.IndexOf(wanted, StringComparison.Ordinal);
        TestAssert.True(start >= 0 && start == actual.LastIndexOf(wanted, StringComparison.Ordinal),
            "Each complete source label paints once: " + label);
        var result = new List<(Vector2 Origin, double Size)>();
        foreach (var character in characters.Where(c => c.Offset >= start && c.Offset < start + wanted.Length))
        {
            var font = character.Font.Font;
            TestAssert.True(font.TryGetGlyphBounds(character.Glyph, out var ink), "Pinned fonts supply the source glyph ink.");
            double scale = character.Size / font.UnitsPerEm;
            foreach (double x in new[] { ink.XMin, ink.XMax })
            foreach (double y in new[] { ink.YMin, ink.YMax })
            {
                var point = Vector2.Transform(new((float)(character.X + x * scale), (float)(character.Y + y * scale)), character.Transform);
                TestAssert.True(point.X >= character.Clip.X - .01d && point.X <= character.Clip.X + character.Clip.Width + .01d &&
                    point.Y >= character.Clip.Y - .01d && point.Y <= character.Clip.Y + character.Clip.Height + .01d,
                    "Every source glyph must remain inside the paper and effective clips: " + label);
            }
            result.Add((Vector2.Transform(new((float)character.X, (float)character.Y), character.Transform), character.Size));
        }
        TestAssert.Equal(wanted.Length, result.Count);
        return result;
    }
}
