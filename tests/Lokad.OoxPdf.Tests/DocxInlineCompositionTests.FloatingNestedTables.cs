using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFloatingNestedTableLabelsStayWithinTheirPaintedClips()
    {
        // The native flat-table source already qualifies the floating frame.
        // Zero-margin wrappers preserve the supported child content. This
        // portable regression checks actual glyph ink, not extraction alone;
        // it makes no new Office-geometry claim for the nested witnesses.
        foreach (int depth in new[] { 0, 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("TableGridOrigins/Containers/floating-table-ordered", mode, false, false,
                (xml, w, _) =>
                {
                    var owner = xml.Descendants(w + "txbxContent").Single();
                    var original = owner.Element(w + "tbl")!;
                    var table = new XElement(original);
                    for (int level = 0; level < depth; level++)
                    {
                        var properties = new XElement(table.Element(w + "tblPr")!);
                        properties.Element(w + "tblBorders")!.RemoveNodes();
                        var margins = new XElement(w + "tcMar", new[] { "left", "right", "top", "bottom" }
                            .Select(edge => new XElement(w + edge, new XAttribute(w + "w", "0"),
                                new XAttribute(w + "type", "dxa"))));
                        var endParagraph = new XElement(w + "p", new XElement(w + "pPr",
                            new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                                new XAttribute(w + "line", "20"), new XAttribute(w + "lineRule", "exact")),
                            new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "2")))));
                        var cell = new XElement(w + "tc",
                            new XElement(w + "tcPr", new XElement(w + "tcW", new XAttribute(w + "w", "3600"),
                                new XAttribute(w + "type", "dxa")), margins), table, endParagraph);
                        table = new XElement(w + "tbl", properties, new XElement(table.Element(w + "tblGrid")!),
                            new XElement(w + "tr", cell));
                    }
                    original.ReplaceWith(table);
                });
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            var labels = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalTextClips(page))
            foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts)))
            {
                string label = show.Text.Trim();
                if (label is not ("BOX TABLE NORTH" or "BOX TABLE SOUTH")) continue;
                TestAssert.True(labels.Add(label), "Each source label paints once.");
                TestAssert.True(Math.Abs(transform.M12) < .000001d && Math.Abs(transform.M21) < .000001d,
                    "The source has axis-aligned glyph frames.");
                var font = show.Font.Font;
                double scale = show.Size / font.UnitsPerEm;
                foreach (char character in label)
                {
                    if (char.IsWhiteSpace(character)) continue;
                    TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(character), out var ink),
                        "The portable font supplies each source glyph.");
                    double first = (show.Y + ink.YMin * scale) * transform.M22 + transform.M32;
                    double last = (show.Y + ink.YMax * scale) * transform.M22 + transform.M32;
                    TestAssert.True(Math.Min(first, last) >= clip.Y - .01d &&
                        Math.Max(first, last) <= clip.Y + clip.Height + .01d,
                        $"A complete floating-table label must survive its ancestor clips: depth={depth}/{mode}/{label}.");
                }
            }
            TestAssert.Equal(2, labels.Count);
        }
    }
}
