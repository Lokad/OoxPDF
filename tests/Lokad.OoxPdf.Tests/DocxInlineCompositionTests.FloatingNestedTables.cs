using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;
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
            AssertFloatingLabelsPaint(source, mode, ["BOX TABLE NORTH", "BOX TABLE SOUTH"], "depth=" + depth);
        }
    }

    public static void DocxFloatingTableCellInlineBoxLabelsRemainVisibleInFinalAndReview()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-cell-inline-box.docx"));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            AssertFloatingLabelsPaint(DocxTests.ReadDocx(path, mode), mode,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH", "INNER BOX LABEL"], "inline child");
    }

    public static void DocxFloatingChildPicturesKeepCompleteSourceAndClearLabelInk()
    {
        foreach (string name in new[] { "inline", "wide", "block" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "TableGridOrigins", "Containers", "floating-cell-" + name + "-picture.docx"));
            var page = AssertFloatingLabelsPaint(DocxTests.ReadDocx(path, mode), mode,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH", "INNER BOX LABEL"], name + " picture");
            var pictures = ReadPictureBounds(page).ToArray();
            TestAssert.Equal(1, pictures.Length);
            var (bounds, clip) = pictures[0];
            TestAssert.True(bounds.Width > 0d && bounds.Height > 0d && Math.Abs(bounds.Width / bounds.Height - 2d) < .001d,
                "The complete source picture retains its aspect ratio.");
            TestAssert.True(bounds.X >= clip.X - .01d && bounds.Y >= clip.Y - .01d &&
                bounds.X + bounds.Width <= clip.X + clip.Width + .01d &&
                bounds.Y + bounds.Height <= clip.Y + clip.Height + .01d,
                $"A fitting source picture cannot be lost to its child-box clip: {name}/{mode}.");
        }
    }

    private static PdfPage AssertFloatingLabelsPaint(DocxDocument source, OoxPdfDocxMarkupMode mode,
        IReadOnlyList<string> expectedLabels, string context)
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
        const string number = @"[-+\d.eE]+";
        // These fixtures contain opaque pictures with separate inline advances.
        // A picture may not intersect label ink, regardless of paint order.
        var fills = Regex.Matches(page.Content, @"(?m)^(?<x>" + number + @") (?<y>" + number +
            @") (?<w>" + number + @") (?<h>" + number + @") re f\r?$").Cast<Match>().Select(match =>
            (Offset: match.Index, X: ReadNumber(match, "x"), Y: ReadNumber(match, "y"),
                Width: ReadNumber(match, "w"), Height: ReadNumber(match, "h")))
            .Concat(ReadPictureBounds(page).Select(picture =>
                (Offset: int.MaxValue, X: picture.Bounds.X, Y: picture.Bounds.Y,
                    Width: picture.Bounds.Width, Height: picture.Bounds.Height))).ToArray();
        var labels = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts)))
        {
            string label = show.Text.Trim();
            if (!expectedLabels.Contains(label, StringComparer.Ordinal)) continue;
            TestAssert.True(labels.Add(label), "Each source label paints once.");
            TestAssert.True(Math.Abs(transform.M12) < .000001d && Math.Abs(transform.M21) < .000001d,
                "The source has axis-aligned glyph frames.");
            TestAssert.True(Math.Abs(transform.M11 - 1f) < .000001d && Math.Abs(transform.M22 - 1f) < .000001d &&
                Math.Abs(transform.M31) < .000001d && Math.Abs(transform.M32) < .000001d,
                "These generated cell fills and glyphs use physical page coordinates.");
            int paintOffset = page.Content.IndexOf(block, StringComparison.Ordinal);
            var font = show.Font.Font;
            double scale = show.Size / font.UnitsPerEm;
            double cursor = show.X;
            foreach (char character in label)
            {
                ushort glyph = font.MapCodePoint(character);
                double glyphX = cursor;
                cursor += font.GetAdvanceWidth(glyph) * scale;
                if (char.IsWhiteSpace(character)) continue;
                TestAssert.True(font.TryGetGlyphBounds(glyph, out var ink),
                    "The portable font supplies each source glyph.");
                double first = (show.Y + ink.YMin * scale) * transform.M22 + transform.M32;
                double last = (show.Y + ink.YMax * scale) * transform.M22 + transform.M32;
                TestAssert.True(Math.Min(first, last) >= clip.Y - .01d &&
                    Math.Max(first, last) <= clip.Y + clip.Height + .01d,
                    $"A complete floating-table label must survive its ancestor clips: {context}/{mode}/{label}.");
                double left = glyphX + ink.XMin * scale, right = glyphX + ink.XMax * scale;
                TestAssert.True(!fills.Any(fill => fill.Offset > paintOffset &&
                    Math.Min(fill.X + fill.Width, right) - Math.Max(fill.X, left) > .01d &&
                    Math.Min(fill.Y + fill.Height, Math.Max(first, last)) - Math.Max(fill.Y, Math.Min(first, last)) > .01d),
                    $"A cell background or picture must not cover a source label: {context}/{mode}/{label}.");
            }
        }
        TestAssert.Equal(expectedLabels.Count, labels.Count);
        return page;

        static double ReadNumber(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    }

    private static IEnumerable<(PdfRectangle Bounds, PdfRectangle Clip)> ReadPictureBounds(PdfPage page)
    {
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        {
            if (!page.Images.Any(image => block == "/" + image.ResourceName + " Do")) continue;
            TestAssert.True(Math.Abs(transform.M12) < .000001d && Math.Abs(transform.M21) < .000001d,
                "The public picture sources have axis-aligned geometry.");
            Vector2 first = Vector2.Transform(Vector2.Zero, transform);
            Vector2 last = Vector2.Transform(Vector2.One, transform);
            yield return (new PdfRectangle(Math.Min(first.X, last.X), Math.Min(first.Y, last.Y),
                Math.Abs(last.X - first.X), Math.Abs(last.Y - first.Y)), clip);
        }
    }
}
