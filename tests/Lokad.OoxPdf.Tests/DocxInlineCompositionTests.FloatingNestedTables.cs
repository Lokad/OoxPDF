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
            AssertCompleteFloatingPicture(page, name + "/" + mode);
        }
    }

    public static void DocxFloatingChildPictureAlignmentKeepsContentAcrossRunSplits()
    {
        foreach (string alignment in new[] { "left", "center", "right", "both" })
        foreach (bool split in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("TableGridOrigins/Containers/floating-cell-wide-picture", mode, false, split,
                (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "txbxContent").Last().Elements(w + "p").Single();
                    var properties = paragraph.Element(w + "pPr")!;
                    properties.Element(w + "jc")?.Remove();
                    properties.Add(new XElement(w + "jc", new XAttribute(w + "val", alignment)));
                });
            var page = AssertFloatingLabelsPaint(source, mode,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH", "INNER BOX LABEL"], alignment + "/split=" + split);
            AssertCompleteFloatingPicture(page, alignment + "/" + split + "/" + mode);
        }
    }

    public static void DocxFloatingChildUnsupportedMathKeepsSupportedPictureAndCaption()
    {
        foreach (bool sameParagraph in new[] { true, false })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("TableGridOrigins/Containers/floating-cell-wide-picture", mode, false, false,
                (xml, w, m) =>
                {
                    var owner = xml.Descendants(w + "txbxContent").Last();
                    var equation = new XElement(m + "oMath", new XElement(m + "future"));
                    if (sameParagraph) owner.Elements(w + "p").Single().Elements(w + "r").Last().AddBeforeSelf(equation);
                    else owner.Add(new XElement(w + "p", equation));
                });
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var page = AssertFloatingLabelsPaint(source, mode,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH", "INNER BOX LABEL"], "unsupported math/" + sameParagraph,
                diagnostics.Add);
            AssertCompleteFloatingPicture(page, "unsupported math/" + sameParagraph + "/" + mode);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "The unsupported formula retains its warning without publishing an unused math resource.");
        }
    }

    private static void AssertCompleteFloatingPicture(PdfPage page, string context)
    {
        var (bounds, clip) = ReadPictureBounds(page).Single();
        TestAssert.True(bounds.Width > 0d && bounds.Height > 0d && Math.Abs(bounds.Width / bounds.Height - 2d) < .001d,
            "The complete source picture retains its aspect ratio.");
        TestAssert.True(bounds.X >= clip.X - .01d && bounds.Y >= clip.Y - .01d &&
            bounds.X + bounds.Width <= clip.X + clip.Width + .01d &&
            bounds.Y + bounds.Height <= clip.Y + clip.Height + .01d,
            "A fitting source picture cannot be lost to its child-box clip: " + context);
    }

    private static PdfPage AssertFloatingLabelsPaint(DocxDocument source, OoxPdfDocxMarkupMode mode,
        IReadOnlyList<string> expectedLabels, string context,
        Action<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>? diagnosticSink = null)
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var page = renderer.RenderPagesAsync(source, diagnosticSink, default).ToBlockingEnumerable().Single();
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
        var paints = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).SelectMany(frame =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, frame.Block, page.Fonts))
                .Select(show => (frame.Block, frame.Clip, frame.Transform, Show: show))).ToArray();
        // Justification and equivalent run splits may emit a label in several
        // text shows. Match its complete source characters, then check each
        // show's ink in its own frame rather than requiring one Tj per label.
        string text = string.Concat(paints.SelectMany(paint => paint.Show.Text).Where(c => !char.IsWhiteSpace(c)));
        var labels = expectedLabels.Select(label =>
        {
            string compact = string.Concat(label.Where(c => !char.IsWhiteSpace(c)));
            int start = text.IndexOf(compact, StringComparison.Ordinal);
            TestAssert.True(start >= 0 && start == text.LastIndexOf(compact, StringComparison.Ordinal),
                "Each complete source label paints once: " + context + "/" + mode + "/" + label);
            return (Start: start, End: start + compact.Length);
        }).ToArray();
        int characterIndex = 0;
        foreach (var (block, clip, transform, show) in paints)
        {
            int start = characterIndex;
            characterIndex += show.Text.Count(c => !char.IsWhiteSpace(c));
            if (!labels.Any(label => label.Start < characterIndex && label.End > start)) continue;
            TestAssert.True(Math.Abs(transform.M12) < .000001d && Math.Abs(transform.M21) < .000001d,
                "The source has axis-aligned glyph frames.");
            TestAssert.True(Math.Abs(transform.M11 - 1f) < .000001d && Math.Abs(transform.M22 - 1f) < .000001d &&
                Math.Abs(transform.M31) < .000001d && Math.Abs(transform.M32) < .000001d,
                "These generated cell fills and glyphs use physical page coordinates.");
            int paintOffset = page.Content.IndexOf(block, StringComparison.Ordinal);
            var font = show.Font.Font;
            double scale = show.Size / font.UnitsPerEm;
            double cursor = show.X;
            int index = start;
            foreach (char character in show.Text)
            {
                ushort glyph = font.MapCodePoint(character);
                double glyphX = cursor;
                cursor += font.GetAdvanceWidth(glyph) * scale;
                if (char.IsWhiteSpace(character)) continue;
                bool inLabel = labels.Any(label => index >= label.Start && index < label.End);
                index++;
                if (!inLabel) continue;
                TestAssert.True(font.TryGetGlyphBounds(glyph, out var ink),
                    "The portable font supplies each source glyph.");
                double first = (show.Y + ink.YMin * scale) * transform.M22 + transform.M32;
                double last = (show.Y + ink.YMax * scale) * transform.M22 + transform.M32;
                TestAssert.True(Math.Min(first, last) >= clip.Y - .01d &&
                    Math.Max(first, last) <= clip.Y + clip.Height + .01d,
                    $"A complete floating-table label must survive its ancestor clips: {context}/{mode}/{show.Text}.");
                double left = glyphX + ink.XMin * scale, right = glyphX + ink.XMax * scale;
                TestAssert.True(left >= clip.X - .01d && right <= clip.X + clip.Width + .01d,
                    $"A floating-table label must retain its horizontal ink: {context}/{mode}/{show.Text}.");
                TestAssert.True(!fills.Any(fill => fill.Offset > paintOffset &&
                    Math.Min(fill.X + fill.Width, right) - Math.Max(fill.X, left) > .01d &&
                    Math.Min(fill.Y + fill.Height, Math.Max(first, last)) - Math.Max(fill.Y, Math.Min(first, last)) > .01d),
                    $"A cell background or picture must not cover a source label: {context}/{mode}/{show.Text}.");
            }
        }
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
