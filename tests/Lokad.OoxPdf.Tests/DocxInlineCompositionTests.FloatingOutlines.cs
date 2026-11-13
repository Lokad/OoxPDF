using System.Globalization;
using System.Numerics;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFloatingBoxOutlinesKeepCompleteStrokesOutsideContentInsets()
    {
        // Retained Word floating-table-ordered PDF: all four .5pt edges paint
        // at the authored 216x60pt frame. Before the fix every edge was clipped
        // away by the default content insets despite complete table-label text.
        foreach (string fixture in new[] { "floating-table-ordered", "floating-cell-inline-box" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (bool explicitInsets in new[] { false, true })
        {
            var source = ReadFloatingOutlineSource(fixture, mode, explicitInsets, omitOutline: false);
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            AssertSourceLabelsPaint(page, mode, ["BOX TABLE NORTH", "BOX TABLE SOUTH"], fixture);
            double height = fixture == "floating-table-ordered" ? 60d : 120d;
            var (bounds, clip) = ReadFloatingOutline(page, height).Single();
            const double halfStroke = .25d;
            TestAssert.True(bounds[0] - halfStroke >= clip.X - .001d &&
                bounds[1] - halfStroke >= clip.Y - .001d &&
                bounds[0] + bounds[2] + halfStroke <= clip.X + clip.Width + .001d &&
                bounds[1] + bounds[3] + halfStroke <= clip.Y + clip.Height + .001d,
                "Every edge of the authored floating frame must survive its ancestor clips: " + fixture + "/" + mode);
        }
    }

    public static void DocxFloatingBoxOutlinesKeepNoStrokeAndChildClipOwnership()
    {
        foreach (bool omitOutline in new[] { false, true })
        {
            var source = ReadFloatingOutlineSource("floating-cell-inline-box", OoxPdfDocxMarkupMode.Final,
                explicitInsets: true, omitOutline);
            var drawing = source.FloatingDrawings.Single();
            TestAssert.Equal("152400", drawing.TextBoxInsetLeftValue ?? string.Empty);
            TestAssert.Equal("38100", drawing.TextBoxInsetTopValue ?? string.Empty);
            TestAssert.Equal(omitOutline ? string.Empty : "6350", drawing.TextBoxStrokeWidthEmuValue ?? string.Empty);
            TestAssert.Equal(omitOutline ? string.Empty : "000000", drawing.TextBoxStrokeColorHex ?? string.Empty);
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            AssertSourceLabelsPaint(page, OoxPdfDocxMarkupMode.Final,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH", "INNER BOX LABEL"], "outline=" + !omitOutline);
            TestAssert.Equal(omitOutline ? 0 : 1, ReadFloatingOutline(page, 120d).Count());
            // Painting the outside frame must not enlarge the children's inset
            // viewport. The retained nested inline outline lives inside it.
            var inner = DocxNoteClippingTests.ReadPhysicalPaintFrames(page)
                .Single(frame => frame.Block.EndsWith("120 30 re S", StringComparison.Ordinal));
            double inset = omitOutline ? 0d : .25d;
            TestAssert.True(Math.Abs(inner.Clip.X - (72d + 12d + inset)) < .001d &&
                inner.Clip.Width <= 216d - 18d - 2d * inset + .001d && inner.Clip.Height < 120d,
                "Descendants keep their original inset and table-cell clips.");
        }
    }

    public static void DocxFloatingBoxReviewOutlinesShareThePrintedFrameWithTheirChildren()
    {
        foreach (string fixture in new[] { "floating-table-ordered", "floating-cell-inline-box" })
        foreach (bool omitOutline in new[] { false, true })
        {
            var finalSource = ReadFloatingOutlineSource(fixture, OoxPdfDocxMarkupMode.Final, false, omitOutline);
            var reviewSource = ReadFloatingOutlineSource(fixture, OoxPdfDocxMarkupMode.AllMarkup, false, omitOutline);
            var final = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(finalSource, null, default).ToBlockingEnumerable().Single();
            var review = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(reviewSource, null, default).ToBlockingEnumerable().Single();
            AssertSourceLabelsPaint(review, OoxPdfDocxMarkupMode.AllMarkup,
                ["BOX TABLE NORTH", "BOX TABLE SOUTH"], fixture + "/review outline");
            var finalLabel = DocxCommentsTests.ReadEmbeddedGlyphTextShows(final)
                .First(show => show.Text.Contains("BOX TABLE NORTH", StringComparison.Ordinal));
            var reviewLabel = DocxCommentsTests.ReadEmbeddedGlyphTextShows(review)
                .First(show => show.Text.Contains("BOX TABLE NORTH", StringComparison.Ordinal));
            double scale = reviewLabel.Size / finalLabel.Size;
            TestAssert.True(scale > 0d && scale < 1d, "The real comment requires scaled Review painting.");
            double height = fixture == "floating-table-ordered" ? 60d : 120d;
            // Text size and graphics numbers round independently to .001pt.
            // This .02pt numerical envelope is not an Office visual threshold.
            var frames = ReadFloatingOutline(review, height * scale, 216d * scale, .02d).ToArray();
            TestAssert.Equal(omitOutline ? 0 : 1, frames.Length);
            if (omitOutline) continue;
            var (bounds, clip) = frames[0];
            double halfStroke = .25d * scale;
            TestAssert.True(bounds[0] - halfStroke >= clip.X - .001d &&
                bounds[1] - halfStroke >= clip.Y - .001d &&
                bounds[0] + bounds[2] + halfStroke <= clip.X + clip.Width + .001d &&
                bounds[1] + bounds[3] + halfStroke <= clip.Y + clip.Height + .001d,
                "All four scaled frame strokes survive independently of the child content clip.");
            var original = ReadFloatingOutline(final, height).Single().Bounds;
            TestAssert.True(Math.Abs((reviewLabel.X - bounds[0]) - (finalLabel.X - original[0]) * scale) < .02d &&
                Math.Abs((bounds[1] + bounds[3] - reviewLabel.Y) -
                    (original[1] + original[3] - finalLabel.Y) * scale) < .02d,
                "The outline and its child baseline retain the same inset and printed scale. Absolute Review anchor residuals remain separate.");
        }
    }

    private static DocxDocument ReadFloatingOutlineSource(string fixture, OoxPdfDocxMarkupMode mode,
        bool explicitInsets, bool omitOutline) =>
        ReadMutation("TableGridOrigins/Containers/" + fixture, mode, false, false, (xml, _, _) =>
        {
            XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var shape = xml.Descendants(wps + "wsp").First();
            if (explicitInsets)
            {
                var body = shape.Element(wps + "bodyPr")!;
                foreach (var (name, points) in new[] { ("lIns", 12), ("rIns", 6), ("tIns", 3), ("bIns", 5) })
                    body.SetAttributeValue(name, points * 12700);
            }
            if (omitOutline) shape.Element(wps + "spPr")!.Element(a + "ln")!.Remove();
        });

    private static IEnumerable<(double[] Bounds, PdfRectangle Clip)> ReadFloatingOutline(PdfPage page, double height,
        double width = 216d, double tolerance = .001d)
    {
        foreach (var frame in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        {
            if (!frame.Block.EndsWith(" re S", StringComparison.Ordinal)) continue;
            var bounds = frame.Block.Split(' ').Take(4).Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray();
            if (Math.Abs(bounds[2] - width) > tolerance || Math.Abs(bounds[3] - height) > tolerance) continue;
            TestAssert.Equal(Matrix3x2.Identity, frame.Transform);
            yield return (bounds, frame.Clip);
        }
    }
}
