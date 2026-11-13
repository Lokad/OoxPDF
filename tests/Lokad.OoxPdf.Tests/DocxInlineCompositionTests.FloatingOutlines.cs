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

    private static IEnumerable<(double[] Bounds, PdfRectangle Clip)> ReadFloatingOutline(PdfPage page, double height)
    {
        foreach (var frame in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        {
            if (!frame.Block.EndsWith(" re S", StringComparison.Ordinal)) continue;
            var bounds = frame.Block.Split(' ').Take(4).Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray();
            if (Math.Abs(bounds[2] - 216d) > .001d || Math.Abs(bounds[3] - height) > .001d) continue;
            TestAssert.Equal(Matrix3x2.Identity, frame.Transform);
            yield return (bounds, frame.Clip);
        }
    }
}
