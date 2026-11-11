using System.Xml.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxTerminalInlinePicturesRetainCompleteFramesAcrossContainersAndRuns()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (width, pictures) in new[] { (24, 1), (240, 1), (88, 2) })
        foreach (bool trailingSpace in new[] { false, true })
        foreach (bool merge in new[] { false, true })
        {
            var source = ReadTerminalPicture(container, mode, width, trailingSpace, merge, pictures);
            var renderer = PictureRunBoundaryRenderer(mode, mathFont);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var page = pages.Single(page => page.Images.Count != 0);
            var frames = Regex.Matches(page.Content,
                @"(?<Width>[-\d.]+) 0 0 (?<Height>[-\d.]+) (?<X>[-\d.]+) (?<Y>[-\d.]+) cm\s+/Im\d+ Do");
            TestAssert.Equal(pictures, frames.Count);
            foreach (Match frame in frames)
            {
                double Number(string name) => double.Parse(frame.Groups[name].Value, CultureInfo.InvariantCulture);
                TestAssert.True(Number("X") >= source.MarginLeftPoints - .001d &&
                    Number("X") + Number("Width") <= page.Width - source.MarginRightPoints + .001d &&
                    Number("Y") >= 0d && Number("Y") + Number("Height") <= page.Height,
                    "Fitting pictures stay visible when no following word triggers wrapping, including adjacent drawings.");
                TestAssert.True(Math.Abs(Number("Width") - width) < .001d && Math.Abs(Number("Height") - 18d) < .001d,
                    "Wrapping preserves the authored picture dimensions.");
            }
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.Contains("Before", text);
            TestAssert.Contains(new string('i', 100), text);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null,
                cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxTerminalInlinePicturesKeepExistingUnfitWidthFallback()
    {
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadTerminalPicture("body", mode, 600, trailingSpace: false, merge: false);
            var page = PictureRunBoundaryRenderer(mode, mathFont).RenderPagesAsync(source, null, default)
                .ToBlockingEnumerable().Single();
            var frame = ReadPictureRangeFrame(page);
            var before = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.StartsWith("Before", StringComparison.Ordinal));
            // The source object cannot fit either line. Its existing bounded
            // fit/placement remains until independently qualified, rather than
            // repeatedly moving an oversized object to another empty row.
            TestAssert.True(frame.X > before.X + .001d &&
                frame.Width <= source.PageWidthPoints - source.MarginLeftPoints - source.MarginRightPoints + .001d,
                "An oversized source retains its established bounded fallback.");
        }
    }

    private static DocxDocument ReadTerminalPicture(string container, OoxPdfDocxMarkupMode mode,
        int width, bool trailingSpace, bool merge, int pictures = 1) => ReadMutation("docx-picture-rotated-raster", mode, false, false,
        (xml, w, _) =>
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            var paragraph = xml.Descendants(w + "p").Single();
            var runs = paragraph.Elements(w + "r").ToArray();
            runs[0].Element(w + "t")!.Value = "Before " + new string('i', 100) + " ";
            var drawing = runs[1].Element(w + "drawing")!;
            drawing.Descendants(a + "xfrm").Single().Attribute("rot")!.Remove();
            foreach (var extent in drawing.Descendants().Where(node => node.Name == a + "ext" || node.Name == wp + "extent"))
            {
                extent.SetAttributeValue("cx", width * 12700);
                extent.SetAttributeValue("cy", 18 * 12700);
            }
            if (trailingSpace) runs[2].Element(w + "t")!.Value = " ";
            else runs[2].Remove();
            for (int index = 1; index < pictures; index++)
            {
                var sibling = new XElement(runs[1]);
                sibling.Descendants(wp + "docPr").Single().SetAttributeValue("id", index + 1);
                runs[1].AddAfterSelf(sibling);
            }
            if (merge) MergePictureRuns(paragraph, w);
            if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        });
}
