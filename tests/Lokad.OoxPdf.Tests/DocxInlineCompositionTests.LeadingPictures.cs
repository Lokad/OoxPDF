using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxLeadingInlinePicturesKeepFollowingWordInk()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int indent in new[] { 0, 350 })
        foreach (bool split in new[] { false, true })
        {
            var source = ReadMutation("docx-picture-rotated-raster", mode, false, false, (xml, w, _) =>
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                var paragraph = xml.Descendants(w + "p").Single();
                var runs = paragraph.Elements(w + "r").ToArray();
                runs[0].Remove();
                runs[2].Element(w + "t")!.Value = split ? " Af" : " After";
                runs[2].Element(w + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                if (split) paragraph.Add(new XElement(w + "r", new XElement(w + "t", "ter")));
                var drawing = runs[1].Element(w + "drawing")!;
                drawing.Descendants(a + "xfrm").Single().Attribute("rot")!.Remove();
                foreach (var extent in drawing.Descendants().Where(node => node.Name == a + "ext" || node.Name == wp + "extent"))
                {
                    extent.SetAttributeValue("cx", 240 * 12700);
                    extent.SetAttributeValue("cy", 18 * 12700);
                }
                paragraph.AddFirst(new XElement(w + "pPr", new XElement(w + "ind",
                    new XAttribute(w + "firstLine", indent * 20))));
                if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                    new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))),
                    new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
            });
            var page = AssertFloatingLabelsPaint(source, mode, ["After"],
                $"leading picture/{container}/{indent}/{split}", mathFont: mathFont);
            var (bounds, clip) = ReadPictureBounds(page).Single();
            TestAssert.True(bounds.X >= clip.X - .001d && bounds.Y >= clip.Y - .001d &&
                bounds.X + bounds.Width <= clip.X + clip.Width + .001d &&
                bounds.Y + bounds.Height <= clip.Y + clip.Height + .001d,
                "A leading picture and its following word survive their actual ancestor clips.");
            TestAssert.True(Math.Abs(bounds.Width - 240d) < .001d && Math.Abs(bounds.Height - 18d) < .001d,
                "A picture fitting the continuation frame keeps its authored dimensions.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => PictureRunBoundaryRenderer(mode, mathFont)
                .RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxReviewPictureOwnerDescentKeepsFollowingParagraphClear()
    {
        foreach (bool mathFont in new[] { false, true })
        {
            const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.AllMarkup;
            var baselines = new List<double>();
            var pictures = new List<Lokad.OoxPdf.Pdf.PdfRectangle>();
            double expectedGrowth = 0d;
            foreach (int ownerSize in new[] { 11, 18 })
            {
                var source = ReadMutation("PictureRevisions/picture-revision-calibri18", mode, false, false, (xml, w, _) =>
                {
                    var owner = xml.Descendants(w + "drawing").Single().Ancestors(w + "r").First();
                    owner.Element(w + "rPr")!.Element(w + "sz")!.SetAttributeValue(w + "val", ownerSize * 2);
                });
                var renderer = PictureRunBoundaryRenderer(mode, mathFont);
                var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
                pictures.Add(ReadPictureRangeFrame(page));
                var following = renderer.InspectTextEmission(source).Lines.Where(line => line.SourceBlockIndex == 1)
                    .SelectMany(line => line.Segments).First();
                baselines.Add(following.BaselineY);
                var font = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).First().Font.Font;
                // The glyphless source run reserves its own font descent. Read
                // the pinned font program independently; include the renderer's
                // actual printed text scale in Review rather than fitting a delta.
                expectedGrowth = (18d - 11d) * font.Os2.WindowsDescender / font.UnitsPerEm * following.FontSize / 11d;
            }
            TestAssert.True(expectedGrowth > 0d && Math.Abs(baselines[0] - baselines[1] - expectedGrowth) < .001d,
                $"The following paragraph clears source-run descent after the physical text strut: {mode}/{mathFont}, " +
                $"baselines={string.Join(',', baselines)}, expected growth={expectedGrowth}.");
            TestAssert.Equal(pictures[0], pictures[1]);
        }
    }
}
