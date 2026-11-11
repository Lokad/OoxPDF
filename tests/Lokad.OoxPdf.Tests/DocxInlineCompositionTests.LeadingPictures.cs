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
        foreach (int? minimum in new int?[] { null, 16, 32 })
        {
            var source = ReadLeadingPictureWithLineFlow(mode, container, indent, split, minimum);
            var page = AssertFloatingLabelsPaint(source, mode, ["After"],
                $"leading picture/{container}/{indent}/{split}/minimum={minimum}", mathFont: mathFont);
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

    public static void DocxMinimumPictureRowsKeepDeclaredLineFrames()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (int minimum in new[] { 16, 32 })
        {
            var source = ReadLeadingPictureWithLineFlow(OoxPdfDocxMarkupMode.Final, container, 350, false, minimum);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(source).Single(p => p.Images.Count != 0);
            TestAssert.True(TestAssert.NotNull(paragraph.InlineContent).OfType<DocxInlinePicture>()
                .Any(item => ReferenceEquals(item.Image, paragraph.Images.Single())),
                "The reader retains ordinary picture source ownership before any mathematics or metadata is added.");
            var page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
            var lines = container == "body" ? page.Items.OfType<DocxTextLineLayout>().ToArray()
                : page.Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.ToArray();
            TestAssert.Equal(2, lines.Length);
            TestAssert.True(lines.All(line => line.InlineFrame is { } frame && frame.Height >= minimum - .0000001d),
                "Both the empty leading row and the picture/prose row retain the authored minimum.");
            TestAssert.True(lines[0].BaselineY > lines[1].BaselineY,
                "A picture fitting the continuation width advances past the authored first-line inset.");
        }
    }

    public static void DocxFixedPictureRowsKeepFollowingWordsAndPhysicalSlots()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool split in new[] { false, true })
        foreach (int height in new[] { 4, 8 })
        {
            var source = ReadLeadingPictureWithLineFlow(mode, container, 350, split, null,
                exact: 16, pictureHeight: height);
            var page = AssertFloatingLabelsPaint(source, mode, ["After"],
                $"fixed leading picture/{container}/{split}/{height}", mathFont: mathFont);
            var (bounds, clip) = ReadPictureBounds(page).Single();
            TestAssert.True(Math.Abs(bounds.Width - 240d) < .001d && Math.Abs(bounds.Height - height) < .001d,
                "A fitting fixed-slot picture keeps its authored dimensions after first-line relocation.");
            TestAssert.True(bounds.X >= clip.X - .001d && bounds.Y >= clip.Y - .001d &&
                bounds.X + bounds.Width <= clip.X + clip.Width + .001d &&
                bounds.Y + bounds.Height <= clip.Y + clip.Height + .001d,
                "The picture and following word survive their ancestor clips.");
        }
    }

    public static void DocxFixedPictureAdmissionUsesPhysicalSlot()
    {
        foreach (double scale in new[] { 1d, .75d })
        foreach (int pictureHeight in new[] { 4, 10 })
        {
            var source = ReadLeadingPictureWithLineFlow(OoxPdfDocxMarkupMode.Final, "cell", 350, false, null,
                exact: 16, pictureHeight: pictureHeight);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(source).Single(p => p.Images.Count != 0);
            double physicalInset = 8d * scale;
            bool admitted = DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, null,
                new DocxTests.FamilyWidthTextMeasurer(), 11d, 12.8d, 16d, scale, default, out var atoms,
                physicalTextBaselineInset: physicalInset);
            TestAssert.True((pictureHeight == 4) == admitted,
                "Containment uses the placed physical baseline, including scaled frames, rather than a nominal inset.");
            if (!admitted) continue;
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(TestAssert.NotNull(atoms),
                118d * scale, 468d * scale, default, out var rows),
                "The fitting source wraps into its first and continuation frames.");
            TestAssert.True(TestAssert.NotNull(rows).All(row => Math.Abs(row.OccupiedHeight - 16d * scale) < .0000001d &&
                Math.Abs(row.BaselineInset - physicalInset) < .0000001d),
                "Relocation preserves the fixed physical advance and baseline on every row.");
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
    private static DocxDocument ReadLeadingPictureWithLineFlow(OoxPdfDocxMarkupMode mode, string container,
        int indent, bool split, int? minimum, int? exact = null, int pictureHeight = 18) =>
        ReadMutation("docx-picture-rotated-raster", mode, false, false, (xml, w, _) =>
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
                extent.SetAttributeValue("cy", pictureHeight * 12700);
            }
            paragraph.AddFirst(new XElement(w + "pPr", new XElement(w + "ind",
                new XAttribute(w + "firstLine", indent * 20))));
            if ((minimum ?? exact) is { } height) paragraph.Element(w + "pPr")!.Add(new XElement(w + "spacing",
                new XAttribute(w + "line", height * 20), new XAttribute(w + "lineRule", exact is null ? "atLeast" : "exact")));
            if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        });
}
