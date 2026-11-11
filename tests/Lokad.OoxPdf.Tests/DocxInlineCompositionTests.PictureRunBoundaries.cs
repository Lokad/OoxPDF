using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrdinaryPicturesKeepPaintWhenFormattedRunsAreMerged()
    {
        foreach (string fixture in new[] { "docx-picture-rotated-raster", "docx-svg-rotated-inline" })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = PictureRunBoundaryRenderer(mode, mathFont);
            var control = renderer.RenderPagesAsync(ReadMutation(fixture, mode, false, false), null, default)
                .ToBlockingEnumerable().Single();
            var source = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
                MergePictureRuns(xml.Descendants(w + "p").Single(), w));
            var paragraph = source.Paragraphs.Single(item => item.Images.Count != 0);
            var image = paragraph.Images.Single();
            TestAssert.Equal(7, image.SourceTextOffsetInRun);
            TestAssert.True(paragraph.Runs.Any(run => run.SourceRunIndex == image.SourceRunIndex &&
                run.SourceTextOffsetInRun == 7 && run.Text == " After"),
                "Text following a drawing keeps its position inside the source run.");
            var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(control.Content, page.Content);
            TestAssert.Equal(control.Images.Count, page.Images.Count);
            TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control), DocxPageFootnoteNumberingTests.DecodeText(page));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null,
                cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxOrdinaryPicturesKeepSourceOrderWithSeveralDrawingsInOneRun()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Read(bool merge) => ReadMutation("docx-picture-rotated-raster", mode, false, false,
                (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single();
                    var drawing = new XElement(paragraph.Descendants(w + "drawing").Single());
                    drawing.Descendants(a + "xfrm").Single().Attribute("rot")!.Remove();
                    var extent = drawing.Descendants(a + "ext").Single();
                    extent.SetAttributeValue("cx", "914400");
                    extent.SetAttributeValue("cy", "1828800");
                    paragraph.ReplaceNodes(Text("Before "), new XElement(w + "r", drawing), Text(" Between "),
                        new XElement(w + "r", new XElement(drawing)), Text(" After"));
                    if (merge) MergePictureRuns(paragraph, w);
                    XElement Text(string text) => new(w + "r", new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), text));
                });
            var renderer = PictureRunBoundaryRenderer(mode, mathFont);
            var control = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().Single();
            var page = renderer.RenderPagesAsync(Read(true), null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(control.Content, page.Content);
            var images = Regex.Matches(page.Content,
                @"(?<Width>[-\d.]+) 0 0 (?<Height>[-\d.]+) (?<X>[-\d.]+) (?<Y>[-\d.]+) cm\s+/Im\d+ Do")
                .Select(match => (X: Value(match, "X"), Right: Value(match, "X") + Value(match, "Width"))).ToArray();
            TestAssert.Equal(2, images.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            var before = shows.Single(show => show.Text.Trim() == "Before");
            var between = shows.Single(show => show.Text.Trim() == "Between");
            var after = shows.Single(show => show.Text.Trim() == "After");
            TestAssert.True(before.X < images[0].X && between.X >= images[0].Right - .001d &&
                between.X < images[1].X && after.X >= images[1].Right - .001d,
                "Text and both pictures retain their visible source order across run boundaries.");
        }

        static double Value(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    }

    private static DocxRenderer PictureRunBoundaryRenderer(OoxPdfDocxMarkupMode mode, bool mathFont) =>
        new(new PortableRenderingTests.PinnedResolver(mathFont), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    private static void MergePictureRuns(XElement paragraph, XNamespace w)
    {
        var runs = paragraph.Elements(w + "r").ToArray();
        TestAssert.True(runs.Length > 1 && runs.All(run => run.Element(w + "rPr") is null),
            "The source-preserving transformation merges equally formatted runs.");
        var merged = new XElement(w + "r", runs.SelectMany(run => run.Elements()).Select(child => new XElement(child)));
        runs[0].AddBeforeSelf(merged);
        foreach (var run in runs) run.Remove();
    }
}
