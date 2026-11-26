using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxPictureOnlyNotesShareAuthoredAspectRatioFitting()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (double width in new[] { .25d, .5d, 1d })
        {
            var source = ReadMutation("CompositionProbes/mixed-note-" + kind + "-paragraph", mode, false, false,
                (xml, w, _) =>
                {
                    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                    XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(w + "drawing").Any());
                    var drawing = new XElement(paragraph.Descendants(w + "drawing").Single());
                    foreach (var extent in drawing.Descendants().Where(node => node.Name == a + "ext" || node.Name == wp + "extent"))
                    { extent.SetAttributeValue("cx", (long)(width * 12700d)); extent.SetAttributeValue("cy", 18 * 12700); }
                    paragraph.ReplaceNodes(paragraph.Element(w + "pPr"), new XElement(w + "r", drawing));
                }, partOverride: "word/" + kind + "s.xml");
            var renderer = PictureRunBoundaryRenderer(mode, mathFont: true);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var pictures = pages.SelectMany(page => Regex.Matches(page.Content,
                @"(?<Width>[-\d.]+) 0 0 (?<Height>[-\d.]+) [-\d.]+ [-\d.]+ cm\s+/Im\d+ Do"));
            double Number(Match match, string group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
            var picture = pictures.Single(p => Math.Abs(Number(p, "Width") - width) < .001d);
            TestAssert.True(Math.Abs(Number(picture, "Height") - 18d) < .001d,
                "Picture-only note fallback shares the authored aspect ratio: " + kind + "/" + mode + "/" + width);
            string text = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                honorActualText: false)).Select(show => show.Text));
            TestAssert.Contains("Alpha", text);
            TestAssert.Contains("Delta", text);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null,
                canceled.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxUnpositionedStaticPicturesShareBoundedFitAndOwners()
    {
        foreach (bool header in new[] { false, true })
        foreach (double width in new[] { 0d, .5d, 1d, 36d, 600d })
        {
            var source = ReadPictureFit("body", OoxPdfDocxMarkupMode.Final, width, fixedRow: false);
            var original = source.Paragraphs.Single();
            var image = original.Images.Single() with { SourceRunIndex = -1 };
            var paragraph = original with { Runs = [], Images = [image], InlineContent = null };
            var elements = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxParagraphElement(paragraph)] };
            var settings = header ? source.PageSettings with { HeaderBodyElementsByType = elements }
                : source.PageSettings with { FooterBodyElementsByType = elements };
            var body = DocxTests.CreateDocxLayoutParagraph("Body survives.", 12d, 14d);
            source = source with { PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings },
                BodyElements = [new DocxParagraphElement(body)], HeaderParagraphs = header ? [paragraph] : [], FooterParagraphs = header ? [] : [paragraph],
                HeaderBodyElementsByType = settings.HeaderBodyElementsByType, FooterBodyElementsByType = settings.FooterBodyElementsByType };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var page = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
            var placed = page.StaticInlineImages.Single();
            double ratio = width == 0d ? 1d : Math.Min(1d, (page.Width - page.MarginLeft - page.MarginRight) / width);
            TestAssert.True(ReferenceEquals(image, placed.Image) && ReferenceEquals(paragraph, placed.SourceParagraph),
                "Legacy unpositioned fallback retains its actual image and paragraph owners.");
            TestAssert.True(Math.Abs(placed.Width - width * ratio) < 1e-9d && Math.Abs(placed.Height - 18d * ratio) < 1e-9d,
                "Static fallback shares zero-width height preservation and bounded aspect ratios.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), canceled.Token));
        }
    }
}
