using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxAutofitParagraphInsetsKeepCompleteLabelInk()
    {
        foreach (bool fixedWidth in new[] { false, true })
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int gridHint in new[] { 2000, 9360 })
        foreach (bool split in new[] { false, true })
        foreach (var (left, first, right) in new[] { (0, 0, 0), (0, 72, 0), (36, 0, 0), (0, 0, 36), (12, 24, 18) })
        {
            var source = ReadAutofitInsetLabel(mode, fixedWidth, gridHint, left, first, right, split);
            var page = AssertFloatingLabelsPaint(source, mode, ["ALPHA"],
                $"autofit inset/{fixedWidth}/{left}/{first}/{right}/{split}", mathFont: mathFont);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.Equal("ALPHA", string.Concat(shows.Select(show => show.Text)).Trim());
            TestAssert.True(shows.All(show => Math.Abs(show.Y - shows[0].Y) < .001d),
                "A natural cell funds paragraph insets instead of wrapping a fitting short label.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => PictureRunBoundaryRenderer(mode, mathFont)
                .RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadAutofitInsetLabel(OoxPdfDocxMarkupMode mode, bool fixedWidth,
        int gridHint, int left, int first, int right, bool split) =>
        ReadMutation("docx-picture-rotated-raster", mode, false, false, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "p").Single();
            paragraph.ReplaceNodes(new XElement(w + "pPr", new XElement(w + "ind",
                new XAttribute(w + "left", left * 20), new XAttribute(w + "firstLine", first * 20),
                new XAttribute(w + "right", right * 20))),
                new XElement(w + "r", new XElement(w + "t", split ? "AL" : "ALPHA")));
            if (split) paragraph.Add(new XElement(w + "r", new XElement(w + "t", "PHA")));
            var properties = new XElement(w + "tblPr");
            if (fixedWidth) properties.Add(
                new XElement(w + "tblW", new XAttribute(w + "type", "dxa"), new XAttribute(w + "w", "9360")),
                new XElement(w + "tblLayout", new XAttribute(w + "type", "fixed")));
            paragraph.ReplaceWith(new XElement(w + "tbl", properties,
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", gridHint))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        });
}
