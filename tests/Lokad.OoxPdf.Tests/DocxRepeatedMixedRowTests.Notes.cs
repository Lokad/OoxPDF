using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxRepeatedMixedRowTests
{
    public static void DocxNoteTablesSharePrintedFramesWithSurroundingNotes()
    {
        const string caption = "Public note table caption before the pictured calculation.";
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool mathFont in new[] { false, true })
        foreach (string alignment in new[] { "top", "bottom" })
        {
            double? originalPictureRatio = null;
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-" + kind + "-table",
                    mode, false, false, (xml, w, _) =>
                    {
                        var cell = xml.Descendants(w + "tc").Single();
                        cell.Element(w + "tcPr")!.Element(w + "vAlign")!.SetAttributeValue(w + "val", alignment);
                        var mixed = cell.Element(w + "p")!;
                        var paragraph = new XElement(w + "p");
                        if (mixed.Element(w + "pPr") is { } properties) paragraph.Add(new XElement(properties));
                        var run = new XElement(w + "r");
                        if (mixed.Element(w + "r")?.Element(w + "rPr") is { } style) run.Add(new XElement(style));
                        run.Add(new XElement(w + "t", caption)); paragraph.Add(run); mixed.AddBeforeSelf(paragraph);
                        cell.Parent!.AddFirst(new XElement(w + "trPr", new XElement(w + "cantSplit"),
                            new XElement(w + "trHeight", new XAttribute(w + "val", "1440"), new XAttribute(w + "hRule", "atLeast"))));
                    }, partOverride: "word/" + kind + "s.xml");
                // A body comment activates the existing printed Review frame.
                // The original note references and their source indices stay intact.
                source = source with
                {
                    BodyElements = [.. source.BodyElements, new DocxParagraphElement(DocxTests.CreateCommentRangeParagraph("Public review", "900"))],
                    RelatedStories = [.. source.RelatedStories, new DocxRelatedStory(DocxRelatedStoryKind.Comment,
                        "/word/comments.xml", "900", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public review", 10d, 12d))], [], [], null)]
                };
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                var shows = pages.Select(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true, includeTextTransforms: true).ToArray()).ToArray();
                var alpha = Find("Public destination Alpha.");
                var label = Find(caption);
                var bravo = Find("Public destination Bravo.");
                TestAssert.Equal(alpha.Page, label.Page);
                TestAssert.True(alpha.Y > label.Y && (bravo.Page > label.Page || bravo.Page == label.Page && label.Y > bravo.Y),
                    "The complete Alpha table stays after its introduction and before the next source note.");
                DocxInlineCompositionTests.AssertSourceLabelsPaint(pages[label.Page], mode, [caption],
                    $"note table/{kind}/{mathFont}/{alignment}/{mode}");
                var picture = DocxInlineCompositionTests.ReadPictureBounds(pages[label.Page]).Single().Bounds;
                double ratio = picture.Width / label.Size;
                originalPictureRatio ??= ratio;
                TestAssert.True(Math.Abs(ratio - originalPictureRatio.Value) < .001d,
                    "The picture and its caption share the printed story scale (PDF number roundoff only).");
                TestAssert.Equal(2, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                    .Normalize(System.Text.NormalizationForm.FormKC)
                    .Split("ratetime", StringSplitOptions.None).Length - 1);

                (int Page, double Y, double Size) Find(string text)
                {
                    for (int page = 0; page < shows.Length; page++)
                    {
                        string all = string.Concat(shows[page].Select(show => show.Text));
                        int start = all.IndexOf(text, StringComparison.Ordinal);
                        if (start < 0) continue;
                        TestAssert.Equal(start, all.LastIndexOf(text, StringComparison.Ordinal));
                        int offset = 0;
                        foreach (var show in shows[page])
                        {
                            if (start >= offset && start < offset + show.Text.Length) return (page, show.Y, show.Size);
                            offset += show.Text.Length;
                        }
                    }
                    throw new InvalidOperationException("Source note label was not painted: " + text);
                }
            }
        }
    }
}
