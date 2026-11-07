using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxHyperlinksPreserveCompleteEndnoteStoryPlacement()
    {
        foreach (string fixture in new[] { "mixed-note-endnote-paragraph", "mixed-note-endnote-table", "mixed-nested-math-endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            CheckNavigationSpacing("CompositionProbes/" + fixture, mode, "word/endnotes.xml");
    }

    public static void DocxHyperlinksPreserveContextualBodyAndMinimumCellSpacing()
    {
        foreach (string fixture in new[] { "note-contextual-spacing-footnote", "note-contextual-spacing-endnote",
            "table-minimum-line-spacing-row", "table-minimum-line-spacing-grid" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            CheckNavigationSpacing(fixture, mode, null);
    }

    private static void CheckNavigationSpacing(string fixture, OoxPdfDocxMarkupMode mode, string? notePart)
    {
        var renderer = PictureRangeRenderer(mode);
        var control = ReadMutation(fixture, mode, false, false, partOverride: notePart);
        var source = ReadMutation(fixture, mode, false, false, (xml, w, m) =>
        {
            var paragraph = notePart is null ? xml.Descendants(w + "p").First() :
                xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
            var target = notePart is null ? paragraph.Elements(w + "r").First() : paragraph.Elements(m + "oMath").First();
            target.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "990"), new XAttribute(w + "name", "SpacingTarget")));
            target.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "990")));
            var content = notePart is null ? target : paragraph.Elements(m + "oMath").Last();
            var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "SpacingTarget"));
            content.ReplaceWith(link);
            link.Add(content);
        }, partOverride: notePart);
        var before = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().ToArray();
        var after = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(before.Length, after.Length);
        int added = 0;
        for (int index = 0; index < before.Length; index++)
        {
            var first = before[index]; var second = after[index];
            // Compare emitted geometry/content, independent of PDF object IDs or
            // incidental operation grouping. Navigation adds no painted payload.
            TestAssert.True(Glyphs(first).SequenceEqual(Glyphs(second)),
                "A navigation range preserves actual glyph origins and sizes: " + fixture + "/" + mode);
            TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(first), DocxPageFootnoteNumberingTests.DecodeText(second));
            TestAssert.Equal(first.Images.Count, second.Images.Count);
            if (first.Images.Count != 0) TestAssert.Equal(ReadPictureRangeFrame(first), ReadPictureRangeFrame(second));
            foreach (var link in first.Annotations)
                TestAssert.True(second.Annotations.Contains(link), "Existing note targets and hit areas retain their placed ownership.");
            added += second.Annotations.Count - first.Annotations.Count;
        }
        TestAssert.True(added > 0, "The new hyperlink is emitted rather than ignored.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
            .ToBlockingEnumerable().ToArray());

        static IEnumerable<(string Text, double X, double Y, double Size)> Glyphs(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Select(g => (g.Text, g.X, g.Y, g.Size));
    }
}
