using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSectionEndnotePageCountersPreserveReferenceAndFollowingSection()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int count in new[] { 80, 900 })
        {
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(ReadSectionCounter(mode, count, null), null, default).ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(ReadSectionCounter(mode, count, "PAGE"), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(counter.Length >= 3, "The reference and following section occupy distinct pages.");
            TestAssert.Equal(literal.Length, counter.Length);
            TestAssert.Contains("Count 2", string.Concat(counter.Select(DocxPageFootnoteNumberingTests.DecodeText)));
            TestAssert.Contains("Public following section", DocxPageFootnoteNumberingTests.DecodeText(counter[^1]));
            TestAssert.True(!DocxPageFootnoteNumberingTests.DecodeText(counter[^1]).Contains("Count ", StringComparison.Ordinal),
                "Section-owned note continuations precede the following body section.");
            for (int index = 0; index < literal.Length; index++)
            {
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(literal[index], usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(counter[index], usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                TestAssert.True(a.SequenceEqual(b), "Section-end PAGE preserves literal glyphs and pagination: " + mode + "/" + count);
                TestAssert.True(literal[index].Annotations.SequenceEqual(counter[index].Annotations),
                    "The reference and following-section destinations retain their placement.");
            }


        }
    }
    public static void DocxEndnotePlacementUsesDocumentPositionAndPreservesSectionMetadata()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var positions in new (string? Global, string? First, string? Final)[]
            { (null, "sectEnd", "sectEnd"), ("docEnd", "sectEnd", "sectEnd"),
              ("sectEnd", "sectEnd", "docEnd"), ("sectEnd", "docEnd", "docEnd"), ("sectEnd", null, null) })
        {
            var source = ReadSectionCounter(mode, 80, null);
            source = source with
            {
                Settings = source.Settings with { EndnoteReferenceSettings =
                    source.Settings.EndnoteReferenceSettings with { PositionValue = positions.Global } },
                PageSettings = source.PageSettings with { EndnoteReferenceSettings =
                    source.PageSettings.EndnoteReferenceSettings with { PositionValue = positions.Final } },
                BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
                    ? section with { PageSettings = section.PageSettings with { EndnoteReferenceSettings =
                        section.PageSettings.EndnoteReferenceSettings with { PositionValue = positions.First } } } : element).ToArray()
            };
            var pages = PictureRangeRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable()
                .Select(DocxPageFootnoteNumberingTests.DecodeText).ToArray();
            int notePage = Array.FindIndex(pages, text => text.Contains("Public section note", StringComparison.Ordinal));
            int followingPage = Array.FindIndex(pages, text => text.Contains("Public following section", StringComparison.Ordinal));
            TestAssert.True(notePage >= 0 && followingPage >= 0, "Both note and following body content survive.");
            TestAssert.True(positions.Global == "sectEnd" ? notePage < followingPage : notePage == followingPage,
                "Document-wide placement owns the boundary independently of authored section positions: " + positions);
            TestAssert.Equal(positions.First, source.BodyElements.OfType<DocxSectionBreakElement>().Single()
                .PageSettings.EndnoteReferenceSettings.PositionValue);
            TestAssert.Equal(positions.Final, source.PageSettings.EndnoteReferenceSettings.PositionValue);
        }
    }

    private static DocxDocument ReadSectionCounter(OoxPdfDocxMarkupMode mode, int count, string? instruction)
    {
        var source = ReadNoteCounter("endnote", mode, instruction, instruction is null ? "2" : "7", false,
            (xml, w) =>
            {
                var note = xml.Root!.Elements(w + "endnote").Single(n => (string?)n.Attribute(w + "id") == "37");
                var first = note.Elements(w + "p").First();
                foreach (var element in note.Elements().Where(e => !ReferenceEquals(e, first)).ToArray()) element.Remove();
                first.Descendants(w + "t").First().Value = " Public section note " +
                    string.Join(" ", Enumerable.Range(0, count).Select(n => "item" + n.ToString("D3", CultureInfo.InvariantCulture)));
            });
        var settings = source.PageSettings with { EndnoteReferenceSettings =
            source.PageSettings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } };
        return source with { Settings = source.Settings with { EndnoteReferenceSettings =
            source.Settings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } }, BodyElements =
        [
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
            new DocxParagraphElement(source.Paragraphs.First()),
            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public trailing body", 12d, 14d)),
            new DocxSectionBreakElement(settings, DocxSectionBreakType.NextPage, null, null, null, []),
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public following section", 12d, 14d))
        ], RelatedStories = source.RelatedStories.Where(story => !story.IsNormalStoryType || story.Id == "37").ToArray() };
    }
}
