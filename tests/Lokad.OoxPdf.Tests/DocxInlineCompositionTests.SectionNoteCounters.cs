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

    public static void DocxLaterSectionEndnoteUsesReferencePageAfterEarlierContinuation()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadSectionCounter(mode, 900, "PAGE");
            DocxParagraph firstReference = source.Paragraphs.Single(paragraph =>
                paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Endnote));
            var secondReference = firstReference with
            {
                InlineReferences = firstReference.InlineReferences.Select(reference => reference with { Id = "38" }).ToArray()
            };
            var secondNote = ReadSectionCounter(mode, 6, "PAGE").RelatedStories.Single(story => story.Id == "37") with { Id = "38" };
            source = source with
            {
                BodyElements = [.. source.BodyElements, new DocxParagraphElement(secondReference)],
                RelatedStories = [.. source.RelatedStories, secondNote]
            };

            var pages = PictureRangeRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable()
                .Select(DocxPageFootnoteNumberingTests.DecodeText).ToArray();
            int laterSection = Array.FindIndex(pages, text => text.Contains("Public following section", StringComparison.Ordinal));
            TestAssert.True(laterSection >= 4, "The earlier note must insert a continuation before the later reference.");
            TestAssert.Contains("Count 2", string.Concat(pages.Take(laterSection)));
            TestAssert.Contains("Count " + (laterSection + 1).ToString(CultureInfo.InvariantCulture), pages[laterSection]);
        }
    }

    public static void DocxSectionPageCountersAccountForInsertedEndnotePages()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = PictureRangeRenderer(mode);
            var source = WithNoteMathStaticStory(ReadSectionCounter(mode, 900, null), true,
                CounterParagraph("Public footer ", includePageCount: true));
            var body = source.BodyElements.Take(source.BodyElements.Count - 1).ToArray();
            var boundary = (DocxSectionBreakElement)body[^1];
            body[^1] = boundary with { PageSettings = boundary.PageSettings with { DecimalPageNumberStart = 25 } };
            source = source with { BodyElements = [.. body, new DocxParagraphElement(CounterParagraph("Public following section "))] };
            foreach (int? restart in new int?[] { 40, null })
            {
                var document = source with { PageSettings = source.PageSettings with { DecimalPageNumberStart = restart } };
                var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                var text = pages.Select(SectionCounterPageText).ToArray();
                int following = Array.FindIndex(text, page => page.Contains("Publicfollowingsection", StringComparison.Ordinal));
                TestAssert.True(following >= 4, "A late endnote continuation precedes the next section.");
                TestAssert.Contains("item899", string.Concat(text.Take(following)));
                TestAssert.Contains($"Publicfollowingsection{restart ?? 25 + following}", text[following]);
                for (int index = 0; index < pages.Length; index++)
                {
                    int number = restart is { } start && index >= following
                        ? start + index - following : 25 + index;
                    TestAssert.Contains($"Publicfooter{number}of{pages.Length}", text[index]);
                }
            }
        }

        static DocxParagraph CounterParagraph(string label, bool includePageCount = false)
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph(label, 12d, 14d);
            var style = paragraph.Runs.Single();
            var page = style with { Text = "{PAGE}", FieldKind = DocxFieldKind.Page };
            return paragraph with { Runs = includePageCount
                ? [style, page, style with { Text = " of " }, style with { Text = "{NUMPAGES}", FieldKind = DocxFieldKind.NumPages }]
                : [style, page] };
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
