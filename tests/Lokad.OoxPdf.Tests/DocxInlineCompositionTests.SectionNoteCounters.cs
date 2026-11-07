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
            var literal = renderer.RenderPagesAsync(ReadSectionCounter(null), null, default).ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(ReadSectionCounter("PAGE"), null, default).ToBlockingEnumerable().ToArray();
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

            DocxDocument ReadSectionCounter(string? instruction)
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
                return source with { BodyElements =
                [
                    new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                    new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                    .. source.BodyElements,
                    new DocxSectionBreakElement(settings, DocxSectionBreakType.NextPage, null, null, null, []),
                    new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public following section", 12d, 14d))
                ] };
            }
        }
    }
}
