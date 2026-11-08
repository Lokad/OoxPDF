using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCommentAnchorTests
{
    public static void DocxCommentEndpointBeforeContentKeepsNominalPageAndSource()
    {
        var paragraph = Point(after: false);
        var source = Document(paragraph);
        var renderer = Renderer();
        TestAssert.Equal(1d, Scale(source));
        TestAssert.Equal(0d, renderer.InspectLayout(source).MarkupMarginReservePoints);
        TestAssert.Equal(0, renderer.InspectMarkupBalloons(source).Count);
        string text = string.Concat(renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable()
            .Select(DocxPageFootnoteNumberingTests.DecodeText));
        TestAssert.Contains("Tail", text);
        TestAssert.True(!text.Contains("Public comment", StringComparison.Ordinal),
            "An endpoint before source content does not paint a balloon or reserve its page lane.");
        TestAssert.Equal(1, paragraph.InlineReferences.Count);
        TestAssert.Equal(1, paragraph.CommentRanges.Count);
        TestAssert.Equal(4, paragraph.InlineReferences[0].TextOffsetInRun);
        TestAssert.Equal(0, paragraph.CommentRanges[0].EndInlineItemIndex!.Value);
        TestAssert.Equal(1, Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectMarkupBalloons(source).Count);

        var after = Document(Point(after: true));
        TestAssert.True(Scale(after) < 1d, "A collapsed endpoint after source content retains its printed comment.");
        TestAssert.Equal(1, renderer.InspectMarkupBalloons(after).Count);
    }

    public static void DocxCommentEndpointsSharePerReferencePaintingAndLaneOwnership()
    {
        var before = Point(after: false);
        var after = Point(after: true);
        var paragraph = before with
        {
            InlineReferences = [before.InlineReferences[0], after.InlineReferences[0] with { Id = "1" }],
            CommentRanges = [before.CommentRanges[0], after.CommentRanges[0] with { Id = "1" }]
        };
        var source = Document(paragraph);
        TestAssert.True(Scale(source) < 1d, "The surviving source anchor owns the shared printed lane.");
        var renderer = Renderer();
        TestAssert.Equal(1, renderer.InspectMarkupBalloons(source).Count);
        string text = string.Concat(renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable()
            .Select(DocxPageFootnoteNumberingTests.DecodeText));
        TestAssert.True(!text.Contains("Public comment", StringComparison.Ordinal),
            "A surviving neighbor cannot make the empty source anchor paint a balloon.");
        TestAssert.Contains("Second comment", text);
        TestAssert.Equal(2, paragraph.CommentRanges.Count);
        TestAssert.Equal(2, Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectMarkupBalloons(source)
            .Sum(placement => placement.CommentCandidateCount));
    }

    public static void DocxCommentUnknownEndpointsRetainTheirEstablishedFallback()
    {
        var source = Point(after: false);
        foreach (var paragraph in new[]
        {
            source with { InlineContent = null },
            source with { CommentRanges = [] },
            source with { CommentRanges = [source.CommentRanges[0] with { StartInlineItemIndex = null }] },
            source with { FieldReferences = [new(DocxFieldKind.Page, DocxFieldSourceKind.ComplexInstruction,
                "PAGE", null, 0, 0, 0, 0)] },
            source with { InlineReferences = [source.InlineReferences[0],
                new(DocxRelatedStoryKind.Footnote, "2", null, null, 0, 0, 0)] }
        })
            TestAssert.True(Scale(Document(paragraph)) < 1d,
                "Unresolved source views, cross-paragraph ranges, fields and note owners retain their prior print-profile fallback.");
    }

    private static DocxParagraph Point(bool after)
    {
        var paragraph = DocxTests.CreateCommentMarkerParagraph("Tail", "0");
        var run = paragraph.Runs[0] with { SourceRunIndex = 0 };
        return paragraph with
        {
            Runs = [run],
            InlineReferences = [paragraph.InlineReferences[0] with { TextOffsetInRun = 4 }],
            CommentRanges = [new("0", 0, after ? 4 : 0, 0, after ? 4 : 0, 0, 4)
                { StartInlineItemIndex = after ? 1 : 0, EndInlineItemIndex = after ? 1 : 0 }],
            InlineContent = [new DocxInlineText(run, 0, 4)]
        };
    }

    private static DocxDocument Document(DocxParagraph paragraph) =>
        DocxTests.CreateAllMarkupWrapProbeDocument([paragraph]) with
        {
            PageWidthPoints = 612d, PageHeightPoints = 792d,
            RelatedStories = [Story("0", "Public comment"), Story("1", "Second comment")]
        };
    private static DocxRelatedStory Story(string id, string text) => new(DocxRelatedStoryKind.Comment,
        "/word/comments.xml", id, [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d))], [], [], null);
    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) =>
        new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
    private static double Scale(DocxDocument source) => DocxRenderer.ResolveWordCompatiblePrintScale(source,
        DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
}
