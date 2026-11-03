using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxMixedTerminalParagraphsKeepBreaksAndPhysicalFacesAcrossCounts()
    {
        foreach (int count in new[] { 2, 4, 5, 6, 7, 9 })
        foreach (string wrapping in new[] { "none", "leading", "terminal" })
        {
            var paragraphs = Enumerable.Range(0, count).Select(index =>
            {
                bool terminal = index == count - 1;
                string prefix = terminal ? "Closing." : $"P{index + 1}.";
                if (wrapping == "terminal" && terminal || wrapping == "leading" && index == count / 2 && !terminal)
                    prefix += " " + new string('m', 72) + " ending.";
                var paragraph = DocxTests.CreateDocxLayoutParagraph(prefix, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace")],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
                return terminal ? paragraph with
                {
                    Runs = [paragraph.Runs[0], paragraph.Runs[0] with { Text = " tail.", FontFamily = "OtherCommentFace" }]
                } : paragraph;
            }).ToArray();
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var page = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var shows = ReadEmbeddedGlyphTextShows(page)
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(p => string.Concat(p.Runs.Select(r => r.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            var starts = Enumerable.Range(0, count).Select(index => shows.First(show =>
                show.Text.StartsWith(index == count - 1 ? "Closing." : $"P{index + 1}.", StringComparison.Ordinal))).ToArray();
            TestAssert.True(starts.Zip(starts.Skip(1)).All(pair => pair.First.Y > pair.Second.Y),
                "Every authored paragraph begins on a later row, independently of count and wrapping.");
            TestAssert.True(shows.Any(show => show.Text.TrimStart().StartsWith("tail.", StringComparison.Ordinal) &&
                show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 900),
                "The closing run retains its own physical face instead of the paragraph-preview fallback.");
        }
    }

    private static (string Text, double Size, double X, double Y, double Advance)[] PaintedBalloonText(
        DocxRenderer renderer, DocxDocument document) => renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable()
            .SelectMany(ReadEmbeddedGlyphTextShows)
            .Select(show => (show.Text, show.Size, show.X, show.Y, show.Font.MeasureTextPoints(show.Text, show.Size))).ToArray();
}
