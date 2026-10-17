using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxReviewEmptyRunOwnersKeepActualSourceMetadata()
    {
        using var golden = ReadEmptyRevisionGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadNonTextAuthorSource(row, mode);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First();
            TestAssert.Equal(row.GetProperty("ExpectedEmptyOwners").GetInt32(), paragraph.NonRenderingRunOwners.Count);
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Actual empty-run ownership must preserve supported paragraph composition.");
            foreach (var owner in paragraph.NonRenderingRunOwners)
            {
                TestAssert.True(owner.SourceRunIndex >= 0 && owner.Revisions.Count != 0 && owner.Revisions.All(revision =>
                    paragraph.Revisions.Any(item => ReferenceEquals(item, revision))),
                    "Empty ownership must retain the actual source index and revision identities.");
                TestAssert.True(paragraph.Runs.All(run => run.SourceRunIndex != owner.SourceRunIndex),
                    "Metadata cannot invent a visible text run or font purpose for the empty source.");
            }
            if (paragraph.NonRenderingRunOwners.Count == 0) continue;
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default,
                paragraph.Revisions.Select(revision => revision with { }).ToArray()),
                "Value-equal aggregate metadata must not invent empty source ownership.");
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { NonRenderingRunOwners = [] }, default),
                "Removing the actual owner must preserve the conservative unsupported fallback.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
        }
    }

    public static void DocxReviewEmptyRunOwnersComposeWithMathematics()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native review controls use installed Windows fonts.");
        CheckEmptyRevisionOwners(null);
    }

    internal static void CheckEmptyRevisionOwners(IFontResolver? resolver)
    {
        using var golden = ReadEmptyRevisionGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadNonTextAuthorSource(row, mode);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First();
            int equations = row.GetProperty("ExpectedEquationCount").GetInt32();
            TestAssert.Equal(equations, paragraph.Equations.Count);
            var runs = paragraph.Runs.ToArray();
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = ReadEmbeddedGlyphTextShows(page).ToArray();
            string text = string.Concat(shows.Select(show => show.Text));
            foreach (var expected in row.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.Contains(expected.GetString()!, text);
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => mathFonts.Contains(show.Font)).Select(show => show.Text).ToArray();
            if (equations != 0)
            {
                TestAssert.True(math.Contains("1") && math.Contains("2"),
                    "Both source fraction operands must be painted beside the empty revision in both public presets: " + row.GetProperty("Id").GetString() + "/" + mode);
            }
            else TestAssert.Equal(0, math.Length);
            if (mode == OoxPdfDocxMarkupMode.AllMarkup)
            {
                var fill = ReadNonTextAuthorRgb(row, "FillRgb");
                TestAssert.Contains(DocxTests.FormatPdfRgb(fill, "rg"), page.Content);
                var comment = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "0");
                var slots = DocxRenderer.BuildWordReviewAuthorPalette(document, [("0", comment.CommentMetadata!.Author)], default);
                TestAssert.Equal(fill, DocxRenderer.ResolveCommentAuthorFillSnapshot(slots.CommentSlots, "0"));
            }
            TestAssert.True(runs.SequenceEqual(paragraph.Runs, ReferenceEqualityComparer.Instance),
                "Author registration cannot replace visible source runs or add layout content.");
        }
    }

    private static JsonDocument ReadEmptyRevisionGolden() => JsonDocument.Parse(File.ReadAllText(
        Path.Combine(NonTextAuthorCases(), "empty-revision-owners-office.json")));
}
