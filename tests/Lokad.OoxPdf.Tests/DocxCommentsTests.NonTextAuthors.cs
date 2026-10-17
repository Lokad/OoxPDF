using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxReviewNonTextAuthorsMatchIndependentOwnerOrder()
    {
        using var golden = ReadNonTextAuthorGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (!row.GetProperty("QualifiesRetainedOwners").GetBoolean()) continue;
            var document = ReadNonTextAuthorSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var comment = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "0");
            var slots = DocxRenderer.BuildWordReviewCommentAuthorPaletteSlots(document, [("0", comment.CommentMetadata!.Author)], default);
            TestAssert.Equal(ReadNonTextAuthorRgb(row, "FillRgb"), DocxRenderer.ResolveCommentAuthorFillSnapshot(slots, "0"));
        }
    }

    public static void DocxReviewNonTextAuthorsRejectUnownedParagraphMetadata()
    {
        using var golden = ReadNonTextAuthorGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            row.GetProperty("Kind").GetString() is "numbered-mark" or "paragraph-mark" or "run-format" or "image-insertion"))
        {
            var source = ReadNonTextAuthorSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var owner = DocxBlockTraversal.EnumerateBodyParagraphs(source).First();
            var unowned = owner with
            {
                Runs = owner.Runs.Select(run => run with { Revision = null, Revisions = [] }).ToArray(),
                Images = owner.Images.Select(image => image with { Revisions = [] }).ToArray(),
                ParagraphMarkRun = owner.ParagraphMarkRun is { } mark ? mark with { Revision = null, Revisions = [] } : null
            };
            TestAssert.True(unowned.Revisions.Count != 0, "The control must retain aggregate metadata after its actual owners are removed.");
            var changed = source with { BodyElements = source.BodyElements.Select(element =>
                element is DocxParagraphElement paragraph && ReferenceEquals(paragraph.Paragraph, owner)
                    ? new DocxParagraphElement(unowned) : element).ToArray() };
            var slots = DocxRenderer.BuildWordReviewCommentAuthorPaletteSlots(changed, [("0", "Comment Reviewer")], default);
            TestAssert.Equal(((byte)248, (byte)220, (byte)221), DocxRenderer.ResolveCommentAuthorFillSnapshot(slots, "0"));
        }
    }

    public static void DocxReviewNonTextAuthorsRenderIndependentCommentColors()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed Windows fonts.");
        CheckNonTextReviewAuthorColors(null);
    }

    internal static void CheckNonTextReviewAuthorColors(IFontResolver? resolver)
    {
        using var golden = ReadNonTextAuthorGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (!row.GetProperty("QualifiesRetainedOwners").GetBoolean()) continue;
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var source = ReadNonTextAuthorSource(row, mode);
                var owners = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
                var runs = owners.SelectMany(paragraph => paragraph.Runs).ToArray();
                var images = owners.SelectMany(paragraph => paragraph.Images).ToArray();
                var pages = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).ToArray();
                TestAssert.Equal(1, pages.Length);
                string content = pages[0].Content;
                string text = DocxPageFootnoteNumberingTests.DecodeText(pages[0]);
                TestAssert.Contains("Public owner paragraph.", text);
                TestAssert.Contains("Public comment anchor.", text);
                if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                {
                    var fill = ReadNonTextAuthorRgb(row, "FillRgb");
                    TestAssert.Contains(DocxTests.FormatPdfRgb(fill, "rg"), content);
                    TestAssert.Contains(DocxTests.FormatPdfRgb(ReadNonTextAuthorRgb(row, "StrokeRgb"), "RG"), content);
                    var other = fill.Red == 248 ? ((byte)213, (byte)237, (byte)255) : ((byte)248, (byte)220, (byte)221);
                    TestAssert.DoesNotContain(DocxTests.FormatPdfRgb(other, "rg"), content);
                }
                TestAssert.True(runs.Zip(owners.SelectMany(paragraph => paragraph.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)) &&
                    images.Zip(owners.SelectMany(paragraph => paragraph.Images)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                    "Author registration cannot rewrite retained text or picture owners.");
            }
        }
    }

    private static JsonDocument ReadNonTextAuthorGolden() => JsonDocument.Parse(File.ReadAllText(
        Path.Combine(NonTextAuthorCases(), "nontext-authors-office-colors.json")));

    private static string NonTextAuthorCases() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ReviewAuthors"));

    private static DocxDocument ReadNonTextAuthorSource(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string input = Path.Combine(NonTextAuthorCases(), row.GetProperty("Id").GetString() + ".docx");
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        return DocxTests.ReadDocx(input, mode);
    }

    private static (byte Red, byte Green, byte Blue) ReadNonTextAuthorRgb(JsonElement row, string property)
    {
        var values = row.GetProperty(property).EnumerateArray().Select(value => value.GetByte()).ToArray();
        return (values[0], values[1], values[2]);
    }
}
