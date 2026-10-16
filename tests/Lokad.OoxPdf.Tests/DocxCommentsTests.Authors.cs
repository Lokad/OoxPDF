using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxReviewCommentAuthorsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed Windows fonts.");
        CheckReviewAuthorColors(null);
    }

    internal static void CheckReviewAuthorColors(IFontResolver? resolver)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "ReviewAuthors", "review-authors-office-colors.json")));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Fixture").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(1, row.GetProperty("OfficeView").GetProperty("RevisionsMode").GetInt32());
            var pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToArray();
            string content = string.Join("\n", pages.Select(page => page.Content));
            var fill = ReadRgb("FillRgb");
            TestAssert.Contains(DocxTests.FormatPdfRgb(fill, "rg"), content);
            TestAssert.Contains(DocxTests.FormatPdfRgb(ReadRgb("StrokeRgb"), "RG"), content);
            var other = fill.Red == 248 ? ((byte)213, (byte)237, (byte)255) : ((byte)248, (byte)220, (byte)221);
            TestAssert.DoesNotContain(DocxTests.FormatPdfRgb(other, "rg"), content);
            (byte Red, byte Green, byte Blue) ReadRgb(string property)
            {
                var values = row.GetProperty(property).EnumerateArray().Select(value => value.GetByte()).ToArray();
                return (values[0], values[1], values[2]);
            }
        }
    }

    public static void DocxReviewCommentAuthorsKeepTextOwnershipCaseAndCancellation()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "RevisionBars", "revision-bar-comment-scaled.docx"));
        var source = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var comments = new[] { (Id: (string?)"1", Author: (string?)"Public reviewer") };
        var slots = DocxRenderer.BuildWordReviewCommentAuthorPaletteSlots(source, comments, default);
        TestAssert.Equal(((byte)213, (byte)237, (byte)255), DocxRenderer.ResolveCommentAuthorFillSnapshot(slots, "1"));
        var owner = DocxBlockTraversal.EnumerateBodyParagraphs(source).First(p => p.Runs.Any(r => r.Revision is not null));
        var unowned = owner with { Runs = owner.Runs.Select(r => r with { Revision = null, Revisions = [] }).ToArray() };
        var changed = source with
        {
            BodyElements = source.BodyElements.Select(element => element is DocxParagraphElement paragraph && ReferenceEquals(paragraph.Paragraph, owner)
                ? new DocxParagraphElement(unowned) : element).ToArray()
        };
        slots = DocxRenderer.BuildWordReviewCommentAuthorPaletteSlots(changed, comments, default);
        TestAssert.Equal(((byte)248, (byte)220, (byte)221), DocxRenderer.ResolveCommentAuthorFillSnapshot(slots, "1"));
        TestAssert.True(unowned.Revisions.Count != 0, "Unowned paragraph metadata cannot invent a text-author slot.");
        slots = DocxRenderer.BuildCommentAuthorPaletteSlots([("1", "Reviewer"), ("2", "reviewer"), ("3", "Reviewer")]);
        TestAssert.Equal(0, slots["1"]);
        TestAssert.Equal(1, slots["2"]);
        TestAssert.Equal(0, slots["3"]);
        TestAssert.Equal(DocxRenderer.NormalizeRevisionAuthorBucketKey("Reviewer"), DocxRenderer.NormalizeRevisionAuthorBucketKey("reviewer"));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxRenderer.BuildWordReviewCommentAuthorPaletteSlots(source, comments, canceled.Token));
    }
}
