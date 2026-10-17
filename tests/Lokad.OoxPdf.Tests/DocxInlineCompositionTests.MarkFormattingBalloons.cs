using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMarkFormattingBalloonsKeepActualSourceOwners()
    {
        using var golden = ReadMarkFormattingGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var document = ReadMarkFormattingDocument(row, OoxPdfDocxMarkupMode.AllMarkup);
            var paragraph = NumberedCounterParagraphs(document).Single(p => p.ListLabel!.Text == "1.");
            var mark = paragraph.ParagraphMarkRun!;
            var revision = mark.Revisions.Single();
            TestAssert.Equal("Font: Bold", DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, default));
            TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners(paragraph.Revisions, paragraph.Runs, default),
                "Run-only ownership cannot substitute for the actual numbered mark.");
            foreach (var invalid in new[]
            {
                paragraph with { ParagraphMarkRun = null },
                paragraph with { ListLabel = null },
                paragraph with { ParagraphMarkRun = mark with { SourceRunIndex = 0 } },
                paragraph with { ParagraphMarkRun = mark with { Revisions = [revision with { }] } }
            })
                TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(invalid, default) is null,
                    "Missing, misplaced or copied mark metadata retains the prior generic fallback.");
            TestAssert.True(ReferenceEquals(revision, paragraph.ParagraphMarkRun!.Revisions.Single()),
                "Selecting a description preserves actual source revision identity.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, canceled.Token));
        }
        var plain = DocxTests.ReadDocx(NumberingMarkCase("list-decoration-mark-inserted.docx"), OoxPdfDocxMarkupMode.AllMarkup);
        var markOnly = NumberedCounterParagraphs(plain).Single(p => p.ListLabel!.Text == "1.");
        TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(markOnly, default) is null,
            "An inserted mark alone cannot manufacture a formatting description.");
    }

    public static void DocxInlineCompositionMarkFormattingBalloonsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native balloon frames use installed Word-compatible fonts.");
        CheckMarkFormattingBalloons(null);
    }

    internal static void CheckMarkFormattingBalloons(IFontResolver? resolver)
    {
        using var golden = ReadMarkFormattingGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool final in new[] { true, false })
        {
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadMarkFormattingDocument(row, mode);
            var paragraphs = NumberedCounterParagraphs(document);
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
            var marks = paragraphs.Select(p => p.ParagraphMarkRun).ToArray();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            if (final)
                TestAssert.True(!shows.Any(show => show.Text.Trim() == "Formatted:"), "Final does not paint revision balloons.");
            else
            {
                TestAssert.True(shows.Any(show => show.Text.Trim() == "Formatted:"),
                    "An actual inserted numbering mark must not force a generic text-insertion balloon.");
                var title = shows.Single(show => show.Text.Trim() == row.GetProperty("Title").GetString());
                var body = shows.Single(show => show.Text.Trim() == row.GetProperty("Body").GetString());
                TestAssert.True(Math.Abs(title.Y - body.Y) < .001d, "The formatting title/body share a row.");
                if (resolver is null)
                    TestAssert.True(Math.Abs(title.X - row.GetProperty("TitleX").GetDouble()) < 1d &&
                        Math.Abs(body.X - row.GetProperty("BodyX").GetDouble()) < 1d,
                        "Descriptions keep the existing one-point Word review-lane tolerance.");
            }
            var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            TestAssert.Equal(18, shows.Count(show => mathFonts.Contains(show.Font)));
            TestAssert.Equal(final ? 3 : 5, shows.Count(show => Regex.IsMatch(show.Text.Trim(), @"^\d+\.$")));
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)) &&
                marks.Zip(paragraphs.Select(p => p.ParagraphMarkRun)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Balloon selection must not rewrite text or actual paragraph-mark owners.");
        }
    }

    private static JsonDocument ReadMarkFormattingGolden() =>
        JsonDocument.Parse(File.ReadAllText(NumberingMarkCase("numbering-mark-formatting-balloons-office.json")));

    private static DocxDocument ReadMarkFormattingDocument(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string path = NumberingMarkCase(row.GetProperty("File").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, mode);
    }
}
