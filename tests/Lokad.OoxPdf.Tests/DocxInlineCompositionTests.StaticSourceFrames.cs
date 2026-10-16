using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStaticSourceFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent frame geometry uses installed Word-compatible fonts.");
        CheckStaticSourceFrames(null);
    }

    internal static void CheckStaticSourceFrames(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "static-source-frames-office.json")));
        var headers = new Dictionary<string, (double Actual, double Expected)>();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var document = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup);
            var paragraphs = document.HeaderParagraphsByType.Values.Concat(document.FooterParagraphsByType.Values)
                .SelectMany(value => value).ToArray();
            var owners = paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var summary = shows.Single(show => show.Text.StartsWith("Summary: ", StringComparison.Ordinal));
            var relationships = shows.Single(show => show.Text.StartsWith("The two relationships ", StringComparison.Ordinal));
            var title = shows.Single(show => show.Text.Trim() == "Formatted:");
            TestAssert.Contains("share the same input series.", string.Concat(shows.Select(show => show.Text)));
            TestAssert.True(summary.Y > 0d && summary.Y < page.Height && relationships.Y > 0d && relationships.Y < page.Height,
                "Supported static prose stays on its page with bookmarks, links and formatting metadata: " + id);
            bool footer = row.GetProperty("IsFooter").GetBoolean();
            if (footer)
                TestAssert.True(title.Y > summary.Y,
                    "Footer formatting placement belongs above the occupied story, including changes in later paragraphs: " + id);
            if (resolver is null)
            {
                TestAssert.True(Math.Abs(summary.Y - row.GetProperty("SummaryY").GetDouble()) < 1d &&
                    Math.Abs(relationships.Y - row.GetProperty("RelationshipsY").GetDouble()) < 1d,
                    "Static prose retains independent physical Word baselines: " + id);
                if (footer)
                    TestAssert.True(Math.Abs(title.Y - row.GetProperty("TitleY").GetDouble()) < 1d,
                        "Footer formatting follows the measured frame within the existing one-point geometry tolerance: " + id);
                else headers.Add(id, (title.Y, row.GetProperty("TitleY").GetDouble()));
            }
            TestAssert.True(owners.Zip(paragraphs.SelectMany(paragraph => paragraph.Runs))
                .All(pair => ReferenceEquals(pair.First, pair.Second)), "Frame placement cannot rewrite static source runs.");
        }
        if (resolver is null)
        {
            var first = headers["math-run-format-header-bold"];
            var second = headers["header-second-paragraph"];
            TestAssert.True(Math.Abs((first.Actual - second.Actual) - (first.Expected - second.Expected)) < 1d,
                "Header balloons continue following the changed paragraph; the existing absolute header residual remains separate.");
        }
    }
}
