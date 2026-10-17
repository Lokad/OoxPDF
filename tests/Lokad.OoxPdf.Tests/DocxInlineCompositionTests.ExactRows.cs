using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionExactRowsKeepIndependentRawContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent exact rows use installed Windows fonts.");
        // Deliberately clipped controls still contain raw text operations in
        // Word. These checks do not claim that every operation is visible.
        CheckEquationRevisions(null, false, "exact-row-frames-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionExactRowsKeepIndependentContainedFrameAndFollowingFlow()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent exact frames use installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "exact-row-frames.json")));
        var complete = manifest.RootElement.EnumerateArray().ToDictionary(row => row.GetProperty("Id").GetString()!,
            row => row.GetProperty("ContentMustFit").GetBoolean(), StringComparer.Ordinal);
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "exact-row-frames-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            if (!complete[id]) continue;
            var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var expected = row.GetProperty("BalloonFrames")[0];
            TestAssert.True(Math.Abs(balloon.Y + balloon.Height - expected.GetProperty("Y").GetDouble() - expected.GetProperty("Height").GetDouble()) < 1d,
                $"A complete exact cell retains its independently measured printed top: {id}/{balloon.Y + balloon.Height}.");
            if (!row.TryGetProperty("ProseAnchors", out var following)) continue;
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            foreach (var anchor in following.EnumerateArray())
            {
                string prefix = anchor.GetProperty("Prefix").GetString()!;
                var show = shows.Single(show => show.Text.TrimStart().StartsWith(prefix, StringComparison.Ordinal));
                TestAssert.True(Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                    "Following prose retains the independent exact-row advance.");
            }
        }
    }
}
