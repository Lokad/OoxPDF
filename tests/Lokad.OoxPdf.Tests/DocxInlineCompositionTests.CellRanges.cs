using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCellRangesKeepIndependentWordCountsContentAndProvenance()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cell ranges use installed Windows fonts.");
        CheckCellRanges(null);
    }

    internal static void CheckCellRanges(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "cell-ranges-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            TestAssert.Equal(2, document.Tables.Single().Rows.SelectMany(source => source.Cells).Sum(cell => cell.Revisions.Count));
            var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloons = renderer.InspectMarkupBalloons(document);
            TestAssert.Equal(row.GetProperty("BalloonFrames").GetArrayLength(), balloons.Count);
            TestAssert.Equal(2, balloons.Sum(balloon => balloon.CandidateCount));
            TestAssert.Equal(2, balloons.Sum(balloon => balloon.RevisionCandidateCount));
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var math = shows.Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(row.GetProperty("ImageCount").GetInt32(), page.Images.Count);
            string prose = System.Text.RegularExpressions.Regex.Replace(
                string.Join(" ", shows.Where(show => !fonts.Contains(show.Font)).Select(show => show.Text)), @"\s+", " ").Trim();
            foreach (var text in row.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(text.GetString()!, prose);
            string[] labels = shows.Select(show => show.Text.Trim()).Where(text => text is "Inserted Cells" or "Deleted Cells")
                .Order(StringComparer.Ordinal).ToArray();
            string[] expected = row.GetProperty("BalloonTexts").EnumerateArray().Select(text => text.GetString()!).Order(StringComparer.Ordinal).ToArray();
            TestAssert.True(labels.SequenceEqual(expected), "Printed range labels retain independently observed Word multiplicity and revision kinds: " + id);
        }
    }

    public static void DocxInlineCompositionCellRangesKeepOtherMarkupModesSeparate()
    {
        var document = ReadMutation("CompositionProbes/cell-connector-same-author-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false);
        TestAssert.Equal(2, new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document).Count);
        var final = ReadMutation("CompositionProbes/cell-connector-same-author-ins", OoxPdfDocxMarkupMode.Final, false, false);
        TestAssert.Equal(0, new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(final).Count);
    }
}
