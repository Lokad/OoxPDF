using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStaticListLabelsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent list geometry uses installed Word-compatible fonts.");
        CheckStaticListLabels(null);
    }

    internal static void CheckStaticListLabels(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "static-list-labels-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = document.HeaderParagraphsByType.Values.Concat(document.FooterParagraphsByType.Values)
                .SelectMany(value => value).ToArray();
            var sourceRuns = paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var labels = shows.Where(show => show.Text.Trim() is "1." or "2." or "3.").ToArray();
            TestAssert.Equal("1.,2.,3.", string.Join(',', labels.Select(show => show.Text.Trim()).Distinct()));
            foreach (var expected in row.GetProperty("Labels").EnumerateArray())
            {
                string text = expected.GetProperty("Text").GetString()!;
                var paints = labels.Where(show => show.Text.Trim() == text).ToArray();
                TestAssert.True(paints.Select(show => show.Y).Distinct().Count() == 1,
                    "Each source label owns one row, including synthetic bold paint: " + id);
                var label = paints[0];
                TestAssert.True(label.X >= 0d && label.X < page.Width && label.Y > 0d && label.Y < page.Height,
                    "Static list labels stay on their page in both public presets: " + id);
                if (resolver is null)
                    TestAssert.True(Math.Abs(label.X - expected.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(label.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                        "List labels retain independent physical Word geometry: " + id);
            }
            TestAssert.Contains("share the same input series.", string.Concat(shows.Select(show => show.Text)));
            TestAssert.True(sourceRuns.Zip(paragraphs.SelectMany(paragraph => paragraph.Runs))
                .All(pair => ReferenceEquals(pair.First, pair.Second)), "Deriving ordered text items retains original static source runs.");
        }
    }
}
