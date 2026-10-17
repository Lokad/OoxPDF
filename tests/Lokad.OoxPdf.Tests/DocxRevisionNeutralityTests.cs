using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRevisionNeutralityTests
{
    public static void DocxRevisionNeutralityRetainsUnownedAndVisibleFallbacks()
    {
        using var golden = ReadGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            row.GetProperty("Mode").GetString() == "all-word-compatible" && row.GetProperty("Variant").GetString() != "control"))
        {
            var document = ReadSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var element = document.BodyElements.OfType<DocxParagraphElement>().First();
            var paragraph = element.Paragraph;
            TestAssert.True(DocxPlainBodyFrame.IsEligible(document), "Actual non-rendering ownership preserves plain source eligibility.");
            foreach (var unsupported in new[]
            {
                paragraph with { NonRenderingRunOwners = [] },
                paragraph with { Revisions = paragraph.Revisions.Select(revision => revision with { }).ToArray() },
                paragraph with { Runs = paragraph.Runs.Select(run => run with { Revision = paragraph.Revisions.Single() }).ToArray() },
                paragraph with { NonRenderingRunOwners = paragraph.NonRenderingRunOwners.Select(owner => owner with { SourceRunIndex = -1 }).ToArray() }
            })
            {
                var changed = document with { BodyElements = document.BodyElements.Select(item => ReferenceEquals(item, element)
                    ? element with { Paragraph = unsupported } : item).ToArray() };
                TestAssert.True(!DocxPlainBodyFrame.IsEligible(changed),
                    "Unowned/value-equal/visible/unknown-position revisions retain their existing fallback.");
            }
        }
    }

    public static void DocxRevisionNeutralityMatchesIndependentWordBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent geometry uses installed Windows fonts.");
        CheckRendering(null, compareGeometry: true);
    }

    internal static void CheckRendering(IFontResolver? resolver, bool compareGeometry)
    {
        using var golden = ReadGolden();
        var controls = new Dictionary<(string Pair, string Mode), string>();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string modeName = row.GetProperty("Mode").GetString()!;
            var mode = modeName == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadSource(row, mode);
            var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var key = (row.GetProperty("Pair").GetString()!, modeName);
            if (row.GetProperty("Variant").GetString() == "control") controls.Add(key, page.Content);
            else TestAssert.Equal(controls[key], page.Content);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            foreach (var expected in row.GetProperty("TextBaselines").EnumerateArray())
            {
                string text = expected.GetProperty("Text").GetString()!;
                var show = shows.Single(show => show.Text.Trim() == text);
                if (compareGeometry)
                    TestAssert.True(Math.Abs(show.Y - expected.GetProperty("Y").GetDouble()) < .25d,
                        "Minimum-height placement must match independent Word in both presets: " + row.GetProperty("Id").GetString() + "/" + modeName + ", y=" + show.Y.ToString(CultureInfo.InvariantCulture));
            }
            TestAssert.True(page.Fonts.All(resource => resource.Font.Font.FamilyName != "Arial"),
                "An invisible formatted source run cannot add an unused face or affect glyph ownership.");
        }
    }

    private static JsonDocument ReadGolden() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "revision-neutrality-office.json")));
    private static string Cases() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "RevisionNeutrality"));
    private static DocxDocument ReadSource(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string path = Path.Combine(Cases(), row.GetProperty("Id").GetString() + ".docx");
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, mode);
    }
}
