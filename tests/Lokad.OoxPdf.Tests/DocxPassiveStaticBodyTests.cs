using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPassiveStaticBodyTests
{
    public static void DocxPassiveStaticTextPreservesIndependentWordBodyBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent minimum body origins require installed Windows fonts.");
        CheckBodyBaselines(null, compareWord: true);
    }

    internal static void CheckBodyBaselines(IFontResolver? resolver, bool compareWord)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("CompositionProbes/minimum-passive-static-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Case("CompositionProbes/" + id + ".docx");
            string original = Case(row.GetProperty("SourceInput").GetString()!.Replace("tests/Lokad.OoxPdf.Tests/Cases/", "", StringComparison.Ordinal));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(row.GetProperty("SourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(original))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var source = Lines(input);
            var baseline = Lines(original).Where(line => line.Text.StartsWith("Public minimum", StringComparison.Ordinal)).ToArray();
            var body = source.Where(line => line.Text.StartsWith("Public minimum", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(4, body.Length);
            foreach (var before in baseline)
            {
                var after = body.Single(line => line.Text == before.Text);
                TestAssert.Equal(before.Page, after.Page);
                TestAssert.True(Math.Abs(before.X - after.X) < .001d && Math.Abs(before.Y - after.Y) < .001d,
                    "Adding a measured margin story must preserve the ordinary body's glyph origins with native or portable fonts.");
            }
            foreach (var expected in row.GetProperty("WordProseAnchors").EnumerateArray())
            {
                var actual = source.Single(line => line.Text == expected.GetProperty("Text").GetString());
                if (!compareWord) continue;
                TestAssert.Equal(expected.GetProperty("Page").GetInt32(), actual.Page);
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "Body and passive story retain their independent one-point Word origins.");
            }

            (int Page, double X, double Y, string Text)[] Lines(string path) =>
                new DocxRenderer(resolver, mode, geometry).RenderBlankPages(DocxTests.ReadDocx(path, mode), null, default)
                    .SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .GroupBy(show => show.Y).Select(line => (index + 1, line.Min(show => show.X), page.Height - line.Key,
                            Regex.Replace(string.Concat(line.OrderBy(show => show.X).Select(show => show.Text))
                                .Normalize(NormalizationForm.FormKC), @"\s+", " ").Trim())))
                    .ToArray();
        }
    }

    public static void DocxPassiveStaticTextKeepsIntrusiveDynamicAndUnmeasuredFallbacks()
    {
        foreach (string kind in new[] { "header", "footer" })
        {
            string input = Case("CompositionProbes/minimum-passive-" + kind + "-ordinary12.docx");
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
            var layout = Layout(document);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, layout, out _), "A complete plain margin frame admits ordinary minimum leading.");
            var page = layout.Pages[0];
            var intrusive = kind == "header"
                ? layout with { HeaderContentBottomByPage = new Dictionary<int, double> { [0] = page.Height - page.MarginTop - 1d } }
                : layout with { FooterContentTopByPage = new Dictionary<int, double> { [0] = page.MarginBottom + 1d } };
            var unmeasured = layout with { HeaderContentBottomByPage = new Dictionary<int, double>() };
            var annotationLane = layout with { Pages = [page with { MarkupMarginReservePoints = 20d }] };
            foreach (var unsupported in new[] { intrusive, unmeasured, annotationLane })
                TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(document, unsupported, out var unchanged) && ReferenceEquals(unsupported, unchanged),
                    "Intrusive, unmeasured and annotation-lane frames retain the existing fallback.");
            var dynamic = ReadDynamicStory(input, kind);
            var dynamicLayout = Layout(dynamic);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(dynamic, dynamicLayout, out var fallback) && ReferenceEquals(dynamicLayout, fallback),
                "A page-dependent story remains outside plain static-text admission.");
        }
    }

    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, default),
            resolver.Resolve(new FontRequest("Calibri", false, false)), default, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
    }

    private static DocxDocument ReadDynamicStory(string input, string kind)
    {
        string temporary = Path.Combine(Path.GetTempPath(), "ooxpdf-passive-" + Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(input, temporary);
        try
        {
            // Frozen inputs are read-only; the independent temporary copy is
            // intentionally mutable, without changing the frozen source.
            File.SetAttributes(temporary, FileAttributes.Normal);
            using (var archive = ZipFile.Open(temporary, ZipArchiveMode.Update))
            {
                string part = "word/passive-body-" + kind + ".xml";
                var entry = archive.GetEntry(part)!;
                XDocument xml;
                using (var stream = entry.Open()) xml = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var paragraph = xml.Root!.Element(w + "p")!;
                var run = paragraph.Element(w + "r")!;
                run.Remove();
                paragraph.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"), run));
                entry.Delete();
                using var output = archive.CreateEntry(part).Open();
                xml.Save(output);
            }
            return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup);
        }
        finally { File.Delete(temporary); }
    }

    private static string Case(string relative) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", relative));
}
