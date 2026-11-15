using System.Security.Cryptography;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingSectionRecordsKeepContentAndPageOwners()
    {
        string folder = Case("SectionRecords");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "expected.json")));
        foreach (var record in manifest.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.Combine(folder, record.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(record.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var pages = RenderAlignmentProbe(DocxTests.ReadDocx(path, mode), mode);
            var expectedPages = record.GetProperty("Pages").EnumerateArray().ToArray();
            TestAssert.Equal(expectedPages.Length, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index]; var expected = expectedPages[index];
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false).ToArray();
                string painted = CompactRecordText(string.Concat(shows.Select(show => show.Text)));
                string headerText = expected.GetProperty("Header").GetString()!;
                foreach (var value in expected.GetProperty("Text").EnumerateArray())
                {
                    string text = CompactRecordText(value.GetString()!);
                    int count = Regex.Matches(painted, Regex.Escape(text)).Count;
                    int sourceCount = 1 + (CompactRecordText(headerText).Contains(text, StringComparison.Ordinal) ? 1 : 0);
                    TestAssert.Equal(sourceCount, count); // The title also occurs inside its header.
                }
                var header = ReadPaintedSourceLabel(page, headerText);
                var footer = ReadPaintedSourceLabel(page, expected.GetProperty("Footer").GetString()!);
                var note = ReadPaintedSourceLabel(page, expected.GetProperty("Note").GetString()!);
                var prose = ReadPaintedSourceLabel(page, expected.GetProperty("Text")[1].GetString()!);
                TestAssert.True(header.Min(glyph => glyph.Ink.Y) > prose.Max(glyph => glyph.Ink.Y + glyph.Ink.Height) &&
                    prose.Min(glyph => glyph.Ink.Y) > note.Max(glyph => glyph.Ink.Y + glyph.Ink.Height) &&
                    note.Min(glyph => glyph.Ink.Y) > footer.Max(glyph => glyph.Ink.Y + glyph.Ink.Height),
                    "Header, body, note and footer keep separate visible stories: " + record.GetProperty("Id").GetString() + "/" + mode + "/" + index);
                foreach (var math in expected.GetProperty("Math").EnumerateArray()) ReadPaintedSourceLabel(page, math.GetString()!, mathOnly: true);
                foreach (var other in expectedPages.Where((_, otherIndex) => otherIndex != index))
                    TestAssert.True(!painted.Contains(CompactRecordText(other.GetProperty("Note").GetString()!), StringComparison.Ordinal),
                        "Another section's note cannot paint on this source page.");
                TestAssert.Equal(expected.GetProperty("Picture").GetBoolean() ? 1 : 0, page.Images.Count);
                TestAssert.Equal(index == 0 ? 3 : 1, page.Annotations.Count);
                var noteLink = page.Annotations.Last();
                TestAssert.Equal(index, noteLink.Destination?.PageIndex ?? -1);
                string mark = expected.GetProperty("Mark").GetString()!;
                TestAssert.True(shows.Any(show => show.Text.Trim() == mark && show.X >= noteLink.X - .01d &&
                    show.X <= noteLink.X + noteLink.Width + .01d && show.Y >= noteLink.Y - .01d && show.Y <= noteLink.Y + noteLink.Height + .01d),
                    "The source note marker paints at its link, rather than being inferred from another page digit.");
                if (index == 0)
                {
                    TestAssert.Equal(pages.Length - 1, page.Annotations[0].Destination?.PageIndex ?? -1);
                    TestAssert.Equal("https://example.invalid/public-service-record", page.Annotations[1].Uri);
                }
                PdfContentValidator.ValidatePage(page, 0, default);
            }
        }
    }

    public static void PortableRenderingOperationsDecksKeepCompleteLegendsAndPictures()
    {
        string folder = Case("OperationsDecks");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "expected.json")));
        foreach (var record in manifest.RootElement.EnumerateArray())
        {
            string path = Path.Combine(folder, record.GetProperty("Id").GetString() + ".pptx");
            TestAssert.Equal(record.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            using var package = OoxPackage.Open(path, default);
            var document = new PptxReader().Read(package, default);
            var renderer = new PptxRenderer(new PinnedResolver());
            var pages = renderer.RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().ToArray();
            var slides = record.GetProperty("Slides").EnumerateArray().ToArray();
            var names = record.GetProperty("LegendNames").EnumerateArray().Select(value => value.GetString()!).ToArray();
            TestAssert.Equal(2, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index];
                TestAssert.Equal(record.GetProperty("Size")[0].GetDouble(), page.Width);
                TestAssert.Equal(record.GetProperty("Size")[1].GetDouble(), page.Height);
                string painted = CompactRecordText(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false).Select(show => show.Text)));
                foreach (var value in slides[index].GetProperty("Text").EnumerateArray())
                {
                    string text = value.GetString()!;
                    if (names.Contains(text, StringComparer.Ordinal)) continue; // Legend role is checked separately below.
                    TestAssert.Equal(1, Regex.Matches(painted, Regex.Escape(CompactRecordText(text))).Count);
                }
                if (index == 0)
                {
                    var frame = record.GetProperty("ChartBox");
                    foreach (string name in names)
                    {
                        var glyphs = ReadPaintedSourceLabel(page, name, fontPrefix: "CL");
                        TestAssert.True(glyphs.All(glyph => glyph.Ink.X >= frame[0].GetDouble() - .01d &&
                            glyph.Ink.X + glyph.Ink.Width <= frame[0].GetDouble() + frame[2].GetDouble() + .01d &&
                            glyph.Ink.Y >= page.Height - frame[1].GetDouble() - frame[3].GetDouble() - .01d &&
                            glyph.Ink.Y + glyph.Ink.Height <= page.Height - frame[1].GetDouble() + .01d),
                            "Every named legend owns complete visible ink inside the chart: " + record.GetProperty("Id").GetString() + "/" + name);
                    }
                }
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Equal(slides[index].GetProperty("Links").GetInt32(), page.Annotations.Count);
                if (index == 1) TestAssert.Equal("https://example.invalid/public-operations-schedule", page.Annotations.Single().Uri);
                PdfContentValidator.ValidatePage(page, 0, default);
            }
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, package, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void PortableRenderingChartSeriesOrderSurvivesXmlReordering()
    {
        foreach (string name in new[] { "service-operations", "survey-operations" })
        {
            string input = Case("OperationsDecks/" + name + ".pptx");
            PdfPage[] Render(string path)
            {
                using var package = OoxPackage.Open(path, default);
                return new PptxRenderer(new PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default),
                    package, null, default).ToBlockingEnumerable().ToArray();
            }
            var baseline = Render(input);
            foreach (string declaration in new[] { "valid", "missing", "negative", "duplicate", "malformed" })
            {
                string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
                File.Copy(input, path);
                new FileInfo(path).IsReadOnly = false; // Only this owned temporary input is writable.
                try
                {
                    string[] xmlNames;
                    using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
                    {
                        const string part = "ppt/charts/chart1.xml";
                        var entry = zip.GetEntry(part)!;
                        XDocument chart;
                        using (var stream = entry.Open()) chart = XDocument.Load(stream);
                        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                        var series = chart.Descendants(c + "barChart").Single().Elements(c + "ser").ToArray();
                        if (declaration == "missing") series[0].Element(c + "order")!.Remove();
                        else if (declaration != "valid") series[0].Element(c + "order")!.SetAttributeValue("val",
                            declaration == "negative" ? "-1" : declaration == "duplicate" ? "1" : "invalid");
                        var anchor = new XComment("Source series container position");
                        series[0].AddBeforeSelf(anchor);
                        foreach (var item in series) item.Remove();
                        anchor.AddBeforeSelf(series.Reverse());
                        anchor.Remove();
                        xmlNames = series.Reverse().Select(item => item.Element(c + "tx")!.Value).ToArray();
                        entry.Delete();
                        using var writer = new StreamWriter(zip.CreateEntry(part).Open(), new UTF8Encoding(false));
                        chart.Save(writer, SaveOptions.DisableFormatting);
                    }
                    var actual = Render(path);
                    if (declaration == "valid")
                    {
                        for (int index = 0; index < actual.Length; index++)
                            TestAssert.Equal(baseline[index].Content, actual[index].Content);
                    }
                    else
                    {
                        string legend = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual[0], true,
                            honorActualText: false, fontResourceFilter: resource => resource.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
                        TestAssert.Equal(string.Concat(xmlNames), legend);
                    }
                    foreach (var page in actual) PdfContentValidator.ValidatePage(page, 0, default);
                }
                finally { File.Delete(path); }
            }
        }
    }

    private static string CompactRecordText(string text) => Regex.Replace(text, @"\s+", "");
}
