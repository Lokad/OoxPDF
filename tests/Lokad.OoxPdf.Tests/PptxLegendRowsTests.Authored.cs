using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxExplicitSeriesLegendDeletionKeepsThePlottedSeries()
    {
        const string fixture = "PointLegends/multiple-series.pptx";
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture));
        byte[] source = File.ReadAllBytes(path);
        TestAssert.Equal("61C8A00DF69BF961D9D8EF808D741415FC890A6DB17166CBC0801E3210720EC5",
            Convert.ToHexString(SHA256.HashData(source)));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        PdfPage baseline = Render();
        Check(baseline, ["Series 1", "Series 2"]);
        TestAssert.Equal(baseline.Content, Render(Entry(1, "0"), Entry(8, "1")).Content);

        Check(Render(Entry(1, "1")), ["Series 1"]);
        Check(Render(Entry(0, null)), ["Series 2"]);
        Check(Render(Entry(0, "true"), Entry(1, "1")), []);

        XElement Entry(int index, string? value) => new(c + "legendEntry",
            new XElement(c + "idx", new XAttribute("val", index)),
            new XElement(c + "delete", value is null ? null : new XAttribute("val", value)));

        PdfPage Render(params XElement[] entries)
        {
            using var stream = new MemoryStream();
            stream.Write(source);
            stream.Position = 0;
            if (entries.Length > 0)
            {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
                const string part = "ppt/charts/chart1.xml";
                ZipArchiveEntry chartEntry = archive.GetEntry(part)!;
                XDocument chart;
                using (Stream input = chartEntry.Open()) chart = XDocument.Load(input);
                chart.Root!.Element(c + "chart")!.Element(c + "legend")!
                    .Element(c + "legendPos")!.AddAfterSelf(entries);
                chartEntry.Delete();
                using Stream output = archive.CreateEntry(part).Open();
                chart.Save(output);
            }

            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        }

        static void Check(PdfPage page, string[] expectedNames)
        {
            var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal));
            TestAssert.True(names.Select(show => show.Text).SequenceEqual(expectedNames),
                "The legend paints precisely the surviving series names in order.");
            var fills = Regex.Matches(page.Content, @"[-\d.]+ [-\d.]+ ([-\d.]+) ([-\d.]+) re f")
                .Select(match => (Width: double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                    Height: double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture))).ToArray();
            TestAssert.Equal(expectedNames.Length, fills.Count(fill => fill.Width is > 3d and < 10d && fill.Width == fill.Height));
            TestAssert.Equal(12, fills.Count(fill => fill.Width is > 15d and < 100d && fill.Height > 100d));
        }
    }

    public static void PptxAuthoredPolarLegendKeepsEveryCompleteNameInsideItsClip()
    {
        // This is the unchanged public source whose independent Office export
        // paints six names in an authored144x108pt legend. A whole-page image
        // gate also passed the old candidate even though four names were clipped.
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PolarLegends", "authored-top.pptx"));
        TestAssert.Equal("367BD89B5D867DB15B95DA48BF26F1A502982A6011612F5A76CA1D51AE3FB634",
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var package = OoxPackage.Open(path, default);
        var document = new PptxReader().Read(package, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
        var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(string.Concat(Enumerable.Range(1, 6).Select(i => $"Category {i}")),
            string.Concat(names.Select(show => show.Text)));
        double[] clip = ReadPolarLegendClip(page, 144);
        foreach (var show in names)
        {
            var font = show.Font.Font;
            double scale = show.Size / font.UnitsPerEm;
            double x = show.X;
            foreach (char character in show.Text)
            {
                ushort glyph = font.MapCodePoint(character);
                bool measured = font.TryGetGlyphBounds(glyph, out var ink);
                TestAssert.True(measured || char.IsWhiteSpace(character), "The pinned font supplies bounds for every painted letter and digit.");
                if (measured)
                    TestAssert.True(x + ink.XMin * scale >= clip[0] - .01d &&
                        x + ink.XMax * scale <= clip[0] + clip[2] + .01d &&
                        show.Y + ink.YMin * scale >= clip[1] - .01d &&
                        show.Y + ink.YMax * scale <= clip[1] + clip[3] + .01d,
                        $"Every glyph in {show.Text} must paint inside the actual legend clip; emitted text alone is insufficient.");
                x += font.GetAdvanceWidth(glyph) * scale;
            }
        }
        TestAssert.True(!diagnostics.Any(item => item.Message.Contains("overflowing entries", StringComparison.Ordinal)),
            "All six names fit the supplied container without shrinking their text.");
    }
}
