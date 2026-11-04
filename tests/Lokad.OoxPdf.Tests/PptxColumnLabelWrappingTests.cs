using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxColumnLabelWrappingTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases"));

    internal static void CheckAutomatic()
    {
        using var controls = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "ColumnLabelWrapping/office-labels.json")));
        foreach (var control in controls.RootElement.EnumerateArray().Where(item => item.GetProperty("Automatic").GetBoolean()))
        {
            string id = control.GetProperty("Id").GetString()!;
            PdfPage page = Render(File.ReadAllBytes(Path.Combine(Cases, "ColumnLabelWrapping", id + ".pptx")));
            string text = Normalize(string.Concat(Shows(page).Select(show => show.Text)));
            foreach (var label in control.GetProperty("Labels").EnumerateArray())
                TestAssert.True(text.Contains(Normalize(label.GetProperty("Text").GetString()!), StringComparison.Ordinal),
                    id + ": every Office source label survives automatic wrapping.");
            var shape = control.GetProperty("Shape");
            CheckFrames(page, shape.GetProperty("ShapeLeft").GetDouble(),
                page.Height - shape.GetProperty("ShapeTop").GetDouble() - shape.GetProperty("ShapeHeight").GetDouble(),
                shape.GetProperty("ShapeWidth").GetDouble(), shape.GetProperty("ShapeHeight").GetDouble(), id);
        }
    }

    internal static void CheckComposition()
    {
        foreach (string name in new[] { "keys-point-override-valid", "fixed-series-negative-keys", "authored-series-point-keys" })
        {
            byte[] input = File.ReadAllBytes(Path.Combine(Cases, "ComposedBarPalettes", name + ".pptx"));
            PdfPage page = Render(input);
            string text = string.Concat(Shows(page).Select(show => show.Text));
            TestAssert.True(text.Contains("96", StringComparison.Ordinal) && text.Contains("145", StringComparison.Ordinal),
                name + ": both final automatic value suffixes remain encoded.");
            CheckFrames(page, 72d, 72d, 720d, 432d, name, positiveOnly: true);
        }
    }

    public static void PptxColumnLabelSourceRetainsExplicitWrappingConstraints()
    {
        foreach (string wrap in new[] { "none", "square", "future" })
        {
            XElement labels = XElement.Parse($"""
                <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:txPr><a:bodyPr wrap="{wrap}"/></c:txPr>
                </c:dLbls>
                """);
            TestAssert.Equal(wrap, PptxSceneBuilder.ReadChartTextBodyProperties(labels).WrapValue);
        }
        TestAssert.Equal(string.Empty, PptxSceneBuilder.ReadChartTextBodyProperties(null).WrapValue);
    }

    internal static void CheckUnicode()
    {
        byte[] source = File.ReadAllBytes(Path.Combine(Cases, "ColumnLabelWrapping/single-long-word.pptx"));
        using var stream = new MemoryStream();
        stream.Write(source); stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            xml.Descendants(c + "barChart").First().Element(c + "ser")!.Element(c + "tx")!.Element(c + "v")!.Value =
                "AnnualRevenueForecast\U0001f600e\u0301\nNext line";
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        var shows = Shows(Render(stream.ToArray())).ToArray();
        string text = string.Concat(shows.Select(show => show.Text));
        TestAssert.True(Normalize(text).Contains("AnnualRevenueForecast\U0001f600e\u0301Nextline", StringComparison.Ordinal),
            "Explicit breaks and supported supplementary/combining characters survive wrapping.");
        TestAssert.Equal(8, shows.Count(show => show.Text.Contains("e\u0301", StringComparison.Ordinal)));
        TestAssert.True(shows.All(show => !show.Text.StartsWith('\u0301')),
            "A combining sequence remains on one measured line.");
    }

    private static void CheckFrames(PdfPage page, double x, double y, double width, double height, string id, bool positiveOnly = false)
    {
        foreach (var show in Shows(page))
        {
            // Count advances independently from layout. Ignoring negative kerning
            // gives a conservative upper bound; 2pt covers that difference, not
            // a visual threshold. Whole-value clipping exceeds it by many points.
            double advance = show.Text.EnumerateRunes().Sum(rune =>
                (double)show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint(rune.Value))) * show.Size / show.Font.Font.UnitsPerEm;
            if (positiveOnly && show.Text.Contains('-')) continue;
            TestAssert.True(show.X >= x - 2d && show.X + advance <= x + width + 2d &&
                show.Y >= y && show.Y <= y + height,
                id + ": emitted label text must fit inside its chart frame: " + show.Text);
        }
    }

    private static string Normalize(string text) => string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
    private static IEnumerable<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)> Shows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false, fontResourceFilter: resource => resource.StartsWith("CBD", StringComparison.Ordinal));

    private static PdfPage Render(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
