using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class PptxComposedAxisTests
{
    internal static void CheckSameSideAdmission()
    {
        // Independent PowerPoint16 controls: same-side-calibri7/9, retained in
        // artifacts/tnumeric-side-spacing-1105/office. Both secondary ink blocks
        // start at78.504pt. Preserve the existing1pt geometry tolerance across
        // portable font substitution; full Office comparisons remain separate.
        foreach (double size in new[] { 7d, 9d })
        foreach (bool columns in new[] { false, true })
        {
            var baseline = NumericShows(RenderSameSide(size, columns));
            var overlay = NumericShows(RenderSameSide(size, columns, overlayLegend: true));
            TestAssert.True(baseline.Select(show => show.Text).SequenceEqual(
                new[] { "0", "25", "50", "75", "100", "0", "50", "100", "150", "200" }),
                "Both authored numeric intervals retain every caption.");
            TestAssert.True(Math.Abs(baseline.Skip(5).Min(show => show.X) - 78.504d) <= 1d,
                "The outer numeric strip stays inside its independently observed automatic frame.");
            TestAssert.True(baseline.Select(show => show.Text).SequenceEqual(overlay.Select(show => show.Text)),
                "Overlay furniture preserves the complete numeric captions.");
            for (int i = 0; i < baseline.Length; i++)
                TestAssert.True(Math.Abs(baseline[i].X - overlay[i].X) < .002d,
                    "An overlay legend cannot switch ownership of the same-side numeric reserve.");
        }
    }

    internal static void CheckAutoScaleAdmission()
    {
        foreach (bool columns in new[] { false, true })
        {
            var small = NumericShows(RenderSameSide(12d, columns, automaticMaximum: 100d));
            var large = NumericShows(RenderSameSide(12d, columns, automaticMaximum: 100000d));
            TestAssert.True(small.Skip(5).Any(show => show.Text == "100") &&
                large.Skip(5).Any(show => show.Text == "100000"),
                "The secondary axis uses its own automatic data range.");
            TestAssert.True(large.Take(5).Min(show => show.X) > small.Take(5).Min(show => show.X) + 10d,
                "Wider secondary captions reserve actual space before the primary strip.");
            TestAssert.True(large.Skip(5).Min(show => show.X) >= 72d,
                "Automatically scaled secondary captions remain inside the chart frame.");
        }
    }

    internal static void CheckOppositeSideAdmission()
    {
        // Independent opposite-side-calibri9 Office PDF: primary left78.504pt,
        // secondary right785.593pt. Use the same1pt portable geometry gate.
        var shows = NumericShows(RenderSameSide(9d, columns: false, oppositeSides: true));
        TestAssert.True(shows.Select(show => show.Text).SequenceEqual(
            new[] { "0", "25", "50", "75", "100", "0", "50", "100", "150", "200" }),
            "Opposite-side bands retain both complete authored intervals.");
        TestAssert.True(Math.Abs(shows.Take(5).Min(show => show.X) - 78.504d) <= 1d &&
            Math.Abs(shows.Skip(5).Max(show => show.Right) - 785.593d) <= 1d,
            "Each side uses its own independently observed single-axis band.");
    }

    private static (string Text, double X, double Right)[] NumericShows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal))
        .Select(show => (show.Text, show.X, show.X + show.Font.MeasureTextPoints(show.Text, show.Size))).ToArray();

    private static PdfPage RenderSameSide(double fontSize, bool columns, bool overlayLegend = false,
        double? automaticMaximum = null, bool oppositeSides = false)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ComposedAxes"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office-labels.json")));
        var row = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "distinct-secondary");
        byte[] source = File.ReadAllBytes(Path.Combine(cases, row.GetProperty("Input").GetString()!));
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(source)));
        using var sourceStream = new MemoryStream(source);
        using var archive = new ZipArchive(sourceStream, ZipArchiveMode.Read);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
        });
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/charts/chart1.xml"]));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        xml.Descendants(c + "plotArea").Single().Element(c + "layout")?.RemoveNodes();
        foreach (var axis in xml.Descendants(c + "valAx"))
        {
            axis.Element(c + "axPos")!.SetAttributeValue("val", "l");
            axis.Element(c + "tickLblPos")!.SetAttributeValue("val", "low");
            axis.Element(c + "txPr")?.Remove();
            axis.Add(new XElement(c + "txPr", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr",
                    new XAttribute("sz", (int)(fontSize * 100d)), new XElement(a + "latin", new XAttribute("typeface", "Calibri")))))));
        }
        if (oppositeSides)
        {
            var axis = xml.Descendants(c + "valAx").Last();
            axis.Element(c + "axPos")!.SetAttributeValue("val", "r");
            axis.Element(c + "tickLblPos")!.SetAttributeValue("val", "high");
        }
        var secondary = xml.Descendants(c + "lineChart").Single();
        if (automaticMaximum is double maximum)
        {
            var scaling = xml.Descendants(c + "valAx").Last().Element(c + "scaling")!;
            scaling.Elements().Where(element => element.Name == c + "min" || element.Name == c + "max").Remove();
            var values = secondary.Descendants(c + "val").Single().Descendants(c + "pt").Select(point => point.Element(c + "v")!).ToArray();
            for (int i = 0; i < values.Length; i++)
                values[i].Value = (maximum * i / (values.Length - 1)).ToString(System.Globalization.CultureInfo.InvariantCulture);
            xml.Descendants(c + "valAx").Last().Element(c + "majorUnit")?.Remove();
        }
        if (columns)
        {
            secondary.Name = c + "barChart";
            secondary.AddFirst(new XElement(c + "barDir", new XAttribute("val", "col")));
            secondary.Element(c + "grouping")?.SetAttributeValue("val", "clustered");
        }
        if (overlayLegend)
            xml.Descendants(c + "chart").Single().Add(new XElement(c + "legend",
                new XElement(c + "legendPos", new XAttribute("val", "b")), new XElement(c + "overlay", new XAttribute("val", "1"))));
        parts["ppt/charts/chart1.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    internal static void CheckLabels(bool hidePrimary)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ComposedAxes"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office-labels.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            byte[] source = File.ReadAllBytes(Path.Combine(cases, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(source)));
            using var stream = new MemoryStream();
            stream.Write(source); stream.Position = 0;
            if (hidePrimary)
            {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
                var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var axis = xml.Descendants(c + "valAx").Single(axis =>
                    (string?)axis.Element(c + "axId")?.Attribute("val") == row.GetProperty("PrimaryAxisId").GetString());
                axis.Element(c + "tickLblPos")!.SetAttributeValue("val", "none");
                entry.Delete();
                using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
                xml.Save(output);
            }
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
            string[] actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false, fontResourceFilter:
                name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(show => show.Text).Order(StringComparer.Ordinal).ToArray();
            string[] expected = row.GetProperty(hidePrimary ? "SecondaryLabels" : "OfficeNumericLabels")
                .EnumerateArray().Select(label => label.GetString()!).Order(StringComparer.Ordinal).ToArray();
            TestAssert.True(expected.SequenceEqual(actual),
                "Only actual visible source axes may emit numeric strips: " + row.GetProperty("Id").GetString() +
                "/hidePrimary=" + hidePrimary + "; expected " + string.Join(",", expected) + "; actual " + string.Join(",", actual));
        }
    }
}
