using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string AdmissionCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ManualAxisAdmission"));
    private const string AdmissionFaceBoundary = "category12-y15";

    public static void PptxManualInnerNumericStripsFitTheirIndependentOfficeSide()
    {
        using var oracle = AdmissionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("XAdmitted").GetBoolean()))
        {
            PdfPage page = RenderAdmission(row);
            var box = AdmissionBounds(page);
            bool right = row.GetProperty("NumericRight").GetBoolean();
            double expected = AdmissionInset(row, page);
            double[] chart = AdmissionNumbers(row, "ChartFrame");
            TestAssert.True(Math.Abs((right ? chart[2] - box.Right : box.Left - chart[0]) - expected) <= 1d,
                row.GetProperty("Id").GetString() + ": the native strip clearance follows the selected face's measured advance.");
            double[] native = AdmissionNumbers(row, "NativeFrame");
            TestAssert.True(Math.Abs(box.Right - box.Left - native[2] + native[0]) < .002d &&
                Math.Abs(box.Top - native[1]) < .002d && Math.Abs(box.Bottom - native[3]) < .002d,
                "Admitting X preserves independently authored width and Y.");
            var shows = AdmissionNumericShows(page);
            TestAssert.True(shows.Length > 0 && shows.All(show => show.X >= chart[0] &&
                show.X + show.Font.MeasureTextPoints(show.Text, show.Size) <= chart[2]),
                "Every emitted numeric fragment stays inside the chart frame.");
        }
    }

    public static void PptxManualInnerVerticalAdmissionMatchesIndependentOfficeFrames()
    {
        using var oracle = AdmissionOracle();
        // Calibri crosses the endcap boundary in this fixture; the pinned
        // DejaVu face fits it. Frozen Office qualification retains its absolute
        // rectangle gate, while the portable check below retains source content.
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("CheckY").GetBoolean() &&
            row.GetProperty("Id").GetString() != AdmissionFaceBoundary))
        {
            var box = AdmissionBounds(RenderAdmission(row));
            double[] expected = AdmissionNumbers(row, "NativeFrame");
            TestAssert.True(Math.Abs(box.Top - expected[1]) <= 1d && Math.Abs(box.Bottom - expected[3]) <= 1d,
                row.GetProperty("Id").GetString() + ": native visible-axis admission and hidden-axis ownership.");
        }
    }

    public static void PptxManualInnerFaceBoundaryKeepsSourceCaptions()
    {
        using var oracle = AdmissionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            row.GetProperty("Id").GetString() is AdmissionFaceBoundary or "axes12-global6-y15" or "axes12-global24-y15"))
        {
            PdfPage page = RenderAdmission(row);
            var categories = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text).ToArray();
            TestAssert.True(categories.SequenceEqual(Enumerable.Range(1, 8).Select(index => "W" + index)),
                "Face substitution and chart defaults preserve every emitted source category.");
            TestAssert.True(AdmissionNumericShows(page).Select(show => show.Text).SequenceEqual(new[] { "7", "32", "57", "82", "107" }),
                "The face-dependent admission boundary preserves all source intervals.");
        }
    }

    public static void PptxManualInnerAdmissionPreservesQualifiedAuthoredFrames()
    {
        using var oracle = AdmissionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("Unchanged").GetBoolean()))
        {
            var box = AdmissionBounds(RenderAdmission(row));
            double[] expected = AdmissionNumbers(row, "NativeFrame");
            TestAssert.True(new[] { box.Left, box.Top, box.Right, box.Bottom }.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .002d),
                row.GetProperty("Id").GetString() + ": global defaults and an X correction must not refit an admitted Y interval.");
        }
    }

    public static void PptxManualAxisAdmissionMovesLinkedLabelsWithTheirSourcePlot()
    {
        using var oracle = AdmissionOracle();
        var row = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "x-left24-y-fit");
        PdfPage page = RenderAdmission(row, link: true);
        var show = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Single(show => show.Text == "L");
        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/manual-caption", annotation.Uri);
        TestAssert.True(annotation.X <= show.X + .002d && annotation.X + annotation.Width >=
            show.X + show.Font.MeasureTextPoints(show.Text, show.Size) - .002d &&
            annotation.Y < show.Y && annotation.Y + annotation.Height > show.Y,
            "The annotation covers the resulting glyph geometry.");
        double[] chart = AdmissionNumbers(row, "ChartFrame");
        double target = chart[0] + AdmissionInset(row, page);
        double center = annotation.X + annotation.Width / 2d;
        TestAssert.True(center >= target && center < target + 45d,
            "The linked source-zero label follows the admitted first category slot.");
    }

    private static double AdmissionInset(JsonElement row, PdfPage page)
    {
        double[] native = AdmissionNumbers(row, "NativeFrame"), chart = AdmissionNumbers(row, "ChartFrame");
        bool right = row.GetProperty("NumericRight").GetBoolean();
        double nativeInset = right ? chart[2] - native[2] : native[0] - chart[0];
        var shows = AdmissionNumericShows(page);
        return nativeInset - row.GetProperty("NativeNumericAdvance").GetDouble() +
            shows.Max(show => show.Font.MeasureTextPoints(show.Text, show.Size));
    }

    private static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] AdmissionNumericShows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();

    private static (double Left, double Top, double Right, double Bottom) AdmissionBounds(PdfPage page)
    {
        var lines = Lines(page);
        var vertical = lines.Single(line => line.Color == new RgbColor(255, 0, 0) &&
            Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d).V;
        var horizontal = lines.Single(line => line.Color == new RgbColor(0, 170, 0) &&
            Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d).V;
        return (Math.Min(horizontal[0], horizontal[2]), page.Height - Math.Max(vertical[1], vertical[3]),
            Math.Max(horizontal[0], horizontal[2]), page.Height - Math.Min(vertical[1], vertical[3]));
    }

    private static double[] AdmissionNumbers(JsonElement row, string name) => row.GetProperty(name).EnumerateArray().Select(value => value.GetDouble()).ToArray();
    private static JsonDocument AdmissionOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(AdmissionCases, "office.json")));

    private static PdfPage RenderAdmission(JsonElement row, bool link = false, string? caseDirectory = null)
    {
        string path = Path.Combine(caseDirectory ?? AdmissionCases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var source = entry.Open(); using var target = new MemoryStream(); source.CopyTo(target); return target.ToArray();
        });
        if (link)
        {
            using var source = new MemoryStream(parts["ppt/charts/chart1.xml"]);
            var xml = XDocument.Load(source);
            var series = xml.Descendants(CrossingC + "barChart").Single().Elements(CrossingC + "ser").First();
            series.Element(CrossingC + "cat")!.AddBeforeSelf(XElement.Parse("""
                <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="admissionLink"/></a:rPr><a:t>L</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="ctr"/><c:showVal val="0"/></c:dLbl></c:dLbls>
                """));
            parts["ppt/charts/chart1.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
            parts["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="admissionLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/manual-caption" TargetMode="External"/></Relationships>
                """);
        }
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
