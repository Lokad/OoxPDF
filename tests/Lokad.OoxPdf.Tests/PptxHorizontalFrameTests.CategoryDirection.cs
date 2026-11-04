using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxHorizontalFrameTests
{
    private static readonly XNamespace DirectionC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DirectionA = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static string CategoryDirectionCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "HorizontalCategoryOrder"));

    public static void PptxHorizontalCategoryBarsKeepIndependentOfficeSourceRows()
    {
        using var oracle = CategoryDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            PdfPage page = RenderCategoryDirection(row);
            var actual = DirectionBars(page);
            var expected = row.GetProperty("NativeBars").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            var axis = PptxColumnIntervalTests.Lines(page).Single(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d).V;
            double left = Math.Min(axis[0], axis[2]), width = Math.Abs(axis[2] - axis[0]);
            double[] nativeX = row.GetProperty("NativePlotX").EnumerateArray().Select(x => x.GetDouble()).ToArray();
            for (int i = 0; i < actual.Length; i++)
            {
                double[] box = expected[i].GetProperty("Box").EnumerateArray().Select(x => x.GetDouble()).ToArray();
                double physicalCenter = page.Height - actual[i].Y - actual[i].Height / 2d;
                TestAssert.True(Math.Abs(physicalCenter - (box[1] + box[3]) / 2d) <= 1d &&
                    Math.Abs(actual[i].Height - box[3] + box[1]) <= 1d,
                    row.GetProperty("Id").GetString() + ": each source bar retains its native category/series row.");
                TestAssert.True(Math.Abs((actual[i].X + actual[i].Width - left) / width -
                    (box[2] - nativeX[0]) / (nativeX[1] - nativeX[0])) < .002d,
                    "Category direction preserves the numeric value associated with each source bar.");
            }
        }
    }

    public static void PptxHorizontalCategoryCaptionsKeepIndependentOfficeOrderAndSpacing()
    {
        using var oracle = CategoryDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(RenderCategoryDirection(row), usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).OrderByDescending(show => show.Y).ToArray();
            var expected = row.GetProperty("NativeCategories").EnumerateArray().ToArray();
            TestAssert.True(actual.Select(show => show.Text).SequenceEqual(expected.Select(item => item.GetProperty("Text").GetString()!)),
                row.GetProperty("Id").GetString() + ": category captions keep every native source row in physical page order.");
            for (int i = 1; i < actual.Length; i++)
                TestAssert.True(Math.Abs(actual[0].Y - actual[i].Y -
                    (expected[i].GetProperty("Top").GetDouble() - expected[0].GetProperty("Top").GetDouble())) <= 1d,
                    "Native row spacing survives portable font baseline differences.");
        }
    }

    public static void PptxReversedHorizontalLabelFramesAndKeysFollowTheirSourceBars()
    {
        using var oracle = CategoryDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("MultiSeries").GetBoolean()))
        foreach (bool key in new[] { false, true })
        {
            PdfPage page = RenderCategoryDirection(row, xml =>
            {
                XElement labels = xml.Descendants(DirectionC + "barChart").Single().Element(DirectionC + "dLbls")!;
                labels.Element(DirectionC + "showLegendKey")!.SetAttributeValue("val", key ? "1" : "0");
                labels.Add(new XElement(DirectionC + "spPr", new XElement(DirectionA + "solidFill",
                    new XElement(DirectionA + "srgbClr", new XAttribute("val", "FFFF00")))));
            });
            var all = PptxChartSpacingTests.ReadBarRectangles(page.Content).ToArray();
            var frames = all.Where(box => box.Width < 150d && box.Height > 20d).OrderBy(box => box.X + box.Width / 2d).ToArray();
            var bars = DirectionBars(page);
            var native = row.GetProperty("NativeBars").EnumerateArray().ToArray();
            TestAssert.Equal(native.Length, frames.Length);
            for (int i = 0; i < frames.Length; i++)
            {
                double center = frames[i].Y + frames[i].Height / 2d;
                TestAssert.True(Math.Abs(center - bars[i].Y - bars[i].Height / 2d) < .002d,
                    "Text/key occupied frames share painted cluster geometry.");
                double[] box = native[i].GetProperty("Box").EnumerateArray().Select(x => x.GetDouble()).ToArray();
                TestAssert.True(Math.Abs(page.Height - center - (box[1] + box[3]) / 2d) <= 1d,
                    "A matching but unreversed label/bar pair does not satisfy native source placement.");
            }
            TestAssert.Equal(key ? frames.Length : 0, all.Count(box => box.Width < 20d && box.Height < 20d));
        }
    }

    public static void PptxReversedHorizontalRichLabelLinksFollowTheirSourceRow()
    {
        using var oracle = CategoryDirectionOracle();
        var row = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "clustered-two-reversed");
        PdfPage page = RenderCategoryDirection(row, xml =>
        {
            XElement series = xml.Descendants(DirectionC + "barChart").Single().Elements(DirectionC + "ser").First();
            series.Add(XElement.Parse("""
                <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="categoryLink"/></a:rPr><a:t>LINK</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="ctr"/><c:showVal val="0"/></c:dLbl>
                </c:dLbls>
                """));
        }, link: true);
        var show = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Single(show => show.Text == "LINK");
        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/category-row", annotation.Uri);
        double advance = show.Font.MeasureTextPoints(show.Text, show.Size);
        TestAssert.True(annotation.X <= show.X + .002d && annotation.X + annotation.Width >= show.X + advance - .002d &&
            annotation.Y < show.Y && annotation.Y + annotation.Height > show.Y,
            "The hyperlink covers the emitted rich-text run after placement.");
        double[] nativeFrame = row.GetProperty("NativeFrameY").EnumerateArray().Select(x => x.GetDouble()).ToArray();
        double physicalCenter = page.Height - annotation.Y - annotation.Height / 2d;
        TestAssert.True(physicalCenter > nativeFrame[0] && physicalCenter < nativeFrame[0] + (nativeFrame[1] - nativeFrame[0]) / 8d,
            "The linked source-zero label remains in the first native category slot.");
    }

    private static (double X, double Y, double Width, double Height)[] DirectionBars(PdfPage page) =>
        PptxChartSpacingTests.ReadBarRectangles(page.Content).Where(box => box.Width > 100d && box.Height is > 2d and < 40d)
            .OrderBy(box => box.X + box.Width).ToArray();

    private static JsonDocument CategoryDirectionOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(CategoryDirectionCases, "office.json")));

    private static PdfPage RenderCategoryDirection(JsonElement row, Action<XDocument>? edit = null, bool link = false)
    {
        string path = Path.Combine(CategoryDirectionCases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var source = entry.Open(); using var target = new MemoryStream(); source.CopyTo(target); return target.ToArray();
        });
        if (edit is not null)
        {
            using var source = new MemoryStream(parts["ppt/charts/chart1.xml"]);
            var xml = XDocument.Load(source); edit(xml);
            parts["ppt/charts/chart1.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
        }
        if (link) parts["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="categoryLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/category-row" TargetMode="External"/></Relationships>
            """);
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
