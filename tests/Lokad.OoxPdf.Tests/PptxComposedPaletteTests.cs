using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxComposedPaletteTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ComposedBarPalettes"));

    internal static void CheckColors(bool keys)
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var expected = row.GetProperty(keys ? "KeyColors" : "BarColors").EnumerateArray().ToArray();
            if (expected.Length == 0) continue;
            byte[] bytes = File.ReadAllBytes(Path.Combine(Cases, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)));
            bool horizontal = row.GetProperty("Horizontal").GetBoolean();
            var actual = Rectangles(Render(bytes)).Where(r => keys
                ? r.Width is > 1d and < 25d && r.Height is > 1d and < 25d
                : horizontal ? r.Width > 50d && r.Height is > 1d and < 80d
                    : r.Width is > 1d and < 80d && r.Height > 50d)
                .GroupBy(r => r.Color).ToDictionary(group => group.Key, group => group.Count());
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.True(expected.Length == actual.Count, id + ": distinct native colors");
            foreach (var color in expected)
            {
                int[] rgb = color.GetProperty("Rgb").EnumerateArray().Select(v => v.GetInt32()).ToArray();
                var matches = actual.Where(pair => Math.Abs(pair.Key.Red - rgb[0]) <= 1 &&
                    Math.Abs(pair.Key.Green - rgb[1]) <= 1 && Math.Abs(pair.Key.Blue - rgb[2]) <= 1).ToArray();
                TestAssert.True(matches.Length == 1, id + ": native color " + string.Join(',', rgb));
                TestAssert.True(color.GetProperty("Count").GetInt32() == matches[0].Value, id + ": native paint count");
            }
        }
    }

    internal static void CheckIndexBoundsAndFallbacks()
    {
        byte[] original = File.ReadAllBytes(Path.Combine(Cases, "index-shift.pptx"));
        foreach (string? index in new[] { (string?)null, "-1", "unknown", int.MaxValue.ToString(CultureInfo.InvariantCulture) })
        {
            byte[] bytes = EditChart(original, xml =>
            {
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var source = xml.Descendants(c + "barChart").Last().Element(c + "ser")!.Element(c + "idx")!;
                if (index is null) source.Remove(); else source.SetAttributeValue("val", index);
            });
            var bars = Rectangles(Render(bytes)).Where(r => r.Width is > 1d and < 80d && r.Height > 50d).ToArray();
            TestAssert.True(bars.Length == 16, "Sparse source indices do not allocate dense series slots or omit bars.");
            if (index != int.MaxValue.ToString(CultureInfo.InvariantCulture))
                TestAssert.True(bars.All(r => r.Color == new RgbColor(79, 129, 189)),
                    "Missing or invalid source identity retains the existing automatic-theme fallback.");
        }
    }

    internal static void CheckSingleSeriesKeys()
    {
        // Separate native controls retain point-styled category keys in a lone
        // series, including explicit false variation. Values are native bytes.
        byte[] original = File.ReadAllBytes(Path.Combine(Cases, "keys-point-override-valid.pptx"));
        foreach (bool varying in new[] { false, true })
        {
            byte[] bytes = EditChart(original, xml =>
            {
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var plots = xml.Descendants(c + "barChart").ToArray();
                plots[1].Remove();
                plots[0].Element(c + "varyColors")!.SetAttributeValue("val", varying ? "1" : "0");
            });
            var keys = Rectangles(Render(bytes)).Where(r => r.Width is > 1d and < 25d && r.Height is > 1d and < 25d)
                .GroupBy(r => r.Color).ToDictionary(group => group.Key, group => group.Count());
            RgbColor purple = new(85, 0, 170);
            TestAssert.True(keys.TryGetValue(purple, out int points) && points == 2,
                "Native label and category-legend keys preserve the explicit point fill: " + varying);
            TestAssert.Equal(16, keys.Values.Sum());
            if (!varying)
            {
                TestAssert.Equal(2, keys.Count);
                TestAssert.Equal(14, keys[new RgbColor(79, 129, 189)]);
            }
            else
            {
                RgbColor[] colors = [purple, new(170, 70, 67), new(137, 165, 78), new(113, 88, 143),
                    new(65, 152, 175), new(219, 132, 61), new(147, 169, 207), new(209, 147, 146)];
                TestAssert.Equal(8, keys.Count);
                foreach (var color in colors)
                    TestAssert.True(keys.Any(pair => pair.Value == 2 && Math.Abs(pair.Key.Red - color.Red) <= 1 &&
                        Math.Abs(pair.Key.Green - color.Green) <= 1 && Math.Abs(pair.Key.Blue - color.Blue) <= 1),
                        "Native varying category key: " + color);
            }
        }
    }

    private static byte[] EditChart(byte[] source, Action<XDocument> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(source); stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            edit(xml);
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        return stream.ToArray();
    }

    private static PdfPage Render(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    private static IEnumerable<(RgbColor Color, double Width, double Height)> Rectangles(PdfPage page)
    {
        var saved = new Stack<RgbColor>();
        RgbColor color = default;
        foreach (string line in page.Content.Split('\n').Select(line => line.Trim()))
        {
            if (line == "q") saved.Push(color);
            else if (line == "Q") color = saved.Pop();
            else if (line.EndsWith(" rg", StringComparison.Ordinal))
            {
                byte[] rgb = line[..^3].Split(' ').Select(value => (byte)Math.Round(
                    double.Parse(value, CultureInfo.InvariantCulture) * 255d)).ToArray();
                color = new(rgb[0], rgb[1], rgb[2]);
            }
            else if (line.EndsWith(" g", StringComparison.Ordinal))
            {
                byte gray = (byte)Math.Round(double.Parse(line[..^2], CultureInfo.InvariantCulture) * 255d);
                color = new(gray, gray, gray);
            }
            else
            {
                var rectangle = Regex.Match(line, @"^([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re f\*?$");
                if (rectangle.Success) yield return (color,
                    Math.Abs(double.Parse(rectangle.Groups[3].Value, CultureInfo.InvariantCulture)),
                    Math.Abs(double.Parse(rectangle.Groups[4].Value, CultureInfo.InvariantCulture)));
            }
        }
    }
}
