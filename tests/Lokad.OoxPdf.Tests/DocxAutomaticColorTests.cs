using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxAutomaticColorTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "AutomaticColor"));

    public static void DocxAutomaticColorMatchesIndependentWordPaletteAndHeldOutThreshold()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray().Take(2))
        foreach (var sample in control.GetProperty("Colors").EnumerateArray())
        {
            string? declared = sample.GetProperty("DeclaredColor").GetString();
            TestAssert.True(RgbColor.TryParse(sample.GetProperty("Fill").GetString(), out var background), "Expected a pinned RGB fill.");
            var actual = DocxTextColor.Read(declared) ?? DocxTextColor.Automatic(background);
            TestAssert.Equal(sample.GetProperty("WordColor").GetString(), Hex(actual));
        }
        TestAssert.True(DocxTextColor.Read(null) is null && DocxTextColor.Read("auto") is null,
            "Absent and explicit automatic foregrounds must retain their ownership until painting.");
        TestAssert.True(DocxTextColor.Read("000000") == default(RgbColor) &&
            DocxTextColor.Read("invalid") == default(RgbColor), "Literal black and malformed-color fallback remain explicit black.");
        TestAssert.Equal(default(RgbColor), DocxTextColor.Automatic(null));
    }

    public static void DocxAutomaticColorPaintsSupportedShadingOwnershipInBothModes()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var samples = control.GetProperty("Colors").EnumerateArray().Where(sample => sample.GetProperty("Supported").GetBoolean()).ToArray();
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(input, mode);
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var found = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var page in renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable())
                foreach (Match block in Regex.Matches(page.Content, @"(?s)BT\s+(.*?)\s+ET"))
                {
                    var fragment = new PdfPage(page.Width, page.Height, block.Value, page.Fonts);
                    string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(fragment).Select(show => show.Text));
                    foreach (var sample in samples)
                    {
                        string label = sample.GetProperty("Label").GetString()!;
                        if (!text.Contains(label, StringComparison.Ordinal)) continue;
                        TestAssert.True(found.TryAdd(label, ReadInk(block.Value)), "Each source label must be painted exactly once: " + label);
                    }
                }
                foreach (var sample in samples)
                {
                    string label = sample.GetProperty("Label").GetString()!;
                    TestAssert.True(found.TryGetValue(label, out string? actual), "Supported content must remain present: " + label);
                    TestAssert.Equal(sample.GetProperty("WordColor").GetString(), actual);
                }
            }
        }
    }

    internal static string ReadInk(string block)
    {
        var color = Regex.Match(block, @"(?m)^(?<r>[\d.]+)(?: (?<g>[\d.]+) (?<b>[\d.]+))? (?<op>rg|g)\r?$");
        TestAssert.True(color.Success, "Expected explicit ink in each generated text object.");
        byte Read(string component) => (byte)Math.Round(double.Parse(color.Groups[component].Value, CultureInfo.InvariantCulture) * 255d);
        byte red = Read("r");
        return Hex(color.Groups["op"].Value == "g" ? new(red, red, red) : new(red, Read("g"), Read("b")));
    }

    private static string Hex(RgbColor color) => $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
}
