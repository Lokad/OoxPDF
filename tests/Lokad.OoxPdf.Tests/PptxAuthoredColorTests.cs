using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAuthoredColorTests
{
    public static void PptxAuthoredTintShadeChartsMatchIndependentOfficePaints() => Check("Chart");

    public static void PptxAuthoredTintShadeShapesMatchOfficeAndPreserveOrder() => Check("Shapes");

    private static void Check(string kind)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "AuthoredColorTransforms"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office.json")));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (row.GetProperty("Kind").GetString() != kind) continue;
            string input = row.GetProperty("Input").GetString()!;
            using var stream = File.OpenRead(Path.Combine(directory, input));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var actual = ReadPaints(page.Content).ToArray();
            var expected = row.GetProperty("ExpectedPaints").EnumerateArray()
                .Select(color => color.EnumerateArray().Select(channel => channel.GetInt32()).ToArray()).ToArray();
            // The chart canvas contributes white paint; shapes include a deliberate
            // white endpoint, so only charts discard their background here.
            if (kind == "Chart") actual = actual.Where(color => !color.SequenceEqual(new[] { 255, 255, 255 })).ToArray();
            else
            {
                TestAssert.Equal(expected.Length + 1, actual.Length);
                TestAssert.True(actual[0].SequenceEqual(new[] { 255, 255, 255 }), "Expected the page background before source shapes.");
                actual = actual.Skip(1).ToArray();
            }
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < expected.Length; i++)
                TestAssert.True(actual[i].SequenceEqual(expected[i]), $"{input} paint {i}: expected {string.Join(',', expected[i])}, got {string.Join(',', actual[i])}.");
        }
    }

    private static IEnumerable<int[]> ReadPaints(string content)
    {
        int[] color = [0, 0, 0];
        var saved = new Stack<int[]>();
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim();
            if (line == "q") saved.Push(color);
            else if (line == "Q") color = saved.Pop();
            else if (line.EndsWith(" rg", StringComparison.Ordinal)) color = line[..^3]
                .Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Channel).ToArray();
            else if (line.EndsWith(" g", StringComparison.Ordinal))
            {
                int gray = Channel(line[..^2]);
                color = [gray, gray, gray];
            }
            else if (line is "f" or "f*" || line.EndsWith(" f", StringComparison.Ordinal) || line.EndsWith(" f*", StringComparison.Ordinal)) yield return color;
        }
        static int Channel(string value) => (int)Math.Round(255 * double.Parse(value, CultureInfo.InvariantCulture), MidpointRounding.AwayFromZero);
    }
}
