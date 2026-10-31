using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxPointLegendTests
{
    public static void PptxPointLegendsMatchIndependentOfficeIdentityAndPaint()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PointLegends"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office.json")));
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            using var stream = File.OpenRead(Path.Combine(directory, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(resolver).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
            text = string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
            int index = 1;
            foreach (var expected in row.GetProperty("CategoryOccurrences").EnumerateArray())
                TestAssert.Equal(expected.GetInt32(), Occurrences(text, "Category" + index++));
            index = 1;
            foreach (var expected in row.GetProperty("SeriesOccurrences").EnumerateArray())
                TestAssert.Equal(expected.GetInt32(), Occurrences(text, "Series" + index++));

            string[] paints = PptxChartPaletteTests.ReadColoredFills(page.Content)
                .Select(color => $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}").ToArray();
            string[] expectedPaints = row.GetProperty("ColoredPaints").EnumerateArray().Select(value => value.GetString()!).ToArray();
            TestAssert.True(paints.Order(StringComparer.Ordinal).SequenceEqual(expectedPaints.Order(StringComparer.Ordinal)),
                id + ": visible point/legend paints must retain independent Office colors and counts.");
            if (id == "no-fill")
                TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == 0d), "Explicit point no-fill must remain transparent.");
            if (row.GetProperty("LegendEntries").GetInt32() == 6)
            {
                int bars = row.GetProperty("VisibleBarCount").GetInt32();
                TestAssert.True(paints.Take(bars).SequenceEqual(paints.Skip(bars)),
                    id + ": category keys must share point paint without enabling automatic varying colors.");
            }
        }
    }

    private static int Occurrences(string text, string value)
    {
        int count = 0;
        for (int position = 0; (position = text.IndexOf(value, position, StringComparison.Ordinal)) >= 0; position += value.Length) count++;
        return count;
    }
}
