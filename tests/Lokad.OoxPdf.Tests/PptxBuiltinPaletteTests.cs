using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxBuiltinPaletteTests
{
    public static void PptxBuiltinPalettesMatchIndependentOfficeAcrossThemesAndPlotKinds() => Check("office.json",
        "style-6", "multi-6", "theme-style6", "preferred106-fallback2", "legacy6-only", "count-varying-points",
        "pie-vary", "doughnut-vary", "line-3", "area-3");

    public static void PptxLineDataLabelKeysMatchIndependentOfficePalette() => Check("label-keys-office.json", "line-single-accent-value-keys");

    private static void Check(string referenceFile, params string[] requiredIds)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "BuiltinPalettes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, referenceFile)));
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var row in PptxChartPaletteTests.ReadRequiredPaintReferenceRows(reference, referenceFile, requiredIds))
        {
            string id = row.GetProperty("Id").GetString()!;
            using var stream = File.OpenRead(Path.Combine(directory, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(resolver).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            string[] paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: row.GetProperty("Strokes").GetBoolean())
                .Select(color => $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}").Order(StringComparer.Ordinal).ToArray();
            string[] expected = row.GetProperty("ColoredPaints").EnumerateArray()
                .Select(value => value.GetString()!).Order(StringComparer.Ordinal).ToArray();
            TestAssert.True(paints.SequenceEqual(expected),
                id + ": actual painted palette and legend colors must match independent Office output.");
        }
    }
}
