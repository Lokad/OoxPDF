using System.Security.Cryptography;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
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
