using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
    public static void MathFontDelimiterAssembliesKeepCompleteInkLogicalTextAndFixedFallbacks()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "PortableFonts", "DejaVuMathTeXGyre.ttf"));
        var font = OpenTypeFont.Load(File.ReadAllBytes(path));
        var mappings = new List<(ushort Glyph, int Unicode)> { (font.MapCodePoint('A'), 'A') };
        ushort lastPart = 0;
        foreach (int point in new[] { 40, 41 })
        {
            ushort canonical = font.MapCodePoint(point);
            mappings.Add((canonical, point));
            for (double height = 0d; font.TryGetMathVerticalVariant(canonical, height, default, out var variant); height = variant.Advance + 1d)
                mappings.Add((variant.Glyph, point));
            TestAssert.True(font.TryGetMathVerticalAssembly(canonical, default, out var assembly),
                "Both growing delimiters have complete font-defined parts.");
            foreach (var part in TestAssert.NotNull(assembly).Parts) mappings.Add((part.Glyph, point));
            lastPart = assembly!.Parts[^1].Glyph;
        }
        var resource = new PdfFontResource("Delimiters", PdfEmbeddedFont.CreateForGlyphMappings(font, mappings, default));
        var missing = new PdfFontResource("MissingLastPart", PdfEmbeddedFont.CreateForGlyphMappings(font,
            mappings.Where(mapping => mapping.Glyph != lastPart).ToArray(), default));
        var literal = new DocxMathRow([new DocxMathRun(new("A", 60d, "008000", false, false, false, null, "Math"), null, null)
            { NormalTextValue = true }]);
        foreach (bool grow in new[] { true, false })
        {
            var delimiter = new DocxMathDelimiter([literal]) { GrowValue = grow,
                ControlText = new("", 12d, "FF0000", false, false, false, null, "Math") };
            var equation = new DocxEquation(new([delimiter]), 0, 0, false, null, null, null) { ParentFontSize = 11d };
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var measured, out var reason,
                resolveControlFont: _ => resource), reason ?? "A growing delimiter can use its complete assembly.");
            var box = TestAssert.NotNull(measured);
            TestAssert.Equal("(A)", string.Concat(box.Glyphs.Select(glyph => char.ConvertFromUtf32(glyph.SourceCodePoint))));
            foreach (var glyph in box.Glyphs.Where(glyph => glyph.SourceCodePoint != 'A'))
            {
                TestAssert.True(ReferenceEquals(glyph.ControlOwner, delimiter) && glyph.FontSize == 12d && glyph.Color.Red == 255 &&
                    (grow ? glyph.AssemblyParts.Count > 1 : glyph.AssemblyParts.Count == 0),
                    "Growing and fixed delimiters retain the actual control and source size/ink.");
                foreach (var part in glyph.AssemblyParts)
                {
                    TestAssert.True(font.TryGetGlyphBounds(part.OriginalGlyph, out var ink), "Every part has complete glyph bounds.");
                    double scale = glyph.FontSize / font.UnitsPerEm;
                    TestAssert.True(glyph.X + part.X + ink.XMax * scale <= box.Width + 1e-9 &&
                        glyph.BaselineY + part.BaselineY + ink.YMax * scale <= box.AboveBaseline + 1e-9 &&
                        glyph.BaselineY + part.BaselineY + ink.YMin * scale >= -box.BelowBaseline - 1e-9,
                        "The measured delimiter encloses every physical part before painting.");
                }
            }
            var graphics = new PdfGraphicsBuilder();
            DocxMathBoxPainter.Paint(box, graphics, 10d, 10d, default);
            var page = new PdfPage(200d, 200d, graphics.ToString(), [resource]);
            PdfContentValidator.ValidatePage(page, 0, default);
            TestAssert.Equal("(A)", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text)));
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [page], default);
            TestAssert.True(output.Length > 0, "Complete delimiter assemblies write as one logical scalar per mark.");
            bool complete = DocxMathLayout.TryMeasure(equation, _ => missing, default, out var absent, out _, resolveControlFont: _ => missing);
            TestAssert.True(grow ? !complete && absent is null : complete,
                "Missing late assembly encoding rejects the whole growing box; fixed glyphs need no parts.");
            if (!grow) continue;
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => resource,
                cancellation.Token, out _, out _, resolveControlFont: _ => resource));
            using var budget = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => resource,
                default, out _, out _, resolveControlFont: _ => resource));
        }
    }
}
