using System.Xml.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrderedPictureRotationKeepsMeasuredFramesAcrossContainers()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Read(bool rotate) => ReadMutation("CompositionProbes/" + container, mode, false, false,
                (xml, w, _) =>
                {
                    if (!rotate) return;
                    var drawing = xml.Descendants(w + "drawing").Single();
                    var extent = drawing.Descendants(wp + "extent").Single();
                    string width = extent.Attribute("cx")!.Value, height = extent.Attribute("cy")!.Value;
                    extent.SetAttributeValue("cx", height);
                    extent.SetAttributeValue("cy", width);
                    drawing.Descendants(a + "xfrm").Single().SetAttributeValue("rot", "5400000");
                }, partOverride: container is "header" or "footer" ? "word/" + container + "1.xml" : null);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().Single();
            var source = Read(true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Contains("0 -1 1 0 ", page.Content);
            // A usable line slot preserves rotation support and its clip. Full
            // containment is a separate visual question; source transforms can
            // extend beyond that slot. Do not move a qualified origin to fit it.
            TestAssert.True(Regex.IsMatch(page.Content, @" re\s+W\s+n\s+0 -1 1 0 "),
                "The rotated raster retains its explicit measured line clip.");
            var bounds = ReadQuarterTurnRasterBounds(page);
            TestAssert.True(bounds.Left >= 0m && bounds.Top >= 0m && bounds.Right <= (decimal)page.Width &&
                bounds.Bottom <= (decimal)page.Height, "The source transform remains on its physical page.");
            TestAssert.True(diagnostics.All(d => d.Id is not ("DOCX_UNSUPPORTED_IMAGE_TRANSFORM" or "DOCX_UNSUPPORTED_EQUATION")),
                "Supported rotation and mathematics share the same placed line frame in " + container + "/" + mode);
            static string Characters(Lokad.OoxPdf.Pdf.PdfPage value) => string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(value, usePageCoordinates: true)
                    .SelectMany(show => show.Text).Where(character => !char.IsWhiteSpace(character)).Order());
            TestAssert.Equal(Characters(before), Characters(page));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxPictureRotationRetainsIndependentOfficeAnchorAcrossFonts()
    {
        // Word16 SaveAs2, 2026-10-14. Exact input below; reference PDF SHA-256:
        // DDD3E07C9102E10652733A8A8EF1A2BDB34E795CA3CC3F928E3F3B84BDD6DC0A.
        // Its rotated raster is [108,36,180,180] in top-left page coordinates.
        // The page-relative anchor is independent of the requested text face.
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-picture-rotated-anchor.docx"));
        TestAssert.Equal("996248AC03106F2CDA91A707333BA1C17C1B552055B3A5D236DC327D77FCEA5B",
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().Single();
            var actual = ReadQuarterTurnRasterBounds(page);
            decimal[] edges = [actual.Left, actual.Top, actual.Right, actual.Bottom];
            decimal[] word = [108m, 36m, 180m, 180m];
            for (int index = 0; index < edges.Length; index++)
                TestAssert.True(Math.Abs(edges[index] - word[index]) <= .001m,
                    "Rotation must preserve the independently measured Office anchor across text faces and modes.");
        }
    }

    private static (decimal Left, decimal Top, decimal Right, decimal Bottom) ReadQuarterTurnRasterBounds(Lokad.OoxPdf.Pdf.PdfPage page)
    {
        const string number = @"[-\d.]+";
        var draw = Regex.Match(page.Content, @"0 -1 1 0 (?<Tx>" + number + ") (?<Ty>" + number +
            @") cm\s+(?:[-\d.]+\s+){3}[-\d.]+ re\s+W\s+n\s+q\s+(?<Width>" + number +
            ") 0 0 (?<Height>" + number + ") (?<X>" + number + ") (?<Y>" + number + @") cm\s+/" +
            Regex.Escape(page.Images.Single().ResourceName) + @" Do");
        TestAssert.True(draw.Success, "An actual quarter-turn raster draw is required.");
        decimal Value(string name) => decimal.Parse(draw.Groups[name].Value, CultureInfo.InvariantCulture);
        decimal left = Value("Y") + Value("Tx"), top = (decimal)page.Height - (Value("Ty") - Value("X"));
        return (left, top, left + Value("Height"), top + Value("Width"));
    }
}
