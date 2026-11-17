using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class SvgTextTests
{
    public static void SvgInheritedPaintPropertiesKeepNearestAndLocalOwners()
    {
        // Literal source values are independent of the production property reader.
        foreach (var (property, value, operation) in new[] {
            ("fill-opacity", ".25", ""), ("stroke-opacity", ".25", ""),
            ("stroke-width", "6", "6 w"), ("stroke-linecap", "round", "1 J"),
            ("stroke-linejoin", "bevel", "2 j"), ("stroke-dasharray", "6 3", "[6 3 ] 0 d"),
            ("stroke-dashoffset", "2", "[6 3 ] 2 d"), ("stroke-miterlimit", "8", "8 M") })
        {
            string Path(string attributes) => "<path d=\"M20 20H120V80H20Z\" fill=\"red\" stroke=\"blue\" " +
                (property == "stroke-dashoffset" ? "stroke-dasharray=\"6 3\" " : "") + attributes + "/>";
            var direct = Paint(Svg(Path(property + "=\"" + value + "\"")));
            if (operation.Length != 0) TestAssert.Contains(operation, direct.ToString());
            else TestAssert.True(direct.ExtGStates.Any(state =>
                property == "fill-opacity" ? state.FillAlpha == .25d && state.StrokeAlpha == 1d :
                    state.FillAlpha == 1d && state.StrokeAlpha == .25d), "The literal opacity owns its painting operation.");
            foreach (bool css in new[] { false, true })
            {
                string declaration = css ? "style=\"" + property + ":" + value + "\"" : property + "=\"" + value + "\"";
                AssertEquivalentSvgPaint(direct, Paint(Svg("<g " + declaration + "><g>" + Path("") + "</g></g>")), property);
                // A local declaration wins over an ancestor, including CSS there.
                AssertEquivalentSvgPaint(direct, Paint(Svg("<g style=\"" + property + ":0\">" +
                    Path(declaration) + "</g>")), property + " local override");
            }
        }
        var nearest = Paint(Svg("<g fill-opacity=\".1\"><g style=\"fill-opacity:.25\"><rect width=\"50\" height=\"50\" fill=\"red\"/></g></g>"));
        AssertEquivalentSvgPaint(Paint(Svg("<rect width=\"50\" height=\"50\" fill=\"red\" fill-opacity=\".25\"/>")), nearest,
            "Inherited opacity selects the nearest value rather than multiplying declarations.");
    }

    public static void SvgInheritedOpacityPreservesTransparentOverlayContent()
    {
        const string picture = "<rect width=\"100\" height=\"100\" fill=\"red\"/>";
        const string overlay = "<rect width=\"100\" height=\"100\" fill=\"white\"/>";
        var direct = Paint(Svg(picture + overlay.Replace("/>", " fill-opacity=\"0\"/>", StringComparison.Ordinal)));
        var inherited = Paint(Svg(picture + "<g fill-opacity=\"0\">" + overlay + "</g>"));
        AssertEquivalentSvgPaint(direct, inherited, "A transparent overlay cannot cover its underlying picture.");
        TestAssert.True(inherited.ExtGStates.Any(state => state.FillAlpha == 0d), "The overlay has zero fill alpha.");
        TestAssert.Equal(0, inherited.Groups.Count); // fill opacity does not create a compositing group.
    }

    public static void SvgZeroWidthStrokesKeepTheirIndependentFill()
    {
        const string definitions = "<defs><linearGradient id=\"g\"><stop offset=\"0\" stop-color=\"red\"/>" +
            "<stop offset=\"1\" stop-color=\"blue\"/></linearGradient></defs>";
        foreach (string fill in new[] { "none", "red", "url(#g)" })
        foreach (string stroke in new[] { "blue", "url(#g)", "url(#missing)" })
        {
            string Path(string attributes) => "<path d=\"M20 20H120V80H20Z\" fill=\"" + fill + "\" " + attributes + "/>";
            var expected = Paint(Svg(definitions + Path("stroke=\"none\"")));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = Paint(Svg(definitions + Path("stroke=\"" + stroke + "\" stroke-width=\"0\"")), diagnostics);
            AssertEquivalentSvgPaint(expected, actual, "Zero width removes only the stroke, including an unused paint server.");
            TestAssert.Equal(0, diagnostics.Count);
            AssertEquivalentSvgPaint(expected, Paint(Svg(definitions + "<g stroke-width=\"0\">" +
                Path("stroke=\"" + stroke + "\"") + "</g>")), "Inherited zero width has the same nonpainting meaning.");
        }
    }

    public static void SvgInheritedGradientStrokeOpacityKeepsNumericFallbackContract()
    {
        const string definitions = "<defs><linearGradient id=\"g\"><stop offset=\"0\" stop-color=\"red\"/>" +
            "<stop offset=\"1\" stop-color=\"blue\"/></linearGradient></defs>";
        const string path = "<path d=\"M20 20H120\" fill=\"none\" stroke=\"url(#g)\" stroke-width=\"4\" {0}/>";
        foreach (string opacity in new[] { ".25", "50%", "bad" })
        {
            var directDiagnostics = new List<OoxPdfDiagnostic>();
            var inheritedDiagnostics = new List<OoxPdfDiagnostic>();
            var direct = Paint(Svg(definitions + string.Format(path, "stroke-opacity=\"" + opacity + "\"")), directDiagnostics);
            var inherited = Paint(Svg(definitions + "<g stroke-opacity=\"" + opacity + "\">" + string.Format(path, "") + "</g>"), inheritedDiagnostics);
            AssertEquivalentSvgPaint(direct, inherited, "Gradient opacity eligibility follows the resolved source value.");
            TestAssert.True(directDiagnostics.SequenceEqual(inheritedDiagnostics), "The same unsupported value keeps the same diagnosis.");
            if (opacity == ".25") TestAssert.True(inherited.ExtGStates.Any(state => state.StrokeAlpha == .25d), "Numeric gradient alpha remains supported.");
            else TestAssert.True(inheritedDiagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "Unqualified gradient alpha keeps its diagnosed fallback.");
        }
    }

    public static void SvgInheritedPaintSurvivesDocxAndPptxContainers()
    {
        foreach (bool docx in new[] { false, true })
        foreach (var pair in new[] {
            ("<g fill-opacity=\"0\"><rect width=\"100\" height=\"100\" fill=\"white\"/></g>",
                "<rect width=\"100\" height=\"100\" fill=\"white\" fill-opacity=\"0\"/>"),
            ("<g stroke-width=\"10\"><path d=\"M20 50H120\" fill=\"none\" stroke=\"blue\"/></g>",
                "<path d=\"M20 50H120\" fill=\"none\" stroke=\"blue\" stroke-width=\"10\"/>") })
        {
            PdfPage Render(string contents)
            {
                string input = Input(docx, Svg("<rect width=\"100\" height=\"100\" fill=\"red\"/>" + contents));
                var resolver = new PortableRenderingTests.PinnedResolver();
                var diagnostics = new List<OoxPdfDiagnostic>();
                using var package = OoxPackage.Open(input, default);
                var page = docx
                    ? new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                        .RenderPagesAsync(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), diagnostics.Add, default).ToBlockingEnumerable().Single()
                    : new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default), package,
                        diagnostics.Add, default).ToBlockingEnumerable().Single();
                TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "Supported paint must reach both containers.");
                return page;
            }
            var inherited = Render(pair.Item1);
            var direct = Render(pair.Item2);
            TestAssert.Equal(direct.Content, inherited.Content);
            TestAssert.True(direct.ExtGStates.SequenceEqual(inherited.ExtGStates), "Container resource alpha follows the same source rule.");
        }
    }

    private static void AssertEquivalentSvgPaint(PdfGraphicsBuilder expected, PdfGraphicsBuilder actual, string reason)
    {
        TestAssert.Equal(expected.ToString(), actual.ToString());
        TestAssert.True(expected.ExtGStates.SequenceEqual(actual.ExtGStates), reason);
    }
}
