using System.Text;
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
        const string path = "<path d=\"M20 20H120V80H20Z\" fill=\"none\" stroke=\"url(#g)\" stroke-width=\"4\" {0}/>";
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
        const string strokePath = "<path d=\"M140 50H220\" fill=\"none\" ";
        foreach (bool docx in new[] { false, true })
        foreach (var pair in new[] {
            ("<g fill-opacity=\"0\"><rect width=\"100\" height=\"100\" fill=\"white\"/></g>",
                "<rect width=\"100\" height=\"100\" fill=\"white\" fill-opacity=\"0\"/>"),
            ("<g stroke-width=\"10\"><path d=\"M20 50H120\" fill=\"none\" stroke=\"blue\"/></g>",
                "<path d=\"M20 50H120\" fill=\"none\" stroke=\"blue\" stroke-width=\"10\"/>"),
            ("<g fill=\"green\" stroke=\"blue\" stroke-width=\"10\"><g fill=\"inherit\">" +
                "<rect x=\"110\" width=\"20\" height=\"80\" fill=\"inherit\" stroke=\"none\"/>" +
                strokePath + "stroke=\"red\" style=\"stroke: inherit\" stroke-width=\"inherit\"/></g></g>",
                "<rect x=\"110\" width=\"20\" height=\"80\" fill=\"green\"/>" +
                strokePath + "stroke=\"blue\" stroke-width=\"10\"/>") })
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

    public static void SvgPrimitivePixelUnitsMatchUserCoordinates()
    {
        string Shapes(string unit) => $$"""
            <g transform="translate(6 3) scale(.95)">
              <rect x="10{{unit}}" y="20{{unit}}" width="45{{unit}}" height="45{{unit}}" fill="red"/>
              <rect x="70" y="20" width="45" height="45" rx="12{{unit}}" ry="9{{unit}}" fill="yellow"/>
              <circle cx="150{{unit}}" cy="42{{unit}}" r="21{{unit}}" fill="green"/>
              <ellipse cx="215{{unit}}" cy="42{{unit}}" rx="27{{unit}}" ry="16{{unit}}" fill="fuchsia"/>
              <line x1="265{{unit}}" x2="350{{unit}}" y1="42{{unit}}" y2="42{{unit}}" stroke="blue" stroke-width="6{{unit}}"/>
            </g>
            """;
        var plain = Paint(Svg(Shapes("")));
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pixels = Paint(Svg(Shapes("px")), diagnostics);
        foreach (string paint in new[] { "1 0 0 rg", "1 1 0 rg", "0 0.502 0 rg", "1 0 1 rg", "0 0 1 RG" })
            TestAssert.Contains(paint, plain.ToString());
        AssertEquivalentSvgPaint(plain, pixels, "Pixels and unitless lengths share the SVG user coordinate system.");
        TestAssert.Equal(0, diagnostics.Count);
    }

    public static void SvgUserSpaceGradientPixelUnitsMatchUserCoordinates()
    {
        string Picture(string unit) => Svg($$"""
            <defs>
              <linearGradient id="linear" gradientUnits="userSpaceOnUse"
                              x1="10{{unit}}" y1="15{{unit}}" x2="90{{unit}}" y2="15{{unit}}">
                <stop offset="0" stop-color="red"/><stop offset="1" stop-color="blue"/>
              </linearGradient>
              <radialGradient id="radial" gradientUnits="userSpaceOnUse"
                              cx="220{{unit}}" cy="45{{unit}}" r="35{{unit}}"
                              fx="220{{unit}}" fy="45{{unit}}">
                <stop offset="0" stop-color="yellow"/><stop offset="1" stop-color="green"/>
              </radialGradient>
            </defs>
            <rect x="10" y="15" width="80" height="60" fill="url(#linear)"/>
            <rect x="20" y="103" width="60" height="36" fill="none" stroke="url(#linear)" stroke-width="6"/>
            <circle cx="220" cy="45" r="35" fill="url(#radial)"/>
            """);

        var diagnostics = new List<OoxPdfDiagnostic>();
        var expected = Paint(Picture(""), diagnostics);
        var actual = Paint(Picture("px"), diagnostics);
        TestAssert.Contains("0.988 0 0.012 rg", expected.ToString());
        TestAssert.Contains("0.012 0 0.988 rg", expected.ToString());
        TestAssert.Equal(1, expected.ShadingPatterns.Count);
        TestAssert.Equal(1, expected.Shadings.Count);
        AssertEquivalentSvgPaint(expected, actual, "User-space gradient coordinates give the same paint with pixel units.");
        TestAssert.Equal(expected.ShadingPatterns[0].Pattern.ResourceKey, actual.ShadingPatterns.Single().Pattern.ResourceKey);
        TestAssert.Equal(expected.Shadings[0].Shading.ResourceKey, actual.Shadings.Single().Shading.ResourceKey);
        TestAssert.Equal(0, diagnostics.Count);
    }

    public static void SvgExplicitViewportDimensionsKeepPaintWithoutViewBox()
    {
        const string content = """
            <rect x="12" y="15" width="76" height="44" fill="red"/>
            <g transform="translate(24 10)">
              <path d="M120 25 L170 64" fill="none" stroke="blue" stroke-width="6"/>
              <text x="135" y="110" font-size="24" fill="#112233">I</text>
            </g>
            """;
        string Picture(string sizing) =>
            $"""<svg xmlns="http://www.w3.org/2000/svg" {sizing} font-family="DejaVu Sans">{content}</svg>""";

        var diagnostics = new List<OoxPdfDiagnostic>();
        var expected = Paint(Picture("width=\"300\" height=\"150\" viewBox=\"0 0 300 150\""), diagnostics);
        string paint = expected.ToString();
        TestAssert.Contains("1 0 0 rg", paint);
        TestAssert.Contains("0 0 1 RG", paint);
        TestAssert.Equal(1, LabelPaths(paint));

        foreach (string sizing in new[]
        {
            "width=\"300\" height=\"150\"",
            "width=\"300px\" height=\"150px\"",
            "width=\"600\" height=\"300\" style=\"width:300px; height:150px\"",
            "width=\"600\" height=\"300\" viewBox=\"0 0 300 150\""
        })
        {
            AssertEquivalentSvgPaint(expected, Paint(Picture(sizing), diagnostics),
                "Explicit dimensions supply the same coordinate frame as the equivalent viewBox.");
        }
        TestAssert.Equal(0, diagnostics.Count);
    }

    public static void SvgUnknownSourceFrameRetainsIgnoredDiagnostic()
    {
        foreach (string sizing in new[]
        {
            "height=\"150\"",
            "width=\"100%\" height=\"150\"",
            "width=\"0\" height=\"150\"",
            "width=\"300\" height=\"150\" viewBox=\"invalid\""
        })
        {
            byte[] source = Encoding.UTF8.GetBytes(
                $"""<svg xmlns="http://www.w3.org/2000/svg" {sizing}><rect width="50" height="50" fill="red"/></svg>""");
            var graphics = new PdfGraphicsBuilder();
            var diagnostics = new List<OoxPdfDiagnostic>();
            TestAssert.True(!PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 400, 160, source, diagnostics.Add, 1, "/test.svg", default),
                $"The source dimensions must resolve for {sizing}.");
            TestAssert.Equal("SVG_UNSUPPORTED_CONTENT", diagnostics.Single().Id);
            TestAssert.Equal("Ignored", diagnostics.Single().Fallback);
            TestAssert.DoesNotContain("1 0 0 rg", graphics.ToString());
        }
    }

    public static void SvgCurrentColorUsesEachPaintedElementsColor()
    {
        var expected = Paint(Svg("""
            <rect x="20" y="20" width="30" height="30" fill="red"/>
            <line x1="40" y1="80" x2="80" y2="80" fill="none" stroke="blue" stroke-width="4"/>
            <text x="20" y="120" font-size="24" fill="#112233">I</text>
            <rect x="75" y="20" width="30" height="30" fill="black"/>
            """));
        var diagnostics = new List<OoxPdfDiagnostic>();
        var actual = Paint(Svg("""
            <g color="#112233" fill="currentColor">
              <rect x="20" y="20" width="30" height="30" color="red"/>
              <line x1="40" y1="80" x2="80" y2="80" fill="none" color="green"
                    style="color:blue;stroke:currentColor" stroke-width="4"/>
              <text x="20" y="120" font-size="24" color="currentColor">I</text>
            </g>
            <rect x="75" y="20" width="30" height="30" fill="currentColor"/>
            """), diagnostics);
        foreach (string operation in new[] { "1 0 0 rg", "0 0 1 RG", LabelColor, "75 140 m" })
            TestAssert.Contains(operation, expected.ToString());
        TestAssert.Equal(1, LabelPaths(expected.ToString()));
        AssertEquivalentSvgPaint(expected, actual, "Each solid paint uses the color of the element being painted.");
        TestAssert.Equal(0, diagnostics.Count);
    }

    public static void SvgMissingPaintServersUseAuthoredFallback()
    {
        const string rectangle = "<rect x=\"20\" y=\"20\" width=\"30\" height=\"30\" fill=\"{0}\"/>";
        const string line = "<line x1=\"40\" y1=\"80\" x2=\"80\" y2=\"80\" fill=\"none\" color=\"blue\" stroke=\"{0}\" stroke-width=\"4\"/>";
        const string text = "<text x=\"20\" y=\"120\" font-size=\"24\" fill=\"{0}\" stroke=\"{1}\">I</text>";
        var literal = Paint(Svg(string.Format(rectangle, "red") + string.Format(line, "blue") +
            string.Format(text, "#112233", "none")));
        var diagnostics = new List<OoxPdfDiagnostic>();
        var fallback = Paint(Svg(string.Format(rectangle, "url(#missing-fill) red") +
            string.Format(line, "url(#missing-stroke) currentColor") +
            string.Format(text, "url(#missing-text) #112233", "url(#missing-outline) none") +
            string.Format(rectangle, "url(#missing-optional) none")), diagnostics);
        TestAssert.Contains("1 0 0 rg", literal.ToString());
        TestAssert.Contains("0 0 1 RG", literal.ToString());
        TestAssert.Equal(1, LabelPaths(literal.ToString()));
        AssertEquivalentSvgPaint(literal, fallback, "An absent server uses its explicit fallback.");
        TestAssert.Equal(0, diagnostics.Count);

        const string stops = "<stop offset=\"0\" stop-color=\"blue\"/><stop offset=\"1\" stop-color=\"yellow\"/>";
        string Definition(string attributes) => "<defs><linearGradient id=\"present\" " + attributes + ">" + stops + "</linearGradient></defs>";
        var gradient = Paint(Svg(Definition("") + string.Format(rectangle, "url(#present)")));
        AssertEquivalentSvgPaint(gradient, Paint(Svg(Definition("") + string.Format(rectangle, "url(#present) red"))),
            "A usable server owns its paint even when a fallback is provided.");
        var unresolved = Paint(Svg(Definition("gradientTransform=\"bogus(1)\"") +
            string.Format(rectangle, "url(#present) red") + string.Format(rectangle, "url(#no-fallback)")), diagnostics);
        TestAssert.DoesNotContain("1 0 0 rg", unresolved.ToString());
        foreach (string id in new[] { "present", "no-fallback" })
            TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains(id, StringComparison.Ordinal)),
                "An unsupported present server and an absent server without fallback still diagnose.");
    }

    private static void AssertEquivalentSvgPaint(PdfGraphicsBuilder expected, PdfGraphicsBuilder actual, string reason)
    {
        TestAssert.Equal(expected.ToString(), actual.ToString());
        TestAssert.True(expected.ExtGStates.SequenceEqual(actual.ExtGStates), reason);
    }
}
