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

    public static void SvgConstantLinearGradientsMatchSolidPaintInBothFormats()
    {
        const string definitions = """
            <defs>
              <linearGradient id="zero" gradientUnits="userSpaceOnUse" x1="80px" y1="24" x2="20%" y2="15%">
                <stop offset=".8" stop-color="red" stop-opacity=".75"/>
                <stop offset=".2" style="stop-color:blue;stop-opacity:.5"/>
              </linearGradient>
              <linearGradient id="one"><stop offset=".4" stop-color="green" stop-opacity=".6"/></linearGradient>
            </defs>
            """;
        string Picture(string fill, string constantStroke, string singleStopStroke) => Svg($"""
            {definitions}
            <rect width="400" height="160" fill="yellow"/>
            <rect x="20" y="20" width="60" height="60" {fill}/>
            <rect x="105" y="20" width="60" height="60" fill="none" stroke-width="8" {constantStroke}/>
            <rect x="190" y="20" width="60" height="60" fill="none" stroke-width="8" {singleStopStroke}/>
            """);
        string expected = Picture("fill=\"blue\" fill-opacity=\".4\"",
            "stroke=\"blue\" stroke-opacity=\".4\"", "stroke=\"green\" stroke-opacity=\".3\"");
        string actual = Picture("fill=\"url(#zero)\" fill-opacity=\".8\"",
            "stroke=\"url(#zero)\" stroke-opacity=\".8\"", "stroke=\"url(#one)\" stroke-opacity=\".5\"");
        AssertSvgPaintInBothFormats(expected, actual, "0 0 1 rg", "0 0 1 RG", "0 0.502 0 RG");
    }

    public static void SvgSingleStopRadialGradientsMatchSolidPaintInBothFormats()
    {
        string Picture(string padFill, string repeatFill, string stroke) => Svg($"""
            <defs>
              <radialGradient id="pad" r="30%"><stop offset=".2" style="stop-color:blue;stop-opacity:.5"/></radialGradient>
              <radialGradient id="repeat" r="15%" spreadMethod="repeat">
                <stop offset=".4" stop-color="green" stop-opacity=".6"/>
              </radialGradient>
            </defs>
            <rect width="400" height="160" fill="yellow"/>
            <rect x="20" y="20" width="60" height="60" {padFill}/>
            <rect x="105" y="20" width="60" height="60" {repeatFill}/>
            <rect x="190" y="20" width="60" height="60" fill="none" stroke-width="8" {stroke}/>
            """);
        string expected = Picture("fill=\"blue\" fill-opacity=\".4\"",
            "fill=\"green\" fill-opacity=\".3\"", "stroke=\"blue\" stroke-opacity=\".4\"");
        string actual = Picture("fill=\"url(#pad)\" fill-opacity=\".8\"",
            "fill=\"url(#repeat)\" fill-opacity=\".5\"", "stroke=\"url(#pad)\" stroke-opacity=\".8\"");
        AssertSvgPaintInBothFormats(expected, actual, "0 0 1 rg", "0 0.502 0 rg", "0 0 1 RG");
    }

    public static void SvgQuotedLocalPaintServersMatchUnquotedReferencesInBothFormats()
    {
        string Picture(string linear, string stroke, string radial, string missing) => Svg($"""
            <defs>
              <linearGradient id="linear"><stop offset="0" stop-color="red"/><stop offset="1" stop-color="blue"/></linearGradient>
              <radialGradient id="radial"><stop offset="0" stop-color="yellow"/><stop offset="1" stop-color="purple"/></radialGradient>
            </defs>
            <rect x="10" y="20" width="60" height="60" fill="{linear} lime"/>
            <rect x="85" y="20" width="60" height="60" fill="none" style="stroke: {stroke}; stroke-width: 8"/>
            <rect x="160" y="20" width="60" height="60" fill="{radial}"/>
            <rect x="235" y="20" width="60" height="60" fill="{missing} green"/>
            <text x="20" y="130" font-size="36" fill="{missing} #112233">I</text>
            """);

        string expected = Picture("url(#linear)", "url(#linear)", "url(#radial)", "url(#missing)");
        string actual = Picture("url('#linear')", "URL( &quot;#linear&quot; )", "url( #radial )", "url( '#missing' )");
        AssertSvgPaintInBothFormats(expected, actual, "/PatternType 2", "/ShadingType 3", "0 0.502 0 rg", LabelColor);
    }

    public static void SvgDecreasingStopsKeepNativePaintAndOpacityInBothFormats()
    {
        string Picture(string last) => Svg($"""
            <defs>
              <linearGradient id="linear" x1="20%" x2="80%">
                <stop offset="0" stop-color="black"/><stop offset="0" stop-color="red"/>
                <stop offset=".75" stop-color="green"/><stop offset="{last}" stop-color="blue"/>
                <stop offset="1" stop-color="blue"/><stop offset="1" stop-color="black"/>
              </linearGradient>
              <radialGradient id="radial" r="35%">
                <stop offset="0" stop-color="red"/>
                <stop offset=".75" stop-color="green" stop-opacity=".45"/><stop offset="{last}" stop-color="blue"/>
                <stop offset="1" stop-color="blue"/><stop offset="1" stop-color="black"/>
              </radialGradient>
            </defs>
            <rect width="400" height="160" fill="yellow"/>
            <rect x="10" y="20" width="65" height="60" fill="url(#linear)"/>
            <rect x="85" y="20" width="65" height="60" fill="none" stroke="url(#linear)" stroke-width="8"/>
            <rect x="160" y="20" width="65" height="60" fill="url(#radial)"/>
            """);

        AssertSvgPaintInBothFormats(Picture(".75"), Picture(".25"), "/PatternType 2", "/ShadingType 3",
            "/S /Luminosity", "/Coords [-1 0 2 0]", "/Bounds [0.375 0.5]");
    }

    public static void SvgGradientTemplatesMatchExpandedPaintInBothFormats()
    {
        const string stops = "<stop offset=\"0\" stop-color=\"currentColor\"/><stop offset=\"1\" stop-color=\"blue\"/>";
        const string alphaStops = "<stop offset=\"0\" stop-color=\"green\" stop-opacity=\".35\"/><stop offset=\"1\" stop-color=\"yellow\"/>";
        string Picture(string definitions) => Svg($"""
            <defs>{definitions}</defs>
            <rect width="400" height="160" fill="#fff0a0"/>
            <rect x="10" y="20" width="60" height="60" fill="url(#linear)"/>
            <rect x="85" y="20" width="60" height="60" fill="none" stroke="url(#stroke)" stroke-width="8"/>
            <rect x="160" y="20" width="60" height="60" fill="url(#shared)"/>
            <rect x="235" y="20" width="60" height="60" fill="url(#radial)"/>
            """);

        string expected = Picture($"""
            <linearGradient id="linear" color="red" gradientUnits="userSpaceOnUse" x1="15" x2="42" spreadMethod="reflect"
                gradientTransform="translate(5 0)">{stops}</linearGradient>
            <linearGradient id="stroke" color="red">{stops}</linearGradient>
            <radialGradient id="shared" color="red" cx="45%" cy="55%" r="40%">{stops}</radialGradient>
            <radialGradient id="radial" cx="45%" cy="55%" r="40%">{alphaStops}</radialGradient>
            """);
        string actual = Picture($"""
            <linearGradient id="linear" href="#axis" x2="42" color="green"><title>Local override</title></linearGradient>
            <linearGradient id="axis" href="#palette" gradientUnits="userSpaceOnUse" x1="15" x2="70"
                gradientTransform="translate(5 0)" spreadMethod="reflect"/>
            <linearGradient id="stroke" xmlns:xlink="http://www.w3.org/1999/xlink" xlink:href="#palette"/>
            <radialGradient id="shared" href="#palette" cx="45%" cy="55%" r="40%"/>
            <radialGradient id="radial" href="#shared">{alphaStops}</radialGradient>
            <linearGradient id="palette" color="red">{stops}</linearGradient>
            """);

        AssertSvgPaintInBothFormats(expected, actual, "/PatternType 2", "/ShadingType 3", "/S /Luminosity");
    }

    public static void SvgGradientTemplateCyclesAndMissingTargetsKeepLocalDeclarations()
    {
        const string definitions = """
            <defs xmlns:xlink="http://www.w3.org/1999/xlink">
              <linearGradient id="cycle-with-paint" href="#cycle"><stop stop-color="red"/></linearGradient>
              <linearGradient id="cycle" href="#cycle-with-paint"/>
              <linearGradient id="outside" href="#cycle-with-paint"/>
              <linearGradient id="relative" href="separate.svg#paint"/>
              <linearGradient id="missing" href="#absent" xlink:href="#cycle-with-paint"/>
              <linearGradient id="discarded" href="#cycle-with-paint"><stop stop-color="unknown-color"/></linearGradient>
              <linearGradient id="self" href="#self"><stop stop-color="blue"/></linearGradient>
            </defs>
            """;
        const string shapes = """
            <rect width="20" height="20" fill="url(#outside)"/>
            <rect x="25" width="20" height="20" fill="url(#self)"/>
            <rect x="50" width="20" height="20" fill="url(#cycle)"/>
            <rect x="75" width="20" height="20" fill="url(#relative)"/>
            <rect x="100" width="20" height="20" fill="url(#missing)"/>
            <rect x="125" width="20" height="20" fill="url(#discarded)"/>
            """;
        var diagnostics = new List<OoxPdfDiagnostic>();
        var actual = Paint(Svg(definitions + shapes), diagnostics);
        var expected = Paint(Svg("""
            <rect width="20" height="20" fill="red"/>
            <rect x="25" width="20" height="20" fill="blue"/>
            """));
        AssertEquivalentSvgPaint(expected, actual, "The usable local palette survives a malformed reference.");
        TestAssert.Equal(4, diagnostics.Count);
        foreach (string id in new[] { "cycle", "relative", "missing", "discarded" })
        {
            TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" &&
                d.Message.Contains($"gradient {id}.", StringComparison.Ordinal)), $"Expected fallback for {id}.");
        }
    }

    public static void SvgGradientTemplateChainsPreserveStopsAndOpacityDiagnostics()
    {
        const int chainLength = 1024;
        const string palette = "<linearGradient id=\"end\" x2=\".6\"><stop offset=\"0\" stop-color=\"red\" stop-opacity=\"50%\"/>" +
            "<stop offset=\"1\" stop-color=\"blue\"/></linearGradient>";
        const string picture = "<rect width=\"60\" height=\"60\" fill=\"url(#end)\"/>";
        var chain = new StringBuilder();
        for (int i = 0; i < chainLength; i++)
        {
            string target = i + 1 == chainLength ? "end" : $"g{i + 1}";
            chain.Append($"<linearGradient id=\"g{i}\" href=\"#{target}\"/>");
        }
        var expectedDiagnostics = new List<OoxPdfDiagnostic>();
        var actualDiagnostics = new List<OoxPdfDiagnostic>();
        var expected = Paint(Svg($"<defs>{palette}</defs>{picture}"), expectedDiagnostics);
        var actual = Paint(Svg($"<defs>{chain}{palette}</defs>{picture.Replace("#end", "#g0", StringComparison.Ordinal)}"), actualDiagnostics);
        AssertEquivalentSvgPaint(expected, actual, "Deep chains should share the authored gradient palette.");
        TestAssert.Contains("ignores stop-opacity on 1 gradient stop", expectedDiagnostics.Single().Message);
        TestAssert.True(expectedDiagnostics.SequenceEqual(actualDiagnostics), "Used template stops retain the same opacity diagnostic.");
        Write(Page(actual));
    }

    private static void AssertSvgPaintInBothFormats(string expected, string actual, params string[] expectedMarkers)
    {
        foreach (var (kind, view, geometry) in new[] {
            (OoxPdfInputKind.Pptx, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
            (OoxPdfInputKind.Docx, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
            (OoxPdfInputKind.Docx, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var options = new OoxPdfOptions
            {
                InputKind = kind, DocxMarkupMode = view, DocxMarkupGeometryMode = geometry,
                FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = diagnostics.Add
            };
            byte[] Render(string svg)
            {
                using var input = File.OpenRead(Input(kind == OoxPdfInputKind.Docx, svg));
                using var output = new MemoryStream();
                OoxPdfConverter.Convert(input, output, options);
                return output.ToArray();
            }

            byte[] control = Render(expected);
            string content = Encoding.Latin1.GetString(control);
            foreach (string marker in expectedMarkers)
            {
                TestAssert.Contains(marker, content);
            }
            TestAssert.True(control.AsSpan().SequenceEqual(Render(actual)),
                $"{kind} {view} must emit the same paint and resources as the control.");
            TestAssert.Equal(0, diagnostics.Count);
        }
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
