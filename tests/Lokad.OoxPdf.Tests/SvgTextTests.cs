using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class SvgTextTests
{
    private const string Label = "KEEP THIS LABEL";
    private const string LabelColor = "0.067 0.133 0.2 rg";
    private static string Svg(string contents) =>
        "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 400 160\" font-family=\"DejaVu Sans\">" + contents + "</svg>";

    public static void SvgLabelsRetainEveryGlyphInDocxAndPptxPictures()
    {
        string svg = Svg("<rect x=\"5\" y=\"5\" width=\"390\" height=\"150\" fill=\"#cce5ff\"/>" +
            "<g font-size=\"16\"><g font-size=\"150%\">" +
            "<text x=\"20\" y=\"80\" font-size=\"8\" style=\"font-size: inherit\" fill=\"#112233\">" + Label + "</text></g></g>");
        foreach (bool docx in new[] { true, false })
        foreach (bool review in docx ? new[] { false, true } : new[] { false })
        {
            string input = Input(docx, svg);
            var diagnostics = new List<OoxPdfDiagnostic>();
            using OoxPackage package = OoxPackage.Open(input, default);
            var resolver = new PortableRenderingTests.PinnedResolver();
            PdfPage page = docx
                ? new DocxRenderer(resolver, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final,
                    review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderPagesAsync(DocxTests.ReadDocx(input, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final),
                        diagnostics.Add, default).ToBlockingEnumerable().Single()
                : new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default), package,
                    diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Contains("0.8 0.898 1 rg", page.Content);
            TestAssert.Equal(13, LabelPaths(page.Content));
            TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"),
                "Ordinary labels must paint, rather than retain the former unsupported-text warning.");
            Write(page); // The real writer also checks balanced content and declared resources.
        }
    }

    public static void SvgLabelsShareInheritedFontSizeColorAndMeasuredAnchors()
    {
        double[] positions = new double[3];
        int index = 0;
        foreach (string anchor in new[] { "start", "middle", "end" })
        {
            string inherited = Svg("<g font-size=\"20\" fill=\"#112233\"><g font-size=\"150%\"><g font-size=\"inherit\">" +
                "<text x=\"200\" y=\"80\" font-size=\".5em\" text-anchor=\"" + anchor + "\">I</text></g></g></g>");
            var graphics = Paint(inherited);
            TestAssert.Equal(Paint(Svg("<text x=\"200\" y=\"80\" font-size=\"15\" fill=\"#112233\" text-anchor=\"" +
                anchor + "\">I</text>")).ToString(), graphics.ToString());
            Match move = Regex.Match(graphics.ToString(), @"(?m)^([-\d.]+) [-\d.]+ m\r?$");
            TestAssert.True(move.Success, "The label must contain a visible glyph contour.");
            positions[index++] = double.Parse(move.Groups[1].Value, CultureInfo.InvariantCulture);
        }
        var face = new PortableRenderingTests.PinnedResolver().Resolve(new FontRequest("DejaVu Sans"));
        var font = TestAssert.NotNull(FontProgramLoader.LoadAsync(face, default).GetAwaiter().GetResult());
        double width = font.GetAdvanceWidth(font.MapCodePoint('I')) * 15d / font.UnitsPerEm;
        TestAssert.True(Math.Abs(positions[0] - positions[1] - width / 2d) < .002d, "Middle anchors use the measured advance.");
        TestAssert.True(Math.Abs(positions[0] - positions[2] - width) < .002d, "End anchors use the measured advance.");
    }

    public static void SvgInitialFontSizePreservesLocalAndRelativeLabelsInBothFormats()
    {
        const string expected = """
            <text x="12" y="30" font-size="16" fill="#112233">INK</text>
            <text x="12" y="71" font-size="24" fill="#2254b4">NEXT</text>
            <text x="12" y="115" font-size="32" fill="green">STAY</text>
            """;
        const string source = """
            <g font-size="32">
              <text x="12" y="30" font-size="7" fill="#112233"
                    style="font-size:InItIaL !important; font-size:8px">INK</text>
              <g font-size="38" style="font-size:initial">
                <g font-size="150%">
                  <text x="12" y="71" font-size="7" fill="#2254b4" style="font-size:inherit">NEXT</text>
                </g>
              </g>
              <text x="12" y="115" fill="green">STAY</text>
            </g>
            """;
        AssertSvgPaintInBothFormats(Svg(expected), Svg(source), pdf =>
        {
            TestAssert.Equal(3, LabelPaths(pdf));
            TestAssert.Equal(4, LabelPaths(pdf, "0.133 0.329 0.706 rg"));
            TestAssert.Equal(4, LabelPaths(pdf, "0 0.502 0 rg"));
        });
    }

    public static void SvgLabelTransformsUseThePictureCoordinateSystem()
    {
        const string transform = "matrix(0.965925826 0.258819045 -0.258819045 0.965925826 5 7)";
        string text = "<text x=\"20\" y=\"80\" font-size=\"24\" fill=\"#112233\">I</text>";
        var group = Paint(Svg("<g transform=\"" + transform + "\">" + text + "</g>"));
        var own = Paint(Svg(text.Replace("<text ", "<text transform=\"" + transform + "\" ", StringComparison.Ordinal)));
        TestAssert.Equal(own.ToString(), group.ToString());
        Match matrix = Regex.Match(group.ToString(), @"(?m)^([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\r?$");
        TestAssert.True(matrix.Success, "The SVG transform must reach the glyph coordinate map.");
        double[] expected = [Math.Cos(Math.PI / 12d), -Math.Sin(Math.PI / 12d), Math.Sin(Math.PI / 12d), Math.Cos(Math.PI / 12d), 5d, 153d];
        for (int i = 0; i < expected.Length; i++)
            TestAssert.True(Math.Abs(double.Parse(matrix.Groups[i + 1].Value, CultureInfo.InvariantCulture) - expected[i]) < .000001d,
                $"Glyph/picture matrix component {i} must be {expected[i]:R}; got {matrix.Groups[i + 1].Value}.");
        TestAssert.Contains("0 0 400 160 re", group.ToString());
        TestAssert.Equal(0, group.StateDepth);
        TestAssert.Equal(1, LabelPaths(group.ToString()));
        const string compound = "translate(5 7) rotate(15)";
        TestAssert.Equal(group.ToString(), Paint(Svg("<g transform=\"" + compound + "\">" + text + "</g>")).ToString());
        TestAssert.Equal(group.ToString(), Paint(Svg(text.Replace("<text ", "<text transform=\"" + compound + "\" ", StringComparison.Ordinal))).ToString());
    }

    public static void SvgTransformFormsPreservePicturePlacement()
    {
        const string translation = """transform="translate(160 10)" """;
        const string scale = """transform="scale(3 2)" """;
        const string sibling = """<rect x="1" y="1" width="5" height="5" fill="#0055ff"/>""";
        string Shape(string transform = "") => $"<rect x=\"5\" y=\"10\" width=\"10\" height=\"10\" fill=\"#ff00aa\" {transform}/>";
        string matrix = Shape("""transform="matrix(3 0 0 2 160 10)" """);
        var placements = new (string Name, string Source)[]
        {
            ("list", Shape("""transform="translate(160 10) scale(3 2)" """)),
            ("nested groups", $"<g {translation}><g {scale}>{Shape()}</g></g>"),
            ("group and local", $"<g {translation}>{Shape(scale)}</g>")
        };

        foreach (bool docx in new[] { true, false })
        {
            var resolver = new PortableRenderingTests.PinnedResolver();
            byte[] Render(string source)
            {
                string path = Input(docx, Svg(source + sibling));
                try
                {
                    using var input = File.OpenRead(path);
                    using var output = new MemoryStream();
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    OoxPdfConverter.Convert(input, output, new OoxPdfOptions
                    {
                        InputKind = docx ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
                        FontResolver = resolver,
                        DiagnosticSink = diagnostics.Add
                    });
                    TestAssert.True(!diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT"), "These SVG transforms are supported.");
                    return output.ToArray();
                }
                finally { File.Delete(path); }
            }
            byte[] expected = Render(matrix);
            string pdf = Encoding.ASCII.GetString(expected);
            TestAssert.Contains("1 0 0.667 rg", pdf);
            TestAssert.Contains("0 0.333 1 rg", pdf);
            foreach (var (name, source) in placements)
            {
                TestAssert.True(Render(source).AsSpan().SequenceEqual(expected),
                    $"{(docx ? "DOCX" : "PPTX")}: a transform {name} must preserve the matrix control's full picture and sibling.");
            }
        }
    }

    public static void SvgRootOverflowClipsThePictureViewport()
    {
        const string shape = """<rect x="370" y="70" width="125" height="40" fill="#ff00ff"/>""";
        string Render(string attributes)
        {
            string svg = Svg(shape).Replace("viewBox=", attributes + " viewBox=", StringComparison.Ordinal);
            string path = Input(docx: false, svg);
            try
            {
                using OoxPackage package = OoxPackage.Open(path, default);
                var renderer = new PptxRenderer(new PortableRenderingTests.PinnedResolver());
                return renderer.RenderPagesAsync(new PptxReader().Read(package, default), package, null, default)
                    .ToBlockingEnumerable().Single().Content;
            }
            finally { File.Delete(path); }
        }

        string visible = Render("overflow=\"hidden\" style=\"overflow: visible\"");
        string clipped = Render("overflow=\"hidden\"");
        string[] visibleLines = visible.Split('\n');
        string clip = clipped.Split('\n').SingleOrDefault(line =>
            line.TrimEnd().EndsWith(" re W n", StringComparison.Ordinal) && !visibleLines.Contains(line)) ?? "";
        int clipIndex = clipped.IndexOf(clip, StringComparison.Ordinal);
        int paintIndex = clipped.IndexOf("1 0 1 rg", StringComparison.Ordinal);
        TestAssert.True(clip.Length > 0 && paintIndex > clipIndex, "The image viewport must clip authored geometry before it paints.");
        TestAssert.Equal(visible, clipped.Remove(clipIndex, clip.Length + 1));
        TestAssert.Equal(clipped, Render("overflow=\"visible\" style=\"overflow: scroll\""));
        TestAssert.Equal(visible, Render(""));
    }

    public static void SvgGradientTransformListsMatchExplicitMatrices()
    {
        foreach (var (gradient, geometry) in new[]
        {
            ("linearGradient", """x1="0" y1="0" x2="60" y2="0" """),
            ("radialGradient", """cx="15" cy="30" r="20" """)
        })
        {
            string Definition(string transform) => Svg($"""
                <defs><{gradient} id="g" gradientUnits="userSpaceOnUse" {geometry} gradientTransform="{transform}">
                  <stop offset="0" stop-color="red"/><stop offset="1" stop-color="blue"/>
                </{gradient}></defs>
                <rect width="250" height="130" fill="url(#g)"/>
                """);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var expected = Paint(Definition("matrix(2 0 0 2 40 10)"), diagnostics);
            var actual = Paint(Definition("translate(40 10) scale(2)"), diagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.True(expected.Shadings.Count > 0 || expected.ToString().Contains(" rg", StringComparison.Ordinal),
                "The reference must paint gradient colors or publish a shading.");
            AssertEquivalentSvgPaint(expected, actual, gradient);
            TestAssert.True(expected.Shadings.Select(resource => resource.Shading.ResourceKey)
                .SequenceEqual(actual.Shadings.Select(resource => resource.Shading.ResourceKey)),
                $"{gradient}: native shading coordinates follow the same transform order as sampled paint.");
        }
    }

    public static void SvgBezierGradientBoundsMatchExplicitUserCoordinates()
    {
        const string arch = "M10 70 C30 -10 70 -10 90 70 L90 90 H10 Z";
        const string asymmetric = "M10 80 C30 53 70 53 90 161 L90 180 H10 Z";
        const string boxLinear = """x1=".5" x2=".5" y1="0" y2="1" """;
        const string userLinear = """x1="50" x2="50" y1="10" y2="90" """;
        foreach (var (path, kind, box, user, paint) in new[]
        {
            // The arch reaches its upper edge at t=.5. The asymmetric curve
            // has a quadratic derivative and reaches its upper edge at t=1/3.
            (arch, "linearGradient", boxLinear, userLinear, """fill="url(#g)" """),
            (asymmetric, "linearGradient", boxLinear, """x1="50" x2="50" y1="65" y2="180" """, """fill="url(#g)" """),
            (arch, "radialGradient", """cx=".5" cy=".5" r=".5" """, """cx="50" cy="50" r="40" """, """fill="url(#g)" """),
            (arch, "linearGradient", boxLinear, userLinear, """fill="none" stroke="url(#g)" stroke-width="6" """)
        })
        {
            string Definition(string units, string geometry) => Svg($"""
                <defs><{kind} id="g" gradientUnits="{units}" {geometry}>
                  <stop offset="0" stop-color="red"/><stop offset="1" stop-color="blue"/>
                </{kind}></defs>
                <path d="{path}" {paint}/>
                """);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfGraphicsBuilder expected = Paint(Definition("userSpaceOnUse", user), diagnostics);
            PdfGraphicsBuilder actual = Paint(Definition("objectBoundingBox", box), diagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.Contains(" c", actual.ToString());
            TestAssert.True(actual.Shadings.Count > 0 || actual.ShadingPatterns.Count > 0 ||
                actual.ToString().Contains(" rg", StringComparison.Ordinal), "The path must paint its gradient.");
            AssertEquivalentSvgPaint(expected, actual, $"{kind} {paint}");
            TestAssert.True(expected.Shadings.Select(resource => resource.Shading.ResourceKey)
                .SequenceEqual(actual.Shadings.Select(resource => resource.Shading.ResourceKey)),
                "Matching curve extents must publish the same shading coordinates.");
            TestAssert.True(expected.ShadingPatterns.Select(resource => resource.Pattern.ResourceKey)
                .SequenceEqual(actual.ShadingPatterns.Select(resource => resource.Pattern.ResourceKey)),
                "Matching curve extents must publish the same pattern projection.");
        }
    }

    public static void SvgLabelOpacityGroupsRemainVectorOnlyAndKeepUnsupportedSiblingsExplicit()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var graphics = Paint(Svg("<rect width=\"400\" height=\"160\" fill=\"#cce5ff\"/>" +
            "<g opacity=\"0.5\"><text x=\"20\" y=\"80\" fill=\"#112233\">I</text></g>" +
            "<text><tspan dx=\"10\">unsupported positioned text</tspan></text><text fill=\"url(#missing)\">unsupported gradient</text>"), diagnostics);
        TestAssert.Equal(1, graphics.Groups.Count);
        TestAssert.Equal(1, LabelPaths(graphics.Groups[0].Group.Content));
        TestAssert.True(!graphics.Groups[0].Group.Content.Contains(" BT", StringComparison.Ordinal) &&
            !graphics.Groups[0].Group.Content.Contains(" Tf", StringComparison.Ordinal), "Isolated labels borrow paths without unbound fonts.");
        TestAssert.True(diagnostics.Any(d => d.Fallback == "Partial" && d.Message.Contains("2 unsupported text elements", StringComparison.Ordinal)),
            "Unsupported text features must not be declared supported because a sibling label painted.");
        TestAssert.True(diagnostics.Any(d => d.Message.Contains("tspan", StringComparison.Ordinal)), "Positioned spans retain their diagnostic.");
        TestAssert.Equal(0, graphics.UsedFontResourceNames.Count);
        Write(Page(graphics));
    }

    public static void SvgLabelsPreserveWhitespaceAndDoNotRevealHiddenText()
    {
        string Text(string value, string extra = "") => Svg("<text x=\"20\" y=\"80\" fill=\"#112233\" " + extra + ">" + value + "</text>");
        TestAssert.Equal(Paint(Text("AB")).ToString(), Paint(Text("A\nB")).ToString());
        TestAssert.Equal(Paint(Text("A B", "xml:space=\"preserve\"")).ToString(), Paint(Text("A\nB", "xml:space=\"preserve\"")).ToString());
        foreach (string hidden in new[] { "display=\"none\"", "visibility=\"hidden\"", "opacity=\"0\"" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var graphics = Paint(Svg("<g " + hidden + "><text><tspan>HIDDEN</tspan></text></g>"), diagnostics);
            TestAssert.Equal(0, Regex.Matches(graphics.ToString(), @"(?m)^f\r?$").Count);
            TestAssert.Equal(0, diagnostics.Count);
        }
    }

    public static void SvgLabelsKeepTextAcrossPlainAndMetadataSpans()
    {
        string Picture(string name, string initials) => Svg($"""
            <text x="20" y="45" font-size="20" fill="#112233">{name}</text>
            <text x="20" y="80" font-size="20" fill="#114499">{initials}</text>
            <rect x="120" y="90" width="20" height="10" fill="lime"/>
            """);
        AssertSvgPaintInBothFormats(Picture("NORTH", "A B"),
            Picture("""NO<tspan id="direction">R<tspan>T</tspan></tspan>H""",
                """A <tspan class="caption" data-key="initials" role="presentation" aria-hidden="true"> B</tspan>"""), pdf =>
            {
                TestAssert.Equal(5, LabelPaths(pdf));
                TestAssert.Equal(2, LabelPaths(pdf, "0.067 0.267 0.6 rg"));
            }, "0 1 0 rg");
    }

    public static void SvgLabelsPrepareFontsOnlyForVisiblePictureUses()
    {
        const string family = "LABEL-FONT";
        string svg = Svg($"<text x=\"20\" y=\"80\" font-family=\"{family}\" font-weight=\"bold\" fill=\"#112233\">{Label}</text>");
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var (transparent, secondVisible) in new[] { (false, false), (true, false), (true, true) })
        {
            string input = Input(false, svg, picture =>
            {
                if (secondVisible)
                {
                    var visible = new XElement(picture);
                    visible.Descendants(p + "cNvPr").Single().SetAttributeValue("id", "1000");
                    visible.Descendants(a + "off").First().SetAttributeValue("x", "4572000");
                    picture.AddAfterSelf(visible);
                }
                if (transparent)
                    picture.Descendants(a + "blip").Single().AddFirst(new XElement(a + "alphaModFix", new XAttribute("amt", 0)));
            });
            var diagnostics = new List<OoxPdfDiagnostic>();
            int requests = 0;
            var resolver = new ObservingResolver(request => { if (request.FamilyName == family) requests++; });
            using OoxPackage package = OoxPackage.Open(input, default);
            PdfPage page = new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default),
                package, diagnostics.Add, default).ToBlockingEnumerable().Single();

            bool paintsLabel = !transparent || secondVisible;
            TestAssert.Equal(paintsLabel ? 1 : 0, requests);
            TestAssert.Equal(paintsLabel ? 1 : 0, diagnostics.Count(d => d.Fallback == "SubstitutedFont" &&
                d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains(family, StringComparison.Ordinal)));
            if (paintsLabel)
                TestAssert.Equal(13, LabelPaths(page.Content));
            else
                TestAssert.DoesNotContain(LabelColor, page.Content);
            Write(page);
        }
    }

    public static void SvgLabelsKeepHarmlessAttributesAndNonRenderingMetadata()
    {
        string text = "<text x=\"20\" y=\"80\" fill=\"#112233\">" + Label + "</text>";
        string plain = Paint(Svg(text)).ToString();
        foreach (string changed in new[] {
            Svg(text.Replace("<text ", "<text class=\"source-label\" id=\"caption\" data-source=\"public\" ", StringComparison.Ordinal)),
            Svg("<g class=\"diagram\">" + text + "</g>"),
            Svg(text).Replace("<svg ", "<svg class=\"diagram\" ", StringComparison.Ordinal),
            Svg(text.Replace("</text>", "<title>NOT A LABEL</title><desc>NOT A LABEL</desc><metadata><style>.unused{fill:red}</style><rect width=\"400\" height=\"160\"/></metadata></text>", StringComparison.Ordinal)) })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var graphics = Paint(changed, diagnostics);
            TestAssert.True(plain == graphics.ToString(), "Non-painting SVG metadata must preserve every visible label path.");
            TestAssert.Equal(13, LabelPaths(graphics.ToString()));
            TestAssert.Equal(0, diagnostics.Count);
        }
        foreach (string stylesheet in new[] {
            Svg("<style>.source-label{fill:red}</style>" + text),
            "<?xml-stylesheet href=\"unavailable.css\" type=\"text/css\"?>" + Svg(text) })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var graphics = Paint(stylesheet, diagnostics);
            TestAssert.True(diagnostics.Any(d => d.Fallback == "Partial" && d.Message.Contains("text", StringComparison.Ordinal)),
                "An actual unsupported stylesheet remains distinct from harmless class metadata.");
            TestAssert.True(!graphics.ToString().Contains(LabelColor, StringComparison.Ordinal), "Do not claim stylesheet-driven glyph paint without implementing its styles.");
        }
    }

    public static void SvgLabelsRejectUnsupportedPositioningAndNonFinitePaintWithoutLosingShapes()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var graphics = Paint(Svg("<rect width=\"400\" height=\"160\" fill=\"#cce5ff\"/>" +
            "<text x=\"10 20\">array positions</text><text dx=\"10%\">viewport percentage</text>" +
            "<text transform=\"matrix(1e308 0 0 1e308 1e308 0)\">overflow</text>" +
            "<text x=\"20\" y=\"80\" fill=\"#112233\">I</text>"), diagnostics);
        TestAssert.Contains("0.8 0.898 1 rg", graphics.ToString());
        TestAssert.Equal(1, LabelPaths(graphics.ToString()));
        TestAssert.True(diagnostics.Any(d => d.Message.Contains("3 unsupported text elements", StringComparison.Ordinal)),
            "Unimplemented positions and non-finite mapped outlines must diagnose while supported siblings survive.");
        TestAssert.Equal(0, graphics.StateDepth);
        Write(Page(graphics));
    }

    public static void SvgLabelFontPreparationRefundsScratchOnRefusalAndCancellation()
    {
        byte[] bytes = Encoding.UTF8.GetBytes(Svg("<text>I</text>"));
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after == 0) return false;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            bool loaded = false;
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                void Prepare() => PptxRenderer.PrepareSvgTextFontsAsync([bytes], new PortableRenderingTests.PinnedResolver(),
                    (_, _) => { loaded = true; return ValueTask.FromResult<OpenTypeFont?>(null); }, null, cancellation.Token).GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Prepare).CancellationToken);
                else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Prepare).Kind);
                TestAssert.True(!loaded, "Preparation refusal/cancellation precedes font acquisition.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void SvgLabelPreparationFailurePreservesDestinationAndCleansTemporaryFiles()
    {
        foreach (bool docx in new[] { true, false })
        foreach (bool cancel in new[] { false, true })
        {
            string svg = Svg("<text x=\"20\" y=\"80\" font-family=\"LABEL-FONT\">" + Label + "</text>");
            string input = Input(docx, svg);
            string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "svg-label-publication-" + Guid.NewGuid().ToString("N"))).FullName;
            string output = Path.Combine(directory, "unchanged.pdf");
            byte[] prior = "UNCHANGED DESTINATION"u8.ToArray(); File.WriteAllBytes(output, prior);
            using var cancellation = new CancellationTokenSource();
            bool requested = false;
            var resolver = new ObservingResolver(request => {
                if (request.FamilyName != "LABEL-FONT") return;
                requested = true;
                if (cancel) cancellation.Cancel();
            });
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => !cancel && requested && after != 0 };
            var options = new OoxPdfOptions { FontResolver = resolver, MemoryBudget = pool };
            void Convert() => OoxPdfConverter.Convert(input, output, options, cancellation.Token);
            if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Convert).CancellationToken);
            else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Convert).Kind);
            TestAssert.True(requested, "The failure must reach the SVG label's own font preparation.");
            TestAssert.True(prior.AsSpan().SequenceEqual(File.ReadAllBytes(output)), "Label failure cannot publish partial diagram output.");
            TestAssert.Equal(1, Directory.GetFiles(directory).Length);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    private sealed class ObservingResolver(Action<FontRequest> observe) : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver inner = new();
        public FontFaceResolution Resolve(FontRequest request) { observe(request); return inner.Resolve(request); }
    }

    private static PdfGraphicsBuilder Paint(string svg, List<OoxPdfDiagnostic>? diagnostics = null)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(svg);
        var fonts = PptxRenderer.PrepareSvgTextFontsAsync([bytes], new PortableRenderingTests.PinnedResolver(),
            FontProgramLoader.LoadAsync, null, default).GetAwaiter().GetResult();
        var graphics = new PdfGraphicsBuilder();
        TestAssert.True(PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 400, 160, bytes,
            diagnostics is null ? null : diagnostics.Add, 1, "/test.svg", default, fonts), "SVG must remain paintable.");
        return graphics;
    }

    private static int LabelPaths(string content, string color = LabelColor)
    {
        int start = content.IndexOf(color, StringComparison.Ordinal);
        TestAssert.True(start >= 0, "The authored label color must reach visible paths.");
        int end = content.IndexOf("\nQ", start, StringComparison.Ordinal);
        return Regex.Matches(content[start..(end < 0 ? content.Length : end)], @"(?m)^f\r?$").Count;
    }

    private static PdfPage Page(PdfGraphicsBuilder graphics) => new(400, 160, graphics.ToString(), [], [],
        graphics.ExtGStates, graphics.Shadings, graphics.Patterns, [], groups: graphics.Groups, shadingPatterns: graphics.ShadingPatterns);

    private static void Write(PdfPage page)
    {
        using var stream = new MemoryStream();
        PdfDocumentWriter.WriteStagedAsync(stream, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), new OoxConversionLimits(), null,
            default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
        TestAssert.True(stream.Length > 0, "The writer must accept the actual label content and group resources.");
    }

    private static string Input(bool docx, string svg, Action<XElement>? editPptxPicture = null)
    {
        string file = docx ? "docx-svg-inline.docx" : "pptx-svg-extensions.pptx";
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", file));
        using var archive = ZipFile.OpenRead(source);
        var parts = archive.Entries.ToDictionary(e => e.FullName, e => {
            using var stream = e.Open(); using var bytes = new MemoryStream(); stream.CopyTo(bytes); return bytes.ToArray();
        });
        foreach (string name in parts.Keys.Where(name => name.EndsWith(".svg", StringComparison.Ordinal)).ToArray())
            parts[name] = Encoding.UTF8.GetBytes(svg);
        if (!docx)
        {
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
            var first = xml.Descendants(p + "pic").First();
            foreach (var sibling in first.Parent!.Elements().Where(e => e.Name == p + "pic" && e != first).ToArray()) sibling.Remove();
            editPptxPicture?.Invoke(first);
            parts["ppt/slides/slide1.xml"] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        }
        return TestFixtures.WriteTempPackage(docx ? ".docx" : ".pptx", parts);
    }
}
