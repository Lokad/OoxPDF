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

internal static class SvgTextTests
{
    private const string Label = "KEEP THIS LABEL";
    private const string LabelColor = "0.067 0.133 0.2 rg";
    private static string Svg(string contents) =>
        "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 400 160\" font-family=\"DejaVu Sans\">" + contents + "</svg>";

    public static void SvgLabelsRetainEveryGlyphInDocxAndPptxPictures()
    {
        string svg = Svg("<rect x=\"5\" y=\"5\" width=\"390\" height=\"150\" fill=\"#cce5ff\"/>" +
            "<text x=\"20\" y=\"80\" font-size=\"24\" fill=\"#112233\">" + Label + "</text>");
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
            string inherited = Svg("<g font-size=\"20\" fill=\"#112233\"><g font-size=\"150%\">" +
                "<text x=\"200\" y=\"80\" font-size=\".5em\" text-anchor=\"" + anchor + "\">I</text></g></g>");
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

    public static void SvgLabelTransformsUseThePictureCoordinateSystem()
    {
        // An explicit matrix isolates glyph/picture axes from the existing
        // path painter's separately unqualified compound-list ordering.
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
        TestAssert.Equal(Paint(Svg("<g transform=\"" + compound + "\">" + text + "</g>")).ToString(),
            Paint(Svg(text.Replace("<text ", "<text transform=\"" + compound + "\" ", StringComparison.Ordinal))).ToString());
    }

    public static void SvgLabelOpacityGroupsRemainVectorOnlyAndKeepUnsupportedSiblingsExplicit()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var graphics = Paint(Svg("<rect width=\"400\" height=\"160\" fill=\"#cce5ff\"/>" +
            "<g opacity=\"0.5\"><text x=\"20\" y=\"80\" fill=\"#112233\">I</text></g>" +
            "<text><tspan>unsupported nested text</tspan></text><text fill=\"url(#missing)\">unsupported gradient</text>"), diagnostics);
        TestAssert.Equal(1, graphics.Groups.Count);
        TestAssert.Equal(1, LabelPaths(graphics.Groups[0].Group.Content));
        TestAssert.True(!graphics.Groups[0].Group.Content.Contains(" BT", StringComparison.Ordinal) &&
            !graphics.Groups[0].Group.Content.Contains(" Tf", StringComparison.Ordinal), "Isolated labels borrow paths without unbound fonts.");
        TestAssert.True(diagnostics.Any(d => d.Fallback == "Partial" && d.Message.Contains("2 unsupported text elements", StringComparison.Ordinal)),
            "Unsupported text features must not be declared supported because a sibling label painted.");
        TestAssert.True(diagnostics.Any(d => d.Message.Contains("tspan", StringComparison.Ordinal)), "Nested text remains diagnosed.");
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

    private static int LabelPaths(string content)
    {
        int start = content.IndexOf(LabelColor, StringComparison.Ordinal);
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

    private static string Input(bool docx, string svg)
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
            parts["ppt/slides/slide1.xml"] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        }
        return TestFixtures.WriteTempPackage(docx ? ".docx" : ".pptx", parts);
    }
}
