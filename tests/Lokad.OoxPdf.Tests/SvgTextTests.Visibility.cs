using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class SvgTextTests
{
    public static void SvgUnmappedHiddenCharactersKeepVisibleLabelPaint()
    {
        const string missing = "\U00013000";
        const string hiddenControl = "\u0250";
        var face = new PortableRenderingTests.PinnedResolver().Resolve(new FontRequest("DejaVu Sans"));
        var font = TestAssert.NotNull(FontProgramLoader.LoadAsync(face, default).GetAwaiter().GetResult());
        TestAssert.Equal((ushort)0, font.MapCodePoint(char.ConvertToUtf32(missing, 0)));
        ushort controlGlyph = font.MapCodePoint(hiddenControl[0]);
        TestAssert.True(controlGlyph != 0 && font.GetAdvanceWidth(0) > 0, "The control must stand for a positive missing-glyph advance.");
        TestAssert.Equal(font.GetAdvanceWidth(0), font.GetAdvanceWidth(controlGlyph));

        static string Picture(string hidden) => Svg($"""
            <text x="20" y="40" font-size="24" font-kerning="none" fill="#112233">NOR<tspan visibility="hidden">{hidden}</tspan>TH</text>
            <text x="160" y="80" font-size="24" font-kerning="none" text-anchor="middle" fill="#2254b4">W<tspan opacity="0">{hidden}</tspan>EST</text>
            <rect x="200" y="100" width="20" height="10" fill="lime"/>
            """);

        AssertSvgPaintInBothFormats(Picture(hiddenControl), Picture(missing), pdf =>
        {
            TestAssert.Equal(5, LabelPaths(pdf));
            TestAssert.Equal(4, LabelPaths(pdf, "0.133 0.329 0.706 rg"));
        }, "0 1 0 rg");

        var diagnostics = new List<OoxPdfDiagnostic>();
        string visible = Paint(Svg($"""<text x="20" y="40" fill="#112233">NOR{missing}TH</text><rect width="10" height="10" fill="lime"/>"""), diagnostics).ToString();
        TestAssert.DoesNotContain(LabelColor, visible);
        TestAssert.Contains("0 1 0 rg", visible);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" &&
            d.Message.Contains("unsupported text element", StringComparison.Ordinal)), "An unmapped visible glyph keeps the existing diagnostic.");
    }

    public static void SvgHiddenGraphicsShareTextVisibilityAdmission()
    {
        string[] primitives = [
            "<path d=\"M10 10 H90 V90 H10 Z\"/>", "<rect x=\"10\" y=\"10\" width=\"80\" height=\"80\"/>",
            "<circle cx=\"50\" cy=\"50\" r=\"40\"/>", "<ellipse cx=\"50\" cy=\"50\" rx=\"40\" ry=\"20\"/>",
            "<line x1=\"10\" y1=\"10\" x2=\"90\" y2=\"90\" stroke=\"red\"/>",
            "<polyline points=\"10,10 90,10 90,90\"/>", "<polygon points=\"10,10 90,10 90,90\"/>"
        ];
        string empty = Paint(Svg("")).ToString();
        foreach (string primitive in primitives)
        foreach (string declaration in new[] { "display=\"none\"", "visibility=\"hidden\"", "style=\"display:none\"", "visibility=\"collapse\"" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            TestAssert.Equal(empty, Paint(Svg(primitive.Insert(primitive.IndexOf(' '), " " + declaration)), diagnostics).ToString());
            TestAssert.Equal(0, diagnostics.Count);
        }
    }

    public static void SvgVisibilityOverridesKeepDisplayAndInheritanceDistinct()
    {
        const string rectangle = "<rect width=\"80\" height=\"60\" fill=\"red\"/>";
        string visible = Paint(Svg(rectangle)).ToString(), empty = Paint(Svg("")).ToString();
        foreach (var sample in new[] {
            (Body: "<g visibility=\"hidden\">" + rectangle + "</g>", Expected: empty),
            (Body: "<g visibility=\"hidden\"><g visibility=\"inherit\">" +
                rectangle.Replace("<rect ", "<rect visibility=\"visible\" style=\"visibility: inherit\" ") + "</g></g>", Expected: empty),
            (Body: "<g visibility=\"hidden\"><g visibility=\"visible\">" + rectangle + "</g></g>", Expected: visible),
            (Body: "<g display=\"none\"><g visibility=\"visible\">" + rectangle + "</g></g>", Expected: empty),
            (Body: rectangle.Replace("<rect ", "<rect display=\"none\" style=\"display:block\" "), Expected: visible),
            (Body: rectangle.Replace("<rect ", "<rect visibility=\"hidden\" style=\"visibility:visible\" "), Expected: visible),
            (Body: rectangle.Replace("<rect ", "<rect visibility=\"visible\" style=\"visibility:hidden\" "), Expected: empty)
        }) TestAssert.Equal(sample.Expected, Paint(Svg(sample.Body)).ToString());
    }

    public static void SvgVisibilitySkipsHiddenFontsAndUnsupportedPayload()
    {
        string svg = Svg("<g display=\"none\"><rect transform=\"unknown(1)\" width=\"80\" height=\"60\"/>" +
            "<use href=\"#unused\"/><text font-family=\"HIDDEN-FONT\"><tspan>Hidden</tspan></text></g>" +
            "<g visibility=\"hidden\"><text visibility=\"inherit\" font-family=\"HIDDEN-FONT\">Also hidden</text></g>");
        var diagnostics = new List<OoxPdfDiagnostic>();
        bool resolved = false, loaded = false;
        var bytes = Encoding.UTF8.GetBytes(svg);
        var fonts = PptxRenderer.PrepareSvgTextFontsAsync([bytes], new PortableRenderingTests.ObservingResolver(_ => resolved = true),
            (_, _) => { loaded = true; return ValueTask.FromResult<OpenTypeFont?>(null); }, diagnostics.Add, default).GetAwaiter().GetResult();
        var graphics = new PdfGraphicsBuilder();
        TestAssert.True(PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 400, 160, bytes, diagnostics.Add, 1, "public.svg", default, fonts), "Hidden payload is a valid empty picture.");
        TestAssert.True(!resolved && !loaded && fonts.Count == 0, "Hidden text cannot acquire font programs.");
        TestAssert.Equal(Paint(Svg("")).ToString(), graphics.ToString());
        TestAssert.Equal(0, diagnostics.Count);
        Paint(Svg("<use href=\"#unused\"/>"), diagnostics);
        TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("use", StringComparison.Ordinal)),
            "A visible unsupported element retains its diagnostic.");
    }

    public static void SvgNonDisplayedTextPathsLeaveTheVisibleLabelIntact()
    {
        static string Picture(string child = "") => Svg($"""
            <defs><path id="track" d="M20 45 H185"/></defs>
            <text x="100" y="82" text-anchor="middle" font-size="24" fill="#112233">LA{child}BEL</text>
            <rect x="200" y="100" width="20" height="10" fill="lime"/>
            """);

        static string PathText(string attributes) => $"""<textPath href="#track" {attributes}>PATH</textPath>""";

        AssertSvgPaintInBothFormats(Picture(),
            [Picture(PathText("display=\"none\"")), Picture(PathText("display=\"inline\" style=\"display: none\""))],
            pdf => TestAssert.Equal(5, LabelPaths(pdf)), "0 1 0 rg");

        // These children still require path layout, even when they have no ink.
        foreach (string attributes in new[] { "display=\"none\" style=\"display: inline\"", "visibility=\"hidden\"" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            string paint = Paint(Picture(PathText(attributes)), diagnostics).ToString();
            TestAssert.DoesNotContain(LabelColor, paint);
            TestAssert.Contains("0 1 0 rg", paint);
            TestAssert.True(diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Fallback == "Partial" &&
                d.Message.Contains("unsupported text element", StringComparison.Ordinal)), "Path layout retains its diagnostic.");
        }
    }

    public static void SvgHiddenGraphicsMatchRemovedPayloadInBothDocumentContainers()
    {
        foreach (bool docx in new[] { false, true })
        foreach (bool review in docx ? new[] { false, true } : new[] { false })
        {
            PdfPage Render(string body)
            {
                string input = Input(docx, Svg(body));
                try
                {
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    using var package = OoxPackage.Open(input, default);
                    var resolver = new PortableRenderingTests.PinnedResolver();
                    var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
                    var page = docx
                        ? new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                            .RenderPagesAsync(DocxTests.ReadDocx(input, mode), diagnostics.Add, default).ToBlockingEnumerable().Single()
                        : new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default).ToBlockingEnumerable().Single();
                    TestAssert.Equal(0, diagnostics.Count);
                    return page;
                }
                finally { File.Delete(input); }
            }
            const string kept = "<rect x=\"200\" y=\"20\" width=\"80\" height=\"60\" fill=\"blue\"/>";
            var ordinary = Render(kept);
            var hidden = Render(kept + "<g style=\"display:none\"><rect width=\"400\" height=\"160\" fill=\"red\"/></g>");
            TestAssert.Equal(ordinary.Content, hidden.Content);
            TestAssert.Contains("0 0 1 rg", hidden.Content);
            PdfContentValidator.ValidatePage(hidden, 0, default);
        }
    }
}
